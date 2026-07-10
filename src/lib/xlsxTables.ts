import type ExcelJS from "exceljs";
import { createRandomId, createStableKeyRowId, type IdGenerator } from "../core/ids";
import { normalizeExcelTableNameKey, validateExcelTableName } from "../core/workbook/tableNames";
import type { FilterExpression, QueryScalar } from "../table/core/query";
import type { TableDataType } from "../table/core/types";
import type {
  CellContent,
  SheetModel,
  StructuredTable,
  StructuredTableColumn,
  TableAggregate,
  WorkbookModel
} from "../types";
import { formatCellAddress, parseRangeAddress } from "./addressing";
import { structuredFormulaToA1 } from "./structuredFormula";
import { nativeTotalsFunctionToAggregate, type NativeTableXmlMetadata } from "./xlsxTableXml";

export type XlsxImportOptions = {
  idGenerator?: IdGenerator;
  tableKeys?: Readonly<Record<string, { columnName: string }>>;
};

type NativeColumnModel = {
  name?: string;
  totalsRowLabel?: string;
  totalsRowFunction?: string;
  filterButton?: boolean;
};

type NativeTableModel = {
  name?: string;
  displayName?: string;
  tableRef?: string;
  headerRow?: boolean;
  totalsRow?: boolean;
  columns?: NativeColumnModel[];
  style?: {
    theme?: string;
    showFirstColumn?: boolean;
    showLastColumn?: boolean;
    showRowStripes?: boolean;
    showColumnStripes?: boolean;
  };
};

export async function importStructuredTablesFromWorksheet(
  worksheet: ExcelJS.Worksheet,
  sheetId: string,
  xmlMetadata: readonly NativeTableXmlMetadata[],
  options: XlsxImportOptions = {}
): Promise<StructuredTable[]> {
  const nativeTables = listWorksheetTables(worksheet);
  const names = new Set<string>();
  const plans = nativeTables.map((nativeTable) => {
    const model = nativeTableModel(nativeTable);
    const name = model.name || model.displayName || nativeTable.name;
    const validation = validateExcelTableName(name);
    if (!validation.valid) throw xlsxTableError("XLSX_TABLE_NAME_INVALID", `Invalid native table name ${name}`);
    if (names.has(validation.normalizedKey)) {
      throw xlsxTableError("XLSX_TABLE_NAME_DUPLICATE", `Duplicate native table name ${name}`);
    }
    names.add(validation.normalizedKey);
    if (!model.tableRef) throw xlsxTableError("XLSX_TABLE_RANGE_INVALID", `Native table ${name} has no range`);
    let range;
    try {
      range = parseRangeAddress(model.tableRef.replace(/\$/g, ""));
    } catch {
      throw xlsxTableError("XLSX_TABLE_RANGE_INVALID", `Native table ${name} has an invalid range`);
    }
    const nativeColumns = Array.isArray(model.columns) ? model.columns : [];
    if (nativeColumns.length === 0 || nativeColumns.length !== range.end.column - range.start.column + 1) {
      throw xlsxTableError("XLSX_TABLE_RANGE_INVALID", `Native table ${name} has inconsistent columns`);
    }
    if (nativeColumns.some((column) => typeof column.name !== "string" || column.name.length === 0)) {
      throw xlsxTableError("XLSX_TABLE_COLUMN_INVALID", `Native table ${name} has an invalid column name`);
    }
    const headerRow = model.headerRow === true || firstRowMatchesColumnNames(worksheet, range.start.row, range.start.column, nativeColumns);
    const totalsRow = model.totalsRow === true;
    const bodyStart = range.start.row + Number(headerRow);
    const bodyEnd = range.end.row - Number(totalsRow);
    if (bodyEnd < bodyStart - 1) {
      throw xlsxTableError("XLSX_TABLE_RANGE_INVALID", `Native table ${name} has invalid row boundaries`);
    }
    return { nativeTable, model, name, range, nativeColumns, headerRow, totalsRow, bodyStart, bodyEnd };
  });

  const createId = options.idGenerator ?? createRandomId;
  const imported: StructuredTable[] = [];
  for (const plan of plans) {
    const metadata = xmlMetadata.find((item) => normalizeName(item.name) === normalizeName(plan.name));
    const tableId = createId("table");
    const columns: StructuredTableColumn[] = plan.nativeColumns.map((nativeColumn, index) => {
      const name = nativeColumn.name!;
      const sheetColumn = plan.range.start.column + index;
      const dataType = inferColumnDataType(worksheet, sheetColumn, plan.bodyStart, plan.bodyEnd);
      const total = metadata?.totals[name];
      const nativeFunction = total?.function ?? nativeColumn.totalsRowFunction;
      const aggregate = nativeFunction ? nativeTotalsFunctionToAggregate(nativeFunction) : undefined;
      const totalsLabel = total?.label ?? nativeColumn.totalsRowLabel;
      return {
        id: createId("table-column"),
        name,
        sheetColumn,
        ...(dataType ? { dataType } : {}),
        ...(aggregate && aggregate !== "none" ? { totalsFunction: aggregate } : {}),
        ...(totalsLabel !== undefined ? { totalsLabel } : {})
      };
    });

    const provisional: StructuredTable = {
      id: tableId,
      name: plan.name,
      sheetId,
      range: plan.range,
      headerRow: plan.headerRow,
      totalsRow: plan.totalsRow,
      columns,
      rowIds: [],
      ...(plan.model.style?.theme ? {
        style: {
          theme: plan.model.style.theme,
          showFirstColumn: plan.model.style.showFirstColumn ?? false,
          showLastColumn: plan.model.style.showLastColumn ?? false,
          showRowStripes: plan.model.style.showRowStripes ?? false,
          showColumnStripes: plan.model.style.showColumnStripes ?? false
        }
      } : {})
    };

    const columnsWithCalculatedFormulas = columns.map((column) => {
      const xmlFormula = metadata?.calculatedColumns[column.name];
      if (!xmlFormula || plan.bodyStart > plan.bodyEnd) return column;
      const translated = structuredFormulaToA1(xmlFormula, provisional, plan.bodyStart);
      if (!translated.ok) throw issueError(translated.issue);
      const expanded = excelCellFormula(worksheet.getCell(plan.bodyStart + 1, column.sheetColumn + 1));
      if (expanded && normalizeFormula(expanded) !== normalizeFormula(translated.formula)) {
        throw xlsxTableError(
          "XLSX_TABLE_FORMULA_MISMATCH",
          `Expanded formula for ${plan.name}[${column.name}] does not match table metadata`
        );
      }
      return { ...column, calculatedFormula: expanded ?? translated.formula };
    });
    const tableWithColumns = { ...provisional, columns: columnsWithCalculatedFormulas };

    applyCustomTotalsFormulas(worksheet, tableWithColumns, metadata);
    const keyConfig = findTableKeyConfig(options.tableKeys, plan.name);
    const keyColumn = keyConfig
      ? columnsWithCalculatedFormulas.find((column) => normalizeName(column.name) === normalizeName(keyConfig.columnName))
      : undefined;
    if (keyConfig && !keyColumn) {
      throw xlsxTableError("XLSX_TABLE_KEY_INVALID", `Configured key column does not exist in ${plan.name}`);
    }
    const rowIds = await createImportedRowIds(
      worksheet,
      plan.name,
      plan.bodyStart,
      plan.bodyEnd,
      keyColumn,
      createId
    );
    const filter = metadata?.filter
      ? remapFilterExpression(metadata.filter, columnsWithCalculatedFormulas)
      : undefined;
    imported.push({
      ...tableWithColumns,
      rowIds,
      ...(keyColumn ? { keyColumnId: keyColumn.id } : {}),
      ...(filter ? { filter } : {})
    });
  }
  return imported;
}

export function addStructuredTablesToWorksheet(
  workbook: WorkbookModel,
  sheet: SheetModel,
  worksheet: ExcelJS.Worksheet
): void {
  validateWorkbookTableNames(workbook.tables);
  for (const table of workbook.tables.filter((candidate) => candidate.sheetId === sheet.id)) {
    validateTableProjection(table, sheet);
    const bodyStart = table.range.start.row + Number(table.headerRow);
    const bodyEnd = table.range.end.row - Number(table.totalsRow);
    const rows: ExcelJS.CellValue[][] = [];
    for (let row = bodyStart; row <= bodyEnd; row += 1) {
      rows.push(table.columns.map((column) => worksheet.getCell(row + 1, column.sheetColumn + 1).value));
    }

    const snapshot = snapshotTableValues(worksheet, table);
    worksheet.addTable({
      name: table.name,
      ref: formatCellAddress(table.range.start),
      headerRow: table.headerRow,
      totalsRow: table.totalsRow,
      columns: table.columns.map((column) => ({
        name: column.name,
        ...(column.totalsLabel !== undefined ? { totalsRowLabel: column.totalsLabel } : {}),
        ...(column.totalsFunction && column.totalsFunction !== "none"
          ? { totalsRowFunction: aggregateToExcelJs(column.totalsFunction) }
          : {}),
        filterButton: true
      })) as ExcelJS.TableColumnProperties[],
      rows,
      style: table.style ? ({
        theme: table.style.theme,
        showFirstColumn: table.style.showFirstColumn ?? false,
        showLastColumn: table.style.showLastColumn ?? false,
        showRowStripes: table.style.showRowStripes ?? false,
        showColumnStripes: table.style.showColumnStripes ?? false
      } as ExcelJS.TableStyleProperties) : undefined
    });
    restoreTableValues(worksheet, snapshot);
  }
}

export function listWorksheetTables(worksheet: ExcelJS.Worksheet): ExcelJS.Table[] {
  const value: unknown = worksheet.getTables();
  if (!Array.isArray(value) || value.some((item) => (
    typeof item !== "object" || item === null || typeof (item as { name?: unknown }).name !== "string"
  ))) {
    throw xlsxTableError("XLSX_TABLE_MODEL_INVALID", "Invalid native table list returned by ExcelJS");
  }
  return value as ExcelJS.Table[];
}

function nativeTableModel(table: ExcelJS.Table): NativeTableModel {
  const model = (table as unknown as { model?: unknown; table?: unknown }).model
    ?? (table as unknown as { table?: unknown }).table;
  if (typeof model !== "object" || model === null) {
    throw xlsxTableError("XLSX_TABLE_MODEL_INVALID", `Invalid native table model for ${table.name}`);
  }
  return model as NativeTableModel;
}

function firstRowMatchesColumnNames(
  worksheet: ExcelJS.Worksheet,
  row: number,
  startColumn: number,
  columns: readonly NativeColumnModel[]
): boolean {
  return columns.every((column, index) => (
    String(excelCellValueToContent(worksheet.getCell(row + 1, startColumn + index + 1).value) ?? "") === column.name
  ));
}

function inferColumnDataType(
  worksheet: ExcelJS.Worksheet,
  column: number,
  bodyStart: number,
  bodyEnd: number
): TableDataType | undefined {
  for (let row = bodyStart; row <= bodyEnd; row += 1) {
    const cell = worksheet.getCell(row + 1, column + 1);
    const value = formulaResultOrValue(cell.value);
    if (value === null || value === undefined || value === "") continue;
    if (value instanceof Date) return excelNumberFormatHasTime(cell.numFmt) ? "datetime" : "date";
    if (typeof value === "number") return "number";
    if (typeof value === "boolean") return "boolean";
    if (typeof value === "string") return "text";
    return "custom";
  }
  return undefined;
}

function excelNumberFormatHasTime(format: string | undefined): boolean {
  if (!format) return false;
  const withoutLiterals = format.replace(/"[^"]*"/g, "").replace(/\\./g, "");
  return /(?:h+|s+|AM\/PM)/i.test(withoutLiterals);
}

function applyCustomTotalsFormulas(
  worksheet: ExcelJS.Worksheet,
  table: StructuredTable,
  metadata: NativeTableXmlMetadata | undefined
): void {
  if (!table.totalsRow || !metadata) return;
  for (const column of table.columns) {
    const formula = metadata.totals[column.name]?.formula;
    if (!formula) continue;
    const translated = structuredFormulaToA1(formula, table, firstBodyRow(table));
    if (!translated.ok) throw issueError(translated.issue);
    worksheet.getCell(table.range.end.row + 1, column.sheetColumn + 1).value = {
      formula: translated.formula.slice(1)
    };
  }
}

async function createImportedRowIds(
  worksheet: ExcelJS.Worksheet,
  tableName: string,
  bodyStart: number,
  bodyEnd: number,
  keyColumn: StructuredTableColumn | undefined,
  createId: IdGenerator
): Promise<string[]> {
  if (!keyColumn) {
    return Array.from({ length: Math.max(0, bodyEnd - bodyStart + 1) }, () => createId("table-row"));
  }
  const values: CellContent[] = [];
  const identities = new Set<string>();
  for (let row = bodyStart; row <= bodyEnd; row += 1) {
    const value = excelCellValueToContent(worksheet.getCell(row + 1, keyColumn.sheetColumn + 1).value);
    if (value === null || (typeof value === "string" && value.trim().length === 0)) {
      throw xlsxTableError("XLSX_TABLE_KEY_INVALID", `Configured key column ${keyColumn.name} contains a blank value`);
    }
    const identity = typedIdentity(value);
    if (identities.has(identity)) {
      throw xlsxTableError("XLSX_TABLE_KEY_INVALID", `Configured key column ${keyColumn.name} contains duplicate values`);
    }
    identities.add(identity);
    values.push(value);
  }
  return Promise.all(values.map((value) => createStableKeyRowId(tableName, keyColumn.name, value)));
}

function findTableKeyConfig(
  tableKeys: XlsxImportOptions["tableKeys"],
  tableName: string
): { columnName: string } | undefined {
  if (!tableKeys) return undefined;
  const entry = Object.entries(tableKeys).find(([name]) => normalizeName(name) === normalizeName(tableName));
  return entry?.[1];
}

function remapFilterExpression(
  filter: FilterExpression,
  columns: readonly StructuredTableColumn[]
): FilterExpression {
  if (filter.kind === "logical") {
    return { ...filter, operands: filter.operands.map((operand) => remapFilterExpression(operand, columns)) };
  }
  if (filter.kind === "not") return { ...filter, operand: remapFilterExpression(filter.operand, columns) };
  const column = columns.find((candidate) => normalizeName(candidate.name) === normalizeName(filter.columnId));
  if (!column) throw xlsxTableError("XLSX_TABLE_FILTER_INVALID", "Native table filter targets an unknown column");
  if (filter.kind === "set") {
    return { ...filter, columnId: column.id, values: filter.values.map((value) => coerceFilterScalar(value, column.dataType)) };
  }
  if (filter.kind === "comparison") {
    return { ...filter, columnId: column.id, value: coerceFilterScalar(filter.value, column.dataType) };
  }
  if (filter.kind === "range") {
    return {
      ...filter,
      columnId: column.id,
      lower: coerceFilterScalar(filter.lower, column.dataType),
      upper: coerceFilterScalar(filter.upper, column.dataType)
    };
  }
  return { ...filter, columnId: column.id };
}

function coerceFilterScalar(value: QueryScalar, dataType: TableDataType | undefined): QueryScalar {
  if (value.type !== "string") return value;
  if (dataType === "number" && value.value.trim() !== "" && Number.isFinite(Number(value.value))) {
    return { type: "number", value: Number(value.value) };
  }
  if (dataType === "boolean" && /^(?:true|false|1|0)$/i.test(value.value)) {
    return { type: "boolean", value: /^(?:true|1)$/i.test(value.value) };
  }
  return value;
}

function validateWorkbookTableNames(tables: readonly StructuredTable[]): void {
  const names = new Set<string>();
  for (const table of tables) {
    const validation = validateExcelTableName(table.name);
    if (!validation.valid) throw xlsxTableError("XLSX_TABLE_NAME_INVALID", `Invalid native table name ${table.name}`);
    if (names.has(validation.normalizedKey)) {
      throw xlsxTableError("XLSX_TABLE_NAME_DUPLICATE", `Duplicate native table name ${table.name}`);
    }
    names.add(validation.normalizedKey);
  }
}

function validateTableProjection(table: StructuredTable, sheet: SheetModel): void {
  if (table.columns.length !== table.range.end.column - table.range.start.column + 1) {
    throw xlsxTableError("XLSX_TABLE_RANGE_INVALID", `Table ${table.name} has inconsistent columns`);
  }
  const expectedRows = table.range.end.row - table.range.start.row + 1
    - Number(table.headerRow) - Number(table.totalsRow);
  if (expectedRows !== table.rowIds.length || expectedRows < 0) {
    throw xlsxTableError("XLSX_TABLE_RANGE_INVALID", `Table ${table.name} has inconsistent rows`);
  }
  if (table.range.end.row >= sheet.rowCount || table.range.end.column >= sheet.columnCount) {
    throw xlsxTableError("XLSX_TABLE_RANGE_INVALID", `Table ${table.name} exceeds its worksheet bounds`);
  }
}

function snapshotTableValues(
  worksheet: ExcelJS.Worksheet,
  table: StructuredTable
): Array<{ row: number; column: number; value: ExcelJS.CellValue }> {
  const values: Array<{ row: number; column: number; value: ExcelJS.CellValue }> = [];
  for (let row = table.range.start.row; row <= table.range.end.row; row += 1) {
    for (let column = table.range.start.column; column <= table.range.end.column; column += 1) {
      values.push({ row, column, value: worksheet.getCell(row + 1, column + 1).value });
    }
  }
  return values;
}

function restoreTableValues(
  worksheet: ExcelJS.Worksheet,
  values: readonly { row: number; column: number; value: ExcelJS.CellValue }[]
): void {
  for (const item of values) worksheet.getCell(item.row + 1, item.column + 1).value = item.value;
}

function aggregateToExcelJs(
  aggregate: Exclude<TableAggregate, "none">
): NonNullable<ExcelJS.TableColumnProperties["totalsRowFunction"]> {
  return ({
    sum: "sum",
    average: "average",
    count: "count",
    countNumbers: "countNums",
    min: "min",
    max: "max",
    standardDeviation: "stdDev",
    variance: "var"
  } as const)[aggregate];
}

function excelCellFormula(cell: ExcelJS.Cell): string | undefined {
  const formula = cell.formula;
  return typeof formula === "string" && formula.length > 0 ? `=${formula}` : undefined;
}

function normalizeFormula(formula: string): string {
  return formula.replace(/\$/g, "").replace(/\s+/g, "").toUpperCase();
}

function formulaResultOrValue(value: ExcelJS.CellValue): unknown {
  if (isRecord(value) && "result" in value) return value.result;
  return value;
}

function excelCellValueToContent(value: ExcelJS.CellValue): CellContent {
  if (value === null || value === undefined || value === "") return null;
  if (typeof value === "string" || typeof value === "number" || typeof value === "boolean") return value;
  if (value instanceof Date) return (value.getTime() - Date.UTC(1899, 11, 30)) / 86_400_000;
  if (isRecord(value)) {
    if (typeof value.formula === "string") return excelCellValueToContent(value.result as ExcelJS.CellValue);
    if (typeof value.sharedFormula === "string") return excelCellValueToContent(value.result as ExcelJS.CellValue);
    if (typeof value.text === "string") return value.text;
    if (Array.isArray(value.richText)) {
      return value.richText.map((part) => isRecord(part) ? String(part.text ?? "") : "").join("");
    }
    if (typeof value.error === "string") return value.error;
  }
  return String(value);
}

function typedIdentity(value: CellContent): string {
  return `${typeof value}:${String(value)}`;
}

function firstBodyRow(table: StructuredTable): number {
  return table.range.start.row + Number(table.headerRow);
}

function normalizeName(value: string): string {
  return normalizeExcelTableNameKey(value);
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

function issueError(issue: { code: string; message: string }): Error {
  return Object.assign(new Error(issue.message), { code: issue.code, issue });
}

function xlsxTableError(code: string, message: string): Error {
  return Object.assign(new Error(message), { code });
}
