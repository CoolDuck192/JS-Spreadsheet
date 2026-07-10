import type {
  CellRange,
  NamedRange,
  SheetModel,
  SheetProtection,
  StructuredTable,
  StructuredTableColumn,
  TableAggregate,
  TableStyle,
  WorkbookModel
} from "../../types";
import type { TableDataType } from "../../table/core/types";
import {
  deserializeQueryRequest,
  serializeQueryRequest,
  type FilterExpression,
  type QueryRequest,
  type TableSort
} from "../../table/core/query";
import { formatCellAddress } from "../../lib/addressing";
import { normalizeExcelTableNameKey, validateExcelTableName } from "./tableNames";

const HEX_COLOR_PATTERN = /^#[0-9a-f]{6}$/i;
const DATA_TYPES = new Set<TableDataType>(["text", "number", "boolean", "date", "datetime", "custom"]);
const TOTALS_FUNCTIONS = new Set<TableAggregate>([
  "none",
  "sum",
  "average",
  "count",
  "countNumbers",
  "min",
  "max",
  "standardDeviation",
  "variance"
]);

export function migrateWorkbookModel(value: unknown): WorkbookModel | null {
  if (!isRecord(value) || (value.version !== 1 && value.version !== 2)) return null;
  if (typeof value.activeSheetId !== "string" || value.activeSheetId.length === 0) return null;
  if (!Array.isArray(value.sheets) || value.sheets.length === 0) return null;

  const sheets: SheetModel[] = [];
  const sheetIds = new Set<string>();
  for (const candidate of value.sheets) {
    const sheet = migrateSheet(candidate);
    if (!sheet || sheetIds.has(sheet.id)) return null;
    sheetIds.add(sheet.id);
    sheets.push(sheet);
  }

  const visibleSheets = sheets.some((sheet) => sheet.isHidden !== true)
    ? sheets
    : sheets.map((sheet, index) => index === 0 ? { ...sheet, isHidden: false } : sheet);
  const activeSheet = visibleSheets.find(
    (sheet) => sheet.id === value.activeSheetId && sheet.isHidden !== true
  ) ?? visibleSheets.find((sheet) => sheet.isHidden !== true) ?? visibleSheets[0];

  const namedRanges = migrateNamedRanges(value.namedRanges, visibleSheets);
  if (!namedRanges) return null;
  const tables = value.version === 1 ? [] : migrateTables(value.tables, visibleSheets);
  if (!tables) return null;

  return {
    version: 2,
    activeSheetId: activeSheet.id,
    sheets: visibleSheets,
    namedRanges,
    tables
  };
}

function migrateSheet(value: unknown): SheetModel | null {
  if (!isRecord(value)) return null;
  if (typeof value.id !== "string" || value.id.length === 0 || typeof value.name !== "string") return null;
  if (!positiveInteger(value.rowCount) || !positiveInteger(value.columnCount) || !isRecord(value.cells)) return null;
  if (!cellRecord(value.cells)) return null;
  if (!optionalRecord(value.formats) || !optionalRecord(value.columnWidths) || !optionalRecord(value.rowHeights)) return null;
  if (!optionalRecord(value.hiddenColumns) || !optionalRecord(value.hiddenRows)) return null;
  if (!optionalBoolean(value.isHidden) || !optionalString(value.tabColor)) return null;
  if (!optionalBoolean(value.freezeTopRow) || !optionalBoolean(value.freezeFirstColumn)) return null;
  if (!optionalRecord(value.comments) || !optionalRecord(value.hyperlinks) || !optionalRecord(value.validations)) return null;
  if (!optionalArray(value.conditionalFormats) || !optionalArray(value.filters) || !optionalArray(value.charts)) return null;
  if (!optionalArray(value.merges) || !optionalRecord(value.protection)) return null;
  if (value.autoFilterRange !== undefined && !isCellRange(value.autoFilterRange)) return null;

  const filters = (value.filters ?? []) as SheetModel["filters"];
  const autoFilterRange = isCellRange(value.autoFilterRange)
    ? cloneRange(value.autoFilterRange)
    : migratedAutoFilterRange(filters);
  return {
    id: value.id,
    name: value.name,
    rowCount: value.rowCount,
    columnCount: value.columnCount,
    isHidden: value.isHidden === true,
    tabColor: normalizeHexColor(value.tabColor as string | undefined),
    cells: value.cells as SheetModel["cells"],
    formats: (value.formats ?? {}) as SheetModel["formats"],
    columnWidths: (value.columnWidths ?? {}) as SheetModel["columnWidths"],
    rowHeights: (value.rowHeights ?? {}) as SheetModel["rowHeights"],
    hiddenColumns: isRecord(value.hiddenColumns) ? booleanFlagRecord(value.hiddenColumns) : {},
    hiddenRows: isRecord(value.hiddenRows) ? booleanFlagRecord(value.hiddenRows) : {},
    freezeTopRow: value.freezeTopRow === true,
    freezeFirstColumn: value.freezeFirstColumn === true,
    comments: (value.comments ?? {}) as SheetModel["comments"],
    hyperlinks: (value.hyperlinks ?? {}) as SheetModel["hyperlinks"],
    validations: (value.validations ?? {}) as SheetModel["validations"],
    conditionalFormats: (value.conditionalFormats ?? []) as SheetModel["conditionalFormats"],
    ...(autoFilterRange ? { autoFilterRange } : {}),
    filters,
    charts: (value.charts ?? []) as SheetModel["charts"],
    merges: (value.merges ?? []) as SheetModel["merges"],
    protection: migrateProtection(value.protection)
  };
}

function migrateNamedRanges(value: unknown, sheets: readonly SheetModel[]): NamedRange[] | null {
  if (value === undefined) return [];
  if (!Array.isArray(value)) return null;
  const sheetById = new Map(sheets.map((sheet) => [sheet.id, sheet]));
  const names = new Set<string>();
  const result: NamedRange[] = [];
  for (const item of value) {
    if (!isRecord(item) || typeof item.name !== "string" || item.name.trim().length === 0) return null;
    if (typeof item.sheetId !== "string" || !isCellRange(item.range)) return null;
    const sheet = sheetById.get(item.sheetId);
    if (!sheet || !rangeInBounds(item.range, sheet)) return null;
    const key = item.name.normalize("NFKC").toLowerCase();
    if (names.has(key)) return null;
    names.add(key);
    result.push({ name: item.name, sheetId: item.sheetId, range: cloneRange(item.range) });
  }
  return result;
}

function migrateTables(value: unknown, sheets: readonly SheetModel[]): StructuredTable[] | null {
  if (!Array.isArray(value)) return null;
  const sheetById = new Map(sheets.map((sheet) => [sheet.id, sheet]));
  const tableIds = new Set<string>();
  const tableNames = new Set<string>();
  const tables: StructuredTable[] = [];

  for (const candidate of value) {
    const table = migrateTable(candidate, sheetById);
    if (!table || tableIds.has(table.id)) return null;
    const nameKey = normalizeExcelTableNameKey(table.name);
    if (tableNames.has(nameKey)) return null;
    tableIds.add(table.id);
    tableNames.add(nameKey);
    tables.push(table);
  }

  for (let index = 0; index < tables.length; index += 1) {
    const table = tables[index];
    const sheet = sheetById.get(table.sheetId)!;
    if (sheet.merges.some((merge) => isRecord(merge) && isCellRange(merge.range) && rangesIntersect(table.range, merge.range))) {
      return null;
    }
    if (tables.slice(0, index).some(
      (other) => other.sheetId === table.sheetId && rangesIntersect(other.range, table.range)
    )) {
      return null;
    }
  }
  return tables;
}

function migrateTable(
  value: unknown,
  sheetById: ReadonlyMap<string, SheetModel>
): StructuredTable | null {
  if (!isRecord(value)) return null;
  if (!nonBlankString(value.id) || !nonBlankString(value.name) || !nonBlankString(value.sheetId)) return null;
  if (!validateExcelTableName(value.name).valid || !isCellRange(value.range)) return null;
  if (typeof value.headerRow !== "boolean" || typeof value.totalsRow !== "boolean") return null;
  const sheet = sheetById.get(value.sheetId);
  if (!sheet || !rangeInBounds(value.range, sheet)) return null;
  const height = value.range.end.row - value.range.start.row + 1;
  const width = value.range.end.column - value.range.start.column + 1;
  if (height < Number(value.headerRow) + Number(value.totalsRow) || !Array.isArray(value.columns)) return null;
  if (value.columns.length !== width || !Array.isArray(value.rowIds)) return null;

  const columns: StructuredTableColumn[] = [];
  const columnIds = new Set<string>();
  const columnNames = new Set<string>();
  for (let index = 0; index < value.columns.length; index += 1) {
    const column = migrateColumn(value.columns[index], value.range.start.column + index, value.totalsRow);
    if (!column || columnIds.has(column.id)) return null;
    const nameKey = column.name.normalize("NFKC").toLowerCase();
    if (columnNames.has(nameKey)) return null;
    columnIds.add(column.id);
    columnNames.add(nameKey);
    columns.push(column);
    if (value.headerRow) {
      const headerAddress = formatCellAddress({ row: value.range.start.row, column: column.sheetColumn });
      if (sheet.cells[headerAddress] !== column.name) return null;
    }
    if (column.totalsLabel !== undefined) {
      const totalsAddress = formatCellAddress({ row: value.range.end.row, column: column.sheetColumn });
      if (sheet.cells[totalsAddress] !== column.totalsLabel) return null;
    }
  }

  const expectedRows = height - Number(value.headerRow) - Number(value.totalsRow);
  if (value.rowIds.length !== expectedRows || !value.rowIds.every(nonBlankString)) return null;
  if (new Set(value.rowIds).size !== value.rowIds.length) return null;
  if (value.keyColumnId !== undefined && (!nonBlankString(value.keyColumnId) || !columnIds.has(value.keyColumnId))) return null;

  const sort = migrateSort(value.sort, columnIds);
  if (sort === null) return null;
  const filter = migrateFilter(value.filter, columnIds);
  if (filter === null) return null;
  const style = migrateStyle(value.style);
  if (style === null) return null;

  return {
    id: value.id,
    name: value.name,
    sheetId: value.sheetId,
    range: cloneRange(value.range),
    headerRow: value.headerRow,
    totalsRow: value.totalsRow,
    columns,
    rowIds: [...value.rowIds],
    ...(value.keyColumnId === undefined ? {} : { keyColumnId: value.keyColumnId }),
    ...(style === undefined ? {} : { style }),
    ...(sort === undefined ? {} : { sort }),
    ...(filter === undefined ? {} : { filter })
  };
}

function migrateColumn(value: unknown, expectedSheetColumn: number, totalsRow: boolean): StructuredTableColumn | null {
  if (!isRecord(value) || !nonBlankString(value.id) || !nonBlankString(value.name)) return null;
  if (value.sheetColumn !== expectedSheetColumn) return null;
  if (value.dataType !== undefined && (typeof value.dataType !== "string" || !DATA_TYPES.has(value.dataType as TableDataType))) return null;
  if (value.calculatedFormula !== undefined && typeof value.calculatedFormula !== "string") return null;
  if (value.totalsFunction !== undefined && (
    typeof value.totalsFunction !== "string" || !TOTALS_FUNCTIONS.has(value.totalsFunction as TableAggregate)
  )) return null;
  if (value.totalsLabel !== undefined && typeof value.totalsLabel !== "string") return null;
  if (value.totalsFunction !== undefined && value.totalsLabel !== undefined) return null;
  if (!totalsRow && (value.totalsFunction !== undefined || value.totalsLabel !== undefined)) return null;
  return {
    id: value.id,
    name: value.name,
    sheetColumn: value.sheetColumn,
    ...(value.dataType === undefined ? {} : { dataType: value.dataType as TableDataType }),
    ...(value.calculatedFormula === undefined ? {} : { calculatedFormula: value.calculatedFormula }),
    ...(value.totalsFunction === undefined ? {} : { totalsFunction: value.totalsFunction as TableAggregate }),
    ...(value.totalsLabel === undefined ? {} : { totalsLabel: value.totalsLabel })
  };
}

function migrateSort(value: unknown, columnIds: ReadonlySet<string>): readonly TableSort[] | undefined | null {
  if (value === undefined) return undefined;
  if (!Array.isArray(value)) return null;
  const result: TableSort[] = [];
  for (const candidate of value) {
    if (!isRecord(candidate) || !nonBlankString(candidate.columnId) || !columnIds.has(candidate.columnId)) return null;
    if (candidate.direction !== "asc" && candidate.direction !== "desc") return null;
    if (candidate.nulls !== undefined && candidate.nulls !== "first" && candidate.nulls !== "last") return null;
    result.push({
      columnId: candidate.columnId,
      direction: candidate.direction,
      ...(candidate.nulls === undefined ? {} : { nulls: candidate.nulls })
    });
  }
  return result;
}

function migrateFilter(value: unknown, columnIds: ReadonlySet<string>): FilterExpression | undefined | null {
  if (value === undefined) return undefined;
  let filter: FilterExpression;
  try {
    const request = deserializeQueryRequest(serializeQueryRequest({
      sorting: [],
      filter: value as FilterExpression,
      grouping: [],
      aggregates: [],
      pagination: { kind: "none" }
    } satisfies QueryRequest));
    filter = request.filter!;
  } catch {
    return null;
  }
  return everyFilterColumn(filter, columnIds) ? filter : null;
}

function everyFilterColumn(filter: FilterExpression, columnIds: ReadonlySet<string>): boolean {
  if (filter.kind === "logical") return filter.operands.every((operand) => everyFilterColumn(operand, columnIds));
  if (filter.kind === "not") return everyFilterColumn(filter.operand, columnIds);
  return columnIds.has(filter.columnId);
}

function migrateStyle(value: unknown): TableStyle | undefined | null {
  if (value === undefined) return undefined;
  if (!isRecord(value) || !nonBlankString(value.theme)) return null;
  for (const key of ["showFirstColumn", "showLastColumn", "showRowStripes", "showColumnStripes"] as const) {
    if (!optionalBoolean(value[key])) return null;
  }
  return {
    theme: value.theme,
    ...(value.showFirstColumn === undefined ? {} : { showFirstColumn: value.showFirstColumn as boolean }),
    ...(value.showLastColumn === undefined ? {} : { showLastColumn: value.showLastColumn as boolean }),
    ...(value.showRowStripes === undefined ? {} : { showRowStripes: value.showRowStripes as boolean }),
    ...(value.showColumnStripes === undefined ? {} : { showColumnStripes: value.showColumnStripes as boolean })
  };
}

function migratedAutoFilterRange(filters: SheetModel["filters"]): CellRange | undefined {
  const withRange = filters.find((filter) => isRecord(filter) && isCellRange(filter.range));
  return withRange && isRecord(withRange) && isCellRange(withRange.range) ? cloneRange(withRange.range) : undefined;
}

function migrateProtection(value: unknown): SheetProtection {
  if (!isRecord(value)) return { isProtected: false, lockedCells: {}, unlockedCells: {} };
  return {
    isProtected: value.isProtected === true,
    lockedCells: isRecord(value.lockedCells) ? booleanFlagRecord(value.lockedCells) : {},
    unlockedCells: isRecord(value.unlockedCells) ? booleanFlagRecord(value.unlockedCells) : {}
  };
}

function rangeInBounds(range: CellRange, sheet: SheetModel): boolean {
  return range.start.row >= 0
    && range.start.column >= 0
    && range.end.row >= range.start.row
    && range.end.column >= range.start.column
    && range.end.row < sheet.rowCount
    && range.end.column < sheet.columnCount;
}

function rangesIntersect(left: CellRange, right: CellRange): boolean {
  return left.start.row <= right.end.row
    && left.end.row >= right.start.row
    && left.start.column <= right.end.column
    && left.end.column >= right.start.column;
}

function isCellRange(value: unknown): value is CellRange {
  return isRecord(value) && isCellCoord(value.start) && isCellCoord(value.end);
}

function isCellCoord(value: unknown): value is CellRange["start"] {
  return isRecord(value) && Number.isInteger(value.row) && Number.isInteger(value.column);
}

function cloneRange(range: CellRange): CellRange {
  return { start: { ...range.start }, end: { ...range.end } };
}

function cellRecord(value: Record<string, unknown>): boolean {
  return Object.values(value).every((cell) => cell === null
    || typeof cell === "string"
    || typeof cell === "boolean"
    || (typeof cell === "number" && Number.isFinite(cell)));
}

function booleanFlagRecord(value: Record<string, unknown>): Record<string, boolean> {
  return Object.fromEntries(Object.entries(value).filter(([, flag]) => flag === true)) as Record<string, boolean>;
}

function normalizeHexColor(color: string | undefined): string | undefined {
  const value = color?.trim();
  if (!value) return undefined;
  const normalized = value.startsWith("#") ? value : `#${value}`;
  return HEX_COLOR_PATTERN.test(normalized) ? normalized.toLowerCase() : undefined;
}

function positiveInteger(value: unknown): value is number {
  return Number.isSafeInteger(value) && (value as number) > 0;
}

function nonBlankString(value: unknown): value is string {
  return typeof value === "string" && value.trim().length > 0;
}

function optionalBoolean(value: unknown): boolean {
  return value === undefined || typeof value === "boolean";
}

function optionalString(value: unknown): boolean {
  return value === undefined || typeof value === "string";
}

function optionalRecord(value: unknown): boolean {
  return value === undefined || isRecord(value);
}

function optionalArray(value: unknown): boolean {
  return value === undefined || Array.isArray(value);
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}
