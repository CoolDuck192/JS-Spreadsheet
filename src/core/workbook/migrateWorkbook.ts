import type {
  CellBorderSide,
  CellBorders,
  CellContent,
  CellFormat,
  CellRange,
  ConditionalFormatCondition,
  ConditionalFormatRule,
  DataValidationRule,
  NamedRange,
  SheetChart,
  SheetFilter,
  SheetMerge,
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
import { formatCellAddress, parseCellAddress } from "../../lib/addressing";
import {
  normalizeExcelTableNameKey,
  validateExcelTableNameForInterop
} from "./tableNames";
import { normalizeNamedRangeLookup } from "./namedRangeNames";

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
  const bounds = { rowCount: value.rowCount, columnCount: value.columnCount };
  if (value.autoFilterRange !== undefined && (
    !isCellRange(value.autoFilterRange) || !rangeInBounds(value.autoFilterRange, bounds)
  )) return null;

  const comments = migrateStringCellRecord(value.comments, bounds);
  const hyperlinks = migrateStringCellRecord(value.hyperlinks, bounds);
  const validations = migrateValidations(value.validations, bounds);
  const conditionalFormats = migrateConditionalFormats(value.conditionalFormats, bounds);
  const filters = migrateSheetFilters(value.filters, bounds);
  const charts = migrateSheetCharts(value.charts, bounds);
  const merges = migrateSheetMerges(value.merges, bounds);
  if (!comments || !hyperlinks || !validations || !conditionalFormats || !filters || !charts || !merges) return null;

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
    comments,
    hyperlinks,
    validations,
    conditionalFormats,
    ...(autoFilterRange ? { autoFilterRange } : {}),
    filters,
    charts,
    merges,
    protection: migrateProtection(value.protection)
  };
}

type SheetBounds = Pick<SheetModel, "rowCount" | "columnCount">;

function migrateStringCellRecord(
  value: unknown,
  bounds: SheetBounds
): Record<string, string> | null {
  if (value === undefined) return {};
  if (!isRecord(value)) return null;
  const result: Record<string, string> = {};
  for (const [address, item] of Object.entries(value)) {
    const canonicalAddress = canonicalCellAddressInBounds(address, bounds);
    if (typeof item !== "string" || !canonicalAddress || Object.hasOwn(result, canonicalAddress)) return null;
    result[canonicalAddress] = item;
  }
  return result;
}

function migrateValidations(
  value: unknown,
  bounds: SheetBounds
): Record<string, DataValidationRule> | null {
  if (value === undefined) return {};
  if (!isRecord(value)) return null;
  const result: Record<string, DataValidationRule> = {};
  for (const [address, item] of Object.entries(value)) {
    const rule = migrateValidationRule(item);
    const canonicalAddress = canonicalCellAddressInBounds(address, bounds);
    if (!rule || !canonicalAddress || Object.hasOwn(result, canonicalAddress)) return null;
    result[canonicalAddress] = rule;
  }
  return result;
}

export function migrateValidationRule(value: unknown): DataValidationRule | null {
  if (!isRecord(value) || !optionalBoolean(value.allowBlank)) return null;
  if (value.type === "list") {
    if (!Array.isArray(value.values) || !value.values.every((item) => typeof item === "string")) return null;
    return {
      type: "list",
      values: [...value.values],
      ...(value.allowBlank === undefined ? {} : { allowBlank: value.allowBlank })
    };
  }
  if (value.type !== "number" && value.type !== "textLength") return null;
  if (!optionalFiniteNumber(value.min) || !optionalFiniteNumber(value.max)) return null;
  if (value.min !== undefined && value.max !== undefined && value.min > value.max) return null;
  const bounds = {
    ...(value.min === undefined ? {} : { min: value.min }),
    ...(value.max === undefined ? {} : { max: value.max }),
    ...(value.allowBlank === undefined ? {} : { allowBlank: value.allowBlank })
  };
  return value.type === "number" ? { type: "number", ...bounds } : { type: "textLength", ...bounds };
}

function migrateConditionalFormats(
  value: unknown,
  bounds: SheetBounds
): ConditionalFormatRule[] | null {
  if (value === undefined) return [];
  if (!Array.isArray(value)) return null;
  const result: ConditionalFormatRule[] = [];
  const ids = new Set<string>();
  for (const item of value) {
    const rule = migrateConditionalFormatRule(item, bounds);
    if (!rule || ids.has(rule.id)) return null;
    ids.add(rule.id);
    result.push(rule);
  }
  return result;
}

export function migrateConditionalFormatRule(
  value: unknown,
  bounds: SheetBounds
): ConditionalFormatRule | null {
  if (!isRecord(value) || !nonBlankString(value.id)) return null;
  if (!isCellRange(value.range) || !rangeInBounds(value.range, bounds)) return null;
  const condition = migrateConditionalFormatCondition(value.condition);
  const format = migrateCellFormat(value.format);
  return condition && format
    ? { id: value.id, range: cloneRange(value.range), condition, format }
    : null;
}

function migrateConditionalFormatCondition(value: unknown): ConditionalFormatCondition | null {
  if (!isRecord(value) || typeof value.type !== "string") return null;
  if (["greaterThan", "lessThan", "equalTo", "textContains"].includes(value.type)) {
    return typeof value.value === "string"
      ? { type: value.type as "greaterThan" | "lessThan" | "equalTo" | "textContains", value: value.value }
      : null;
  }
  if (value.type === "between") {
    return typeof value.value === "string" && typeof value.secondValue === "string"
      ? { type: "between", value: value.value, secondValue: value.secondValue }
      : null;
  }
  if (["blank", "notBlank", "duplicate", "unique"].includes(value.type)) {
    return { type: value.type as "blank" | "notBlank" | "duplicate" | "unique" };
  }
  if (value.type === "top" || value.type === "bottom") {
    return positiveInteger(value.count) ? { type: value.type, count: value.count } : null;
  }
  if (value.type === "dataBar") {
    return typeof value.color === "string" ? { type: "dataBar", color: value.color } : null;
  }
  if (value.type === "colorScale") {
    return typeof value.minColor === "string" && typeof value.maxColor === "string"
      ? { type: "colorScale", minColor: value.minColor, maxColor: value.maxColor }
      : null;
  }
  return null;
}

function migrateCellFormat(value: unknown): CellFormat | null {
  if (!isRecord(value)) return null;
  for (const key of ["bold", "italic", "wrapText"] as const) {
    if (!optionalBoolean(value[key])) return null;
  }
  for (const key of ["fontFamily", "textColor", "backgroundColor"] as const) {
    if (!optionalString(value[key])) return null;
  }
  if (value.fontSize !== undefined && (typeof value.fontSize !== "number" || !Number.isFinite(value.fontSize) || value.fontSize <= 0)) return null;
  if (value.numberFormat !== undefined && !["general", "number", "currency", "percent", "date", "dateTime", "financial", "financial2", "accounting"].includes(value.numberFormat as string)) return null;
  if (value.horizontalAlign !== undefined && !["left", "center", "right"].includes(value.horizontalAlign as string)) return null;
  if (value.verticalAlign !== undefined && !["top", "middle", "bottom"].includes(value.verticalAlign as string)) return null;
  const borders = migrateCellBorders(value.borders);
  if (borders === null) return null;
  return {
    ...(value.bold === undefined ? {} : { bold: value.bold as boolean }),
    ...(value.italic === undefined ? {} : { italic: value.italic as boolean }),
    ...(value.fontFamily === undefined ? {} : { fontFamily: value.fontFamily as string }),
    ...(value.fontSize === undefined ? {} : { fontSize: value.fontSize }),
    ...(value.textColor === undefined ? {} : { textColor: value.textColor as string }),
    ...(value.backgroundColor === undefined ? {} : { backgroundColor: value.backgroundColor as string }),
    ...(value.numberFormat === undefined ? {} : { numberFormat: value.numberFormat as NonNullable<CellFormat["numberFormat"]> }),
    ...(value.horizontalAlign === undefined ? {} : { horizontalAlign: value.horizontalAlign as NonNullable<CellFormat["horizontalAlign"]> }),
    ...(value.verticalAlign === undefined ? {} : { verticalAlign: value.verticalAlign as NonNullable<CellFormat["verticalAlign"]> }),
    ...(value.wrapText === undefined ? {} : { wrapText: value.wrapText as boolean }),
    ...(borders === undefined ? {} : { borders })
  };
}

function migrateCellBorders(value: unknown): CellBorders | undefined | null {
  if (value === undefined) return undefined;
  if (!isRecord(value)) return null;
  const result: CellBorders = {};
  for (const key of ["top", "right", "bottom", "left"] as const) {
    const side = migrateCellBorderSide(value[key]);
    if (side === null) return null;
    if (side !== undefined) result[key] = side;
  }
  return result;
}

function migrateCellBorderSide(value: unknown): CellBorderSide | undefined | null {
  if (value === undefined) return undefined;
  return isRecord(value) && value.style === "thin" && typeof value.color === "string"
    ? { style: "thin", color: value.color }
    : null;
}

function migrateSheetFilters(value: unknown, bounds: SheetBounds): SheetFilter[] | null {
  if (value === undefined) return [];
  if (!Array.isArray(value)) return null;
  const result: SheetFilter[] = [];
  const ids = new Set<string>();
  for (const item of value) {
    const filter = migrateSheetFilter(item, bounds);
    if (!filter || ids.has(filter.id)) return null;
    ids.add(filter.id);
    result.push(filter);
  }
  return result;
}

export function migrateSheetFilter(value: unknown, bounds: SheetBounds): SheetFilter | null {
  if (!isRecord(value) || !nonBlankString(value.id)) return null;
  if (!isCellRange(value.range) || !rangeInBounds(value.range, bounds)) return null;
  if (!Number.isInteger(value.column)
    || (value.column as number) < value.range.start.column
    || (value.column as number) > value.range.end.column) return null;
  if (!["contains", "equals", "greaterThan", "lessThan"].includes(value.operator as string)) return null;
  if (typeof value.value !== "string" || !optionalBoolean(value.hasHeader)) return null;
  if (value.values !== undefined
    && (!Array.isArray(value.values) || !value.values.every((entry) => typeof entry === "string"))) return null;
  return {
    id: value.id,
    range: cloneRange(value.range),
    column: value.column as number,
    operator: value.operator as SheetFilter["operator"],
    value: value.value,
    ...(value.values === undefined ? {} : { values: [...value.values] }),
    ...(value.hasHeader === undefined ? {} : { hasHeader: value.hasHeader as boolean })
  };
}

function migrateSheetCharts(value: unknown, bounds: SheetBounds): SheetChart[] | null {
  if (value === undefined) return [];
  if (!Array.isArray(value)) return null;
  const result: SheetChart[] = [];
  const ids = new Set<string>();
  for (const item of value) {
    if (!isRecord(item) || !nonBlankString(item.id) || ids.has(item.id) || typeof item.title !== "string") return null;
    if (item.type !== "bar" && item.type !== "line" && item.type !== "pie") return null;
    if (!isCellRange(item.range) || !rangeInBounds(item.range, bounds)) return null;
    if (!isCellCoord(item.anchor) || !coordinateInBounds(item.anchor, bounds)) return null;
    ids.add(item.id);
    result.push({
      id: item.id,
      title: item.title,
      type: item.type,
      range: cloneRange(item.range),
      anchor: { ...item.anchor }
    });
  }
  return result;
}

function migrateSheetMerges(value: unknown, bounds: SheetBounds): SheetMerge[] | null {
  if (value === undefined) return [];
  if (!Array.isArray(value)) return null;
  const result: SheetMerge[] = [];
  const ids = new Set<string>();
  for (const item of value) {
    if (!isRecord(item) || !nonBlankString(item.id) || ids.has(item.id)) return null;
    if (!isCellRange(item.range) || !rangeInBounds(item.range, bounds)) return null;
    ids.add(item.id);
    result.push({ id: item.id, range: cloneRange(item.range) });
  }
  return result;
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
    const key = normalizeNamedRangeLookup(item.name);
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
  if (!validateExcelTableNameForInterop(value.name).valid || !isCellRange(value.range)) return null;
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

  const sort = migrateTableSort(value.sort, columnIds);
  if (sort === null) return null;
  const filter = migrateTableFilter(value.filter, columnIds);
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
  if (value.totalsFunction !== undefined && !isSupportedTableAggregate(value.totalsFunction)) return null;
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

export function isSupportedTableAggregate(value: unknown): value is TableAggregate {
  return typeof value === "string" && TOTALS_FUNCTIONS.has(value as TableAggregate);
}

export function migrateTableSort(
  value: unknown,
  columnIds: ReadonlySet<string>
): readonly TableSort[] | undefined | null {
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

export function migrateTableFilter(
  value: unknown,
  columnIds: ReadonlySet<string>
): FilterExpression | undefined | null {
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
    if (request.filter === null) return null;
    filter = request.filter;
  } catch {
    return null;
  }
  return everyFilterColumn(filter, columnIds) && hasSerializableNotInFilters(filter)
    ? filter
    : null;
}

function everyFilterColumn(filter: FilterExpression, columnIds: ReadonlySet<string>): boolean {
  if (filter.kind === "logical") return filter.operands.every((operand) => everyFilterColumn(operand, columnIds));
  if (filter.kind === "not") return everyFilterColumn(filter.operand, columnIds);
  return columnIds.has(filter.columnId);
}

function hasSerializableNotInFilters(filter: FilterExpression): boolean {
  if (filter.kind === "logical") {
    return filter.operands.every(hasSerializableNotInFilters);
  }
  if (filter.kind === "not") return hasSerializableNotInFilters(filter.operand);
  return filter.kind !== "set"
    || filter.operator !== "notIn"
    || (filter.values.length > 0 && filter.values.length <= 2);
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

function rangeInBounds(range: CellRange, sheet: SheetBounds): boolean {
  return range.start.row >= 0
    && range.start.column >= 0
    && range.end.row >= range.start.row
    && range.end.column >= range.start.column
    && range.end.row < sheet.rowCount
    && range.end.column < sheet.columnCount;
}

function coordinateInBounds(coordinate: CellRange["start"], sheet: SheetBounds): boolean {
  return coordinate.row >= 0
    && coordinate.column >= 0
    && coordinate.row < sheet.rowCount
    && coordinate.column < sheet.columnCount;
}

function canonicalCellAddressInBounds(address: string, sheet: SheetBounds): string | null {
  try {
    const coordinate = parseCellAddress(address);
    return coordinateInBounds(coordinate, sheet) ? formatCellAddress(coordinate) : null;
  } catch {
    return null;
  }
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

export function isMigratableCellContent(value: unknown): value is CellContent {
  return value === null
    || typeof value === "string"
    || typeof value === "boolean"
    || (typeof value === "number" && Number.isFinite(value));
}

function cellRecord(value: Record<string, unknown>): boolean {
  return Object.values(value).every(isMigratableCellContent);
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

function optionalBoolean(value: unknown): value is boolean | undefined {
  return value === undefined || typeof value === "boolean";
}

function optionalString(value: unknown): value is string | undefined {
  return value === undefined || typeof value === "string";
}

function optionalFiniteNumber(value: unknown): value is number | undefined {
  return value === undefined || (typeof value === "number" && Number.isFinite(value));
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
