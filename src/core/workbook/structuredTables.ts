import type { TableIssue } from "../commands/types";
import type { IdGenerator } from "../ids";
import type {
  CellCoord,
  CellFormat,
  CellRange,
  StructuredTable,
  StructuredTableColumn,
  TableAggregate,
  TableStyle,
  WorkbookModel
} from "../../types";
import type { ComputedCellValue } from "../../lib/formulaEngine";
import { formatCellAddress } from "../../lib/addressing";
import { translateFormulaReferences } from "../../lib/formulaReferences";
import { getCellContent, getCellReadOnly } from "../../lib/workbook";
import type { FilterExpression, TableSort } from "../../table/core/query";
import { normalizeExcelTableNameKey, validateExcelTableName } from "./tableNames";

export type StructuredTableCommand =
  | {
      type: "table.create";
      sheetId: string;
      range: CellRange;
      name?: string;
      headerRow?: boolean;
      totalsRow?: boolean;
      style?: TableStyle;
    }
  | { type: "table.rename"; tableId: string; name: string }
  | { type: "table.renameColumn"; tableId: string; columnId: string; name: string }
  | { type: "table.resize"; tableId: string; range: CellRange }
  | { type: "table.setHeaderRow"; tableId: string; enabled: boolean }
  | { type: "table.setTotalsRow"; tableId: string; enabled: boolean }
  | { type: "table.setTotalsFunction"; tableId: string; columnId: string; aggregate: TableAggregate }
  | { type: "table.setStyle"; tableId: string; style: TableStyle }
  | { type: "table.setKeyColumn"; tableId: string; columnId?: string }
  | { type: "table.setCalculatedColumn"; tableId: string; columnId: string; formula?: string }
  | { type: "table.setFilter"; tableId: string; filter?: FilterExpression }
  | { type: "table.sort"; tableId: string; sorting: readonly TableSort[] }
  | {
      type: "table.insertRows";
      tableId: string;
      count: number;
      beforeRowId?: string;
      afterRowId?: string;
    }
  | { type: "table.deleteRows"; tableId: string; rowIds: readonly string[] }
  | {
      type: "table.editCells";
      tableId: string;
      edits: readonly { rowId: string; columnId: string; rawText: string }[];
    }
  | { type: "table.convertToRange"; tableId: string };

export type StructuredTableCommandServices = {
  createId: IdGenerator;
  getCellEvaluation(sheetId: string, address: string): ComputedCellValue;
};

export type StructuredTableReduction =
  | { status: "committed"; workbook: WorkbookModel }
  | { status: "unchanged"; workbook: WorkbookModel }
  | { status: "rejected"; workbook: WorkbookModel; issues: readonly TableIssue[] };

const DEFAULT_TABLE_STYLE: TableStyle = Object.freeze({
  theme: "TableStyleLight1",
  showRowStripes: true
});

export function getStructuredTable(workbook: WorkbookModel, tableId: string): StructuredTable | null {
  return workbook.tables.find((table) => table.id === tableId) ?? null;
}

export function getStructuredTableAtCell(
  workbook: WorkbookModel,
  sheetId: string,
  coord: CellCoord
): StructuredTable | null {
  return workbook.tables.find((table) => table.sheetId === sheetId && contains(table.range, coord)) ?? null;
}

export function getStructuredTableForSelection(
  workbook: WorkbookModel,
  sheetId: string,
  selection: CellRange
): StructuredTable | null {
  const normalized = normalize(selection);
  return getStructuredTableAtCell(workbook, sheetId, selection.start)
    ?? workbook.tables.find((table) => table.sheetId === sheetId && intersects(table.range, normalized))
    ?? null;
}

export function getStructuredTableBodyRange(table: StructuredTable): CellRange | null {
  const firstBodyRow = table.range.start.row + (table.headerRow ? 1 : 0);
  const lastBodyRow = table.range.end.row - (table.totalsRow ? 1 : 0);
  return firstBodyRow > lastBodyRow
    ? null
    : {
        start: { row: firstBodyRow, column: table.range.start.column },
        end: { row: lastBodyRow, column: table.range.end.column }
      };
}

export function reduceStructuredTableCommand(
  workbook: WorkbookModel,
  command: StructuredTableCommand,
  services: StructuredTableCommandServices
): StructuredTableReduction {
  switch (command.type) {
    case "table.create":
      return createTable(workbook, command, services);
    case "table.rename":
      return renameTable(workbook, command.tableId, command.name);
    case "table.renameColumn":
      return renameColumn(workbook, command.tableId, command.columnId, command.name);
    case "table.resize":
      return resizeTable(workbook, command.tableId, command.range, services);
    case "table.setHeaderRow":
      return setHeaderRow(workbook, command.tableId, command.enabled);
    case "table.setTotalsRow":
      return setTotalsRow(workbook, command.tableId, command.enabled);
    case "table.setTotalsFunction":
      return setTotalsFunction(workbook, command.tableId, command.columnId, command.aggregate);
    case "table.setStyle":
      return setStyle(workbook, command.tableId, command.style);
    case "table.setKeyColumn":
      return setKeyColumn(workbook, command.tableId, command.columnId);
    case "table.convertToRange":
      return convertToRange(workbook, command.tableId);
    case "table.setCalculatedColumn":
    case "table.setFilter":
    case "table.sort":
    case "table.insertRows":
    case "table.deleteRows":
    case "table.editCells":
      return reject(workbook, "TABLE_COMMAND_UNSUPPORTED", "Structured table command is not available yet");
  }
}

function createTable(
  workbook: WorkbookModel,
  command: Extract<StructuredTableCommand, { type: "table.create" }>,
  services: StructuredTableCommandServices
): StructuredTableReduction {
  const sheet = workbook.sheets.find((candidate) => candidate.id === command.sheetId);
  if (!sheet) return reject(workbook, "TABLE_RANGE_BLOCKED", "Table sheet does not exist");
  const name = command.name ?? nextTableName(workbook);
  const nameIssue = validateName(workbook, name);
  if (nameIssue) return { status: "rejected", workbook, issues: [nameIssue] };
  const range = cloneRange(command.range);
  const rangeIssue = validateRange(workbook, command.sheetId, range);
  if (rangeIssue) return { status: "rejected", workbook, issues: [rangeIssue] };
  const headerRow = command.headerRow ?? true;
  const totalsRow = command.totalsRow ?? false;
  const height = range.end.row - range.start.row + 1;
  if (height < Number(headerRow) + Number(totalsRow)) {
    return reject(workbook, "TABLE_RANGE_BLOCKED", "Table range has no room for its configured rows");
  }
  const headerNames = headerRow
    ? readHeaderNames(workbook, command.sheetId, range)
    : Array.from({ length: range.end.column - range.start.column + 1 }, (_, index) => `Column${index + 1}`);
  if (!headerNames) return reject(workbook, "TABLE_HEADER_INVALID", "Table headers must be nonblank and unique");
  const style = command.style ?? DEFAULT_TABLE_STYLE;
  if (!validStyle(style)) return reject(workbook, "TABLE_RANGE_BLOCKED", "Table style is invalid");

  const columns = headerNames.map((columnName, index): StructuredTableColumn => ({
    id: services.createId("table-column"),
    name: columnName,
    sheetColumn: range.start.column + index
  }));
  const bodyCount = height - Number(headerRow) - Number(totalsRow);
  const table: StructuredTable = {
    id: services.createId("table"),
    name,
    sheetId: command.sheetId,
    range,
    headerRow,
    totalsRow,
    columns,
    rowIds: Array.from({ length: bodyCount }, () => services.createId("table-row")),
    style: cloneStyle(style)
  };
  return { status: "committed", workbook: { ...workbook, tables: [...workbook.tables, table] } };
}

function renameTable(workbook: WorkbookModel, tableId: string, name: string): StructuredTableReduction {
  const table = getStructuredTable(workbook, tableId);
  if (!table) return tableNotFound(workbook);
  const nameIssue = validateName(workbook, name, tableId);
  if (nameIssue) return { status: "rejected", workbook, issues: [nameIssue] };
  if (table.name === name) return { status: "unchanged", workbook };
  return commitTable(workbook, { ...table, name });
}

function renameColumn(
  workbook: WorkbookModel,
  tableId: string,
  columnId: string,
  name: string
): StructuredTableReduction {
  const table = getStructuredTable(workbook, tableId);
  if (!table) return tableNotFound(workbook);
  const columnIndex = table.columns.findIndex((column) => column.id === columnId);
  if (columnIndex < 0) return reject(workbook, "TABLE_COLUMN_NOT_FOUND", "Structured table column does not exist");
  if (name.trim().length === 0 || table.columns.some(
    (column, index) => index !== columnIndex && normalizeHeader(column.name) === normalizeHeader(name)
  )) {
    return reject(workbook, "TABLE_HEADER_INVALID", "Table headers must be nonblank and unique");
  }
  if (table.columns[columnIndex].name === name) return { status: "unchanged", workbook };
  const columns = table.columns.map((column, index) => index === columnIndex ? { ...column, name } : column);
  let nextWorkbook = workbook;
  if (table.headerRow) {
    nextWorkbook = setRawCell(
      nextWorkbook,
      table.sheetId,
      { row: table.range.start.row, column: columns[columnIndex].sheetColumn },
      name
    );
  }
  return commitTable(nextWorkbook, { ...table, columns });
}

function resizeTable(
  workbook: WorkbookModel,
  tableId: string,
  requestedRange: CellRange,
  services: StructuredTableCommandServices
): StructuredTableReduction {
  const table = getStructuredTable(workbook, tableId);
  if (!table) return tableNotFound(workbook);
  const range = cloneRange(requestedRange);
  if (range.start.row !== table.range.start.row || range.start.column !== table.range.start.column) {
    return reject(workbook, "TABLE_RANGE_BLOCKED", "Table resize must preserve its top-left cell");
  }
  const issue = validateRange(workbook, table.sheetId, range, table.id);
  if (issue) return { status: "rejected", workbook, issues: [issue] };
  const height = range.end.row - range.start.row + 1;
  if (height < Number(table.headerRow) + Number(table.totalsRow)) {
    return reject(workbook, "TABLE_RANGE_BLOCKED", "Table range has no room for its configured rows");
  }
  if (sameRange(range, table.range)) return { status: "unchanged", workbook };

  const width = range.end.column - range.start.column + 1;
  const columns = table.columns.slice(0, width).map((column) => ({ ...column }));
  const names = new Set(columns.map((column) => normalizeHeader(column.name)));
  let nextWorkbook = workbook;
  for (let index = columns.length; index < width; index += 1) {
    const sheetColumn = range.start.column + index;
    const headerValue = table.headerRow
      ? getCellContent(workbook, table.sheetId, formatCellAddress({ row: range.start.row, column: sheetColumn }))
      : null;
    let name: string;
    if (table.headerRow && headerValue !== null && (typeof headerValue !== "string" || headerValue.trim().length === 0)) {
      return reject(workbook, "TABLE_HEADER_INVALID", "Table headers must be nonblank text");
    }
    if (typeof headerValue === "string" && headerValue.trim().length > 0) {
      name = headerValue;
      if (names.has(normalizeHeader(name))) {
        return reject(workbook, "TABLE_HEADER_INVALID", "Table headers must be unique");
      }
    } else {
      name = nextColumnName(index + 1, names);
      if (table.headerRow) {
        nextWorkbook = setRawCell(nextWorkbook, table.sheetId, {
          row: range.start.row,
          column: sheetColumn
        }, name);
      }
    }
    names.add(normalizeHeader(name));
    columns.push({ id: services.createId("table-column"), name, sheetColumn });
  }
  const bodyCount = height - Number(table.headerRow) - Number(table.totalsRow);
  const rowIds = table.rowIds.slice(0, bodyCount);
  while (rowIds.length < bodyCount) rowIds.push(services.createId("table-row"));
  const survivingColumnIds = new Set(columns.map((column) => column.id));
  const keyColumnId = table.keyColumnId && survivingColumnIds.has(table.keyColumnId)
    ? table.keyColumnId
    : undefined;
  const sort = table.sort?.filter((entry) => survivingColumnIds.has(entry.columnId));
  const filter = table.filter && everyFilterColumn(table.filter, survivingColumnIds)
    ? table.filter
    : undefined;
  const { keyColumnId: _key, sort: _sort, filter: _filter, ...base } = table;
  return commitTable(nextWorkbook, {
    ...base,
    range,
    columns,
    rowIds,
    ...(keyColumnId === undefined ? {} : { keyColumnId }),
    ...(sort === undefined ? {} : { sort }),
    ...(filter === undefined ? {} : { filter })
  });
}

function setHeaderRow(workbook: WorkbookModel, tableId: string, enabled: boolean): StructuredTableReduction {
  const table = getStructuredTable(workbook, tableId);
  if (!table) return tableNotFound(workbook);
  if (table.headerRow === enabled) return { status: "unchanged", workbook };

  if (enabled) {
    const range = { ...cloneRange(table.range), end: { ...table.range.end, row: table.range.end.row + 1 } };
    const issue = validateRange(workbook, table.sheetId, range, table.id);
    if (issue) return { status: "rejected", workbook, issues: [issue] };
    if (!rowSliceIsEmpty(workbook, table.sheetId, table.range.end.row + 1, table.range)) {
      return reject(workbook, "TABLE_RANGE_BLOCKED", "Enabling headers would overwrite populated cells");
    }
    let next = shiftTableSlice(workbook, table, table.range.start.row, table.range.end.row, 1);
    for (const column of table.columns) {
      next = setRawCell(next, table.sheetId, { row: table.range.start.row, column: column.sheetColumn }, column.name);
    }
    return commitTable(next, { ...table, range, headerRow: true });
  }

  if (table.range.end.row - 1 < table.range.start.row) {
    return reject(workbook, "TABLE_RANGE_BLOCKED", "A table must retain at least one physical row");
  }

  let next = shiftTableSlice(workbook, table, table.range.start.row + 1, table.range.end.row, -1);
  next = clearRowSlice(next, table.sheetId, table.range.end.row, table.range);
  const range = { ...cloneRange(table.range), end: { ...table.range.end, row: table.range.end.row - 1 } };
  return commitTable(next, { ...table, range, headerRow: false });
}

function setTotalsRow(workbook: WorkbookModel, tableId: string, enabled: boolean): StructuredTableReduction {
  const table = getStructuredTable(workbook, tableId);
  if (!table) return tableNotFound(workbook);
  if (table.totalsRow === enabled) return { status: "unchanged", workbook };
  if (enabled) {
    const range = { ...cloneRange(table.range), end: { ...table.range.end, row: table.range.end.row + 1 } };
    const issue = validateRange(workbook, table.sheetId, range, table.id);
    if (issue) return { status: "rejected", workbook, issues: [issue] };
    if (!rowSliceIsEmpty(workbook, table.sheetId, range.end.row, table.range)) {
      return reject(workbook, "TABLE_RANGE_BLOCKED", "Enabling totals would overwrite populated cells");
    }
    return commitTable(workbook, { ...table, range, totalsRow: true });
  }
  if (table.range.end.row - 1 < table.range.start.row) {
    return reject(workbook, "TABLE_RANGE_BLOCKED", "A table must retain at least one physical row");
  }
  const next = clearRowSlice(workbook, table.sheetId, table.range.end.row, table.range);
  const columns = table.columns.map(({ totalsFunction: _function, totalsLabel: _label, ...column }) => column);
  const range = { ...cloneRange(table.range), end: { ...table.range.end, row: table.range.end.row - 1 } };
  return commitTable(next, { ...table, range, totalsRow: false, columns });
}

function setTotalsFunction(
  workbook: WorkbookModel,
  tableId: string,
  columnId: string,
  aggregate: TableAggregate
): StructuredTableReduction {
  const table = getStructuredTable(workbook, tableId);
  if (!table) return tableNotFound(workbook);
  const columnIndex = table.columns.findIndex((column) => column.id === columnId);
  if (columnIndex < 0) return reject(workbook, "TABLE_COLUMN_NOT_FOUND", "Structured table column does not exist");
  if (!table.totalsRow && aggregate !== "none") {
    return reject(workbook, "TABLE_TOTALS_DISABLED", "Enable the totals row before choosing an aggregate");
  }
  const current = table.columns[columnIndex];
  if ((current.totalsFunction ?? "none") === aggregate && current.totalsLabel === undefined) {
    return { status: "unchanged", workbook };
  }
  const columns = table.columns.map((column, index) => {
    if (index !== columnIndex) return column;
    const { totalsFunction: _function, totalsLabel: _label, ...base } = column;
    return aggregate === "none" ? base : { ...base, totalsFunction: aggregate };
  });
  let next = workbook;
  if (table.totalsRow) {
    const coord = { row: table.range.end.row, column: current.sheetColumn };
    const formula = aggregate === "none" ? null : totalsFormula(table, current, aggregate);
    next = setRawCell(next, table.sheetId, coord, formula);
  }
  return commitTable(next, { ...table, columns });
}

function setStyle(workbook: WorkbookModel, tableId: string, style: TableStyle): StructuredTableReduction {
  const table = getStructuredTable(workbook, tableId);
  if (!table) return tableNotFound(workbook);
  if (!validStyle(style)) return reject(workbook, "TABLE_RANGE_BLOCKED", "Table style is invalid");
  if (JSON.stringify(table.style) === JSON.stringify(style)) return { status: "unchanged", workbook };
  return commitTable(workbook, { ...table, style: cloneStyle(style) });
}

function setKeyColumn(
  workbook: WorkbookModel,
  tableId: string,
  columnId: string | undefined
): StructuredTableReduction {
  const table = getStructuredTable(workbook, tableId);
  if (!table) return tableNotFound(workbook);
  if (columnId !== undefined && !table.columns.some((column) => column.id === columnId)) {
    return reject(workbook, "TABLE_COLUMN_NOT_FOUND", "Structured table column does not exist");
  }
  if (table.keyColumnId === columnId) return { status: "unchanged", workbook };
  const { keyColumnId: _current, ...base } = table;
  return commitTable(workbook, columnId === undefined ? base : { ...base, keyColumnId: columnId });
}

function convertToRange(workbook: WorkbookModel, tableId: string): StructuredTableReduction {
  const table = getStructuredTable(workbook, tableId);
  if (!table) return tableNotFound(workbook);
  const sheetIndex = workbook.sheets.findIndex((sheet) => sheet.id === table.sheetId);
  const sheet = workbook.sheets[sheetIndex];
  const formats = { ...sheet.formats };
  for (let row = table.range.start.row; row <= table.range.end.row; row += 1) {
    for (let column = table.range.start.column; column <= table.range.end.column; column += 1) {
      const address = formatCellAddress({ row, column });
      const style = materializedStyle(table, row);
      if (formats[address] !== undefined || Object.keys(style).length > 0) {
        formats[address] = { ...(formats[address] ?? {}), ...style };
      }
    }
  }
  const sheets = workbook.sheets.map((candidate, index) => index === sheetIndex ? { ...candidate, formats } : candidate);
  return {
    status: "committed",
    workbook: { ...workbook, sheets, tables: workbook.tables.filter((candidate) => candidate.id !== tableId) }
  };
}

function validateName(workbook: WorkbookModel, name: string, ownTableId?: string): TableIssue | null {
  const validation = validateExcelTableName(name);
  if (!validation.valid) return issue("TABLE_NAME_INVALID", `Invalid Excel table name (${validation.reason})`);
  return workbook.tables.some(
    (table) => table.id !== ownTableId && normalizeExcelTableNameKey(table.name) === validation.normalizedKey
  ) ? issue("TABLE_NAME_CONFLICT", "Table names must be unique") : null;
}

function validateRange(
  workbook: WorkbookModel,
  sheetId: string,
  range: CellRange,
  ownTableId?: string
): TableIssue | null {
  const sheet = workbook.sheets.find((candidate) => candidate.id === sheetId);
  if (!sheet || !validRange(range)
    || range.end.row >= sheet.rowCount || range.end.column >= sheet.columnCount) {
    return issue("TABLE_RANGE_BLOCKED", "Table range is outside the worksheet");
  }
  if (workbook.tables.some(
    (table) => table.id !== ownTableId && table.sheetId === sheetId && intersects(table.range, range)
  )) return issue("TABLE_RANGE_OVERLAP", "Table ranges cannot overlap");
  if (sheet.merges.some((merge) => intersects(merge.range, range))) {
    return issue("TABLE_MERGE_CONFLICT", "Table range intersects a merged cell");
  }
  if (sheet.protection.isProtected) return issue("TABLE_PROTECTED", "Protected sheets cannot change tables");
  for (let row = range.start.row; row <= range.end.row; row += 1) {
    for (let column = range.start.column; column <= range.end.column; column += 1) {
      if (getCellReadOnly(workbook, sheetId, formatCellAddress({ row, column }))) {
        return issue("TABLE_PROTECTED", "Table range contains a locked cell");
      }
    }
  }
  return null;
}

function readHeaderNames(workbook: WorkbookModel, sheetId: string, range: CellRange): string[] | null {
  const names: string[] = [];
  const keys = new Set<string>();
  for (let column = range.start.column; column <= range.end.column; column += 1) {
    const value = getCellContent(workbook, sheetId, formatCellAddress({ row: range.start.row, column }));
    if (typeof value !== "string" || value.trim().length === 0) return null;
    const key = normalizeHeader(value);
    if (keys.has(key)) return null;
    keys.add(key);
    names.push(value);
  }
  return names;
}

function everyFilterColumn(filter: FilterExpression, columnIds: ReadonlySet<string>): boolean {
  if (filter.kind === "logical") return filter.operands.every((operand) => everyFilterColumn(operand, columnIds));
  if (filter.kind === "not") return everyFilterColumn(filter.operand, columnIds);
  return columnIds.has(filter.columnId);
}

function shiftTableSlice(
  workbook: WorkbookModel,
  table: StructuredTable,
  startRow: number,
  endRow: number,
  rowOffset: number
): WorkbookModel {
  const sheetIndex = workbook.sheets.findIndex((sheet) => sheet.id === table.sheetId);
  const sheet = workbook.sheets[sheetIndex];
  const cells = shiftAddressRecord(sheet.cells, table.range, startRow, endRow, rowOffset, true);
  const formats = shiftAddressRecord(sheet.formats, table.range, startRow, endRow, rowOffset);
  const validations = shiftAddressRecord(sheet.validations, table.range, startRow, endRow, rowOffset);
  const comments = shiftAddressRecord(sheet.comments, table.range, startRow, endRow, rowOffset);
  const hyperlinks = shiftAddressRecord(sheet.hyperlinks, table.range, startRow, endRow, rowOffset);
  const protection = {
    ...sheet.protection,
    lockedCells: shiftAddressRecord(sheet.protection.lockedCells, table.range, startRow, endRow, rowOffset),
    unlockedCells: shiftAddressRecord(sheet.protection.unlockedCells, table.range, startRow, endRow, rowOffset)
  };
  const nextSheet = { ...sheet, cells, formats, validations, comments, hyperlinks, protection };
  return {
    ...workbook,
    sheets: workbook.sheets.map((candidate, index) => index === sheetIndex ? nextSheet : candidate)
  };
}

function shiftAddressRecord<T>(
  record: Readonly<Record<string, T>>,
  tableRange: CellRange,
  startRow: number,
  endRow: number,
  rowOffset: number,
  translateFormulas = false
): Record<string, T> {
  const next = { ...record };
  const sources: Array<{ source: string; target: string; value: T }> = [];
  for (let row = startRow; row <= endRow; row += 1) {
    for (let column = tableRange.start.column; column <= tableRange.end.column; column += 1) {
      const source = formatCellAddress({ row, column });
      const target = formatCellAddress({ row: row + rowOffset, column });
      delete next[source];
      delete next[target];
      if (Object.prototype.hasOwnProperty.call(record, source)) {
        let value = record[source];
        if (translateFormulas && typeof value === "string" && value.startsWith("=")) {
          value = translateFormulaReferences(value, { rowOffset, columnOffset: 0 }) as T;
        }
        sources.push({ source, target, value });
      }
    }
  }
  for (const { target, value } of sources) next[target] = value;
  return next;
}

function clearRowSlice(
  workbook: WorkbookModel,
  sheetId: string,
  row: number,
  tableRange: CellRange
): WorkbookModel {
  const sheetIndex = workbook.sheets.findIndex((sheet) => sheet.id === sheetId);
  const sheet = workbook.sheets[sheetIndex];
  const clear = <T>(record: Readonly<Record<string, T>>): Record<string, T> => {
    const next = { ...record };
    for (let column = tableRange.start.column; column <= tableRange.end.column; column += 1) {
      delete next[formatCellAddress({ row, column })];
    }
    return next;
  };
  const nextSheet = {
    ...sheet,
    cells: clear(sheet.cells),
    formats: clear(sheet.formats),
    validations: clear(sheet.validations),
    comments: clear(sheet.comments),
    hyperlinks: clear(sheet.hyperlinks),
    protection: {
      ...sheet.protection,
      lockedCells: clear(sheet.protection.lockedCells),
      unlockedCells: clear(sheet.protection.unlockedCells)
    }
  };
  return { ...workbook, sheets: workbook.sheets.map((candidate, index) => index === sheetIndex ? nextSheet : candidate) };
}

function rowSliceIsEmpty(workbook: WorkbookModel, sheetId: string, row: number, tableRange: CellRange): boolean {
  const sheet = workbook.sheets.find((candidate) => candidate.id === sheetId)!;
  const records: Readonly<Record<string, unknown>>[] = [
    sheet.cells,
    sheet.formats,
    sheet.validations,
    sheet.comments,
    sheet.hyperlinks,
    sheet.protection.lockedCells,
    sheet.protection.unlockedCells
  ];
  for (let column = tableRange.start.column; column <= tableRange.end.column; column += 1) {
    const address = formatCellAddress({ row, column });
    if (records.some((record) => Object.prototype.hasOwnProperty.call(record, address))) return false;
  }
  return true;
}

function setRawCell(
  workbook: WorkbookModel,
  sheetId: string,
  coord: CellCoord,
  value: string | number | boolean | null
): WorkbookModel {
  const sheetIndex = workbook.sheets.findIndex((sheet) => sheet.id === sheetId);
  const sheet = workbook.sheets[sheetIndex];
  const address = formatCellAddress(coord);
  const cells = { ...sheet.cells };
  if (value === null || value === "") delete cells[address];
  else cells[address] = value;
  return {
    ...workbook,
    sheets: workbook.sheets.map((candidate, index) => index === sheetIndex ? { ...sheet, cells } : candidate)
  };
}

function commitTable(workbook: WorkbookModel, table: StructuredTable): StructuredTableReduction {
  const index = workbook.tables.findIndex((candidate) => candidate.id === table.id);
  if (index < 0) return tableNotFound(workbook);
  const tables = [...workbook.tables];
  tables[index] = table;
  return { status: "committed", workbook: { ...workbook, tables } };
}

function totalsFormula(table: StructuredTable, column: StructuredTableColumn, aggregate: TableAggregate): string {
  const body = getStructuredTableBodyRange(table);
  if (!body) return aggregate === "count" || aggregate === "countNumbers" ? "=0" : "";
  const start = formatCellAddress({ row: body.start.row, column: column.sheetColumn });
  const end = formatCellAddress({ row: body.end.row, column: column.sheetColumn });
  const code: Record<Exclude<TableAggregate, "none">, number> = {
    sum: 109,
    average: 101,
    count: 103,
    countNumbers: 102,
    min: 105,
    max: 104,
    standardDeviation: 107,
    variance: 110
  };
  return `=SUBTOTAL(${code[aggregate as Exclude<TableAggregate, "none">]},${start}:${end})`;
}

function materializedStyle(table: StructuredTable, row: number): Partial<CellFormat> {
  if (table.headerRow && row === table.range.start.row) {
    return { bold: true, textColor: "#0f172a", backgroundColor: "#e2e8f0" };
  }
  if (table.totalsRow && row === table.range.end.row) {
    return { bold: true, backgroundColor: "#f1f5f9" };
  }
  const body = getStructuredTableBodyRange(table);
  if (body && table.style?.showRowStripes && (row - body.start.row) % 2 === 1) {
    return { backgroundColor: "#f8fafc" };
  }
  return {};
}

function nextTableName(workbook: WorkbookModel): string {
  const existing = new Set(workbook.tables.map((table) => normalizeExcelTableNameKey(table.name)));
  for (let index = 1; ; index += 1) {
    const candidate = `Table${index}`;
    if (!existing.has(normalizeExcelTableNameKey(candidate))) return candidate;
  }
}

function nextColumnName(start: number, existing: ReadonlySet<string>): string {
  for (let index = start; ; index += 1) {
    const candidate = `Column${index}`;
    if (!existing.has(normalizeHeader(candidate))) return candidate;
  }
}

function validStyle(style: TableStyle): boolean {
  return style !== null
    && typeof style === "object"
    && typeof style.theme === "string"
    && style.theme.trim().length > 0
    && [style.showFirstColumn, style.showLastColumn, style.showRowStripes, style.showColumnStripes]
      .every((value) => value === undefined || typeof value === "boolean");
}

function cloneStyle(style: TableStyle): TableStyle {
  return { ...style };
}

function normalizeHeader(name: string): string {
  return name.normalize("NFKC").toLowerCase();
}

function validRange(range: CellRange): boolean {
  return Number.isInteger(range.start.row)
    && Number.isInteger(range.start.column)
    && Number.isInteger(range.end.row)
    && Number.isInteger(range.end.column)
    && range.start.row >= 0
    && range.start.column >= 0
    && range.end.row >= range.start.row
    && range.end.column >= range.start.column;
}

function normalize(range: CellRange): CellRange {
  return {
    start: {
      row: Math.min(range.start.row, range.end.row),
      column: Math.min(range.start.column, range.end.column)
    },
    end: {
      row: Math.max(range.start.row, range.end.row),
      column: Math.max(range.start.column, range.end.column)
    }
  };
}

function contains(range: CellRange, coord: CellCoord): boolean {
  return coord.row >= range.start.row && coord.row <= range.end.row
    && coord.column >= range.start.column && coord.column <= range.end.column;
}

function intersects(left: CellRange, right: CellRange): boolean {
  return left.start.row <= right.end.row && left.end.row >= right.start.row
    && left.start.column <= right.end.column && left.end.column >= right.start.column;
}

function sameRange(left: CellRange, right: CellRange): boolean {
  return left.start.row === right.start.row && left.start.column === right.start.column
    && left.end.row === right.end.row && left.end.column === right.end.column;
}

function cloneRange(range: CellRange): CellRange {
  return { start: { ...range.start }, end: { ...range.end } };
}

function issue(code: string, message: string): TableIssue {
  return { code, message };
}

function reject(workbook: WorkbookModel, code: string, message: string): StructuredTableReduction {
  return { status: "rejected", workbook, issues: [issue(code, message)] };
}

function tableNotFound(workbook: WorkbookModel): StructuredTableReduction {
  return reject(workbook, "TABLE_NOT_FOUND", "Structured table does not exist");
}
