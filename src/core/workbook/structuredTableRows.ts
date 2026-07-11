import type {
  CellContent,
  SheetModel,
  StructuredTable,
  WorkbookModel
} from "../../types";
import type { TableSort } from "../../table/core/query";
import type { ComputedCellValue } from "../../lib/formulaEngine";
import { compareDeterministicText } from "../../lib/filters";
import { formatCellAddress } from "../../lib/addressing";
import {
  rewriteFormulaForRectangularRowEdit,
  rewriteRowIntervalForRectangularEdit,
  translateFormulaReferences,
  translateFormulaRowsWithinColumns
} from "../../lib/formulaReferences";
import {
  getStructuredTable,
  getStructuredTableBodyRange,
  regenerateStructuredTableTotals,
  reduceStructuredTableCommand,
  type StructuredTableCommandServices,
  type StructuredTableReduction
} from "./structuredTables";
import { migrateTableSort } from "./migrateWorkbook";

export type TableCellPlanes = Pick<
  SheetModel,
  "cells" | "formats" | "validations" | "comments" | "hyperlinks"
>;

export function insertStructuredTableRows(
  workbook: WorkbookModel,
  tableId: string,
  input: { count: number; beforeRowId?: string; afterRowId?: string },
  services: StructuredTableCommandServices
): StructuredTableReduction {
  const table = getStructuredTable(workbook, tableId);
  if (!table) return rejected(workbook, "TABLE_NOT_FOUND", "Structured table does not exist");
  if (!Number.isSafeInteger(input.count) || input.count <= 0) {
    return rejected(workbook, "TABLE_ROW_COUNT_INVALID", "Inserted row count must be a positive integer");
  }
  if (input.beforeRowId !== undefined && input.afterRowId !== undefined) {
    return rejected(workbook, "TABLE_ROW_ANCHOR_INVALID", "Choose at most one table row anchor");
  }
  let insertionIndex = table.rowIds.length;
  if (input.beforeRowId !== undefined) insertionIndex = table.rowIds.indexOf(input.beforeRowId);
  if (input.afterRowId !== undefined) {
    const anchorIndex = table.rowIds.indexOf(input.afterRowId);
    insertionIndex = anchorIndex < 0 ? -1 : anchorIndex + 1;
  }
  if (insertionIndex < 0) {
    return rejected(workbook, "TABLE_ROW_ANCHOR_INVALID", "Table row anchor does not exist");
  }
  const body = getStructuredTableBodyRange(table);
  const insertionRow = body ? body.start.row + insertionIndex : table.range.start.row + Number(table.headerRow);
  const newBottom = table.range.end.row + input.count;
  const sheet = workbook.sheets.find((candidate) => candidate.id === table.sheetId)!;
  if (newBottom >= sheet.rowCount || !rowsAreEmpty(
    sheet,
    table.range.start.column,
    table.range.end.column,
    table.range.end.row + 1,
    newBottom
  )) {
    return rejected(workbook, "TABLE_RANGE_BLOCKED", "Table row insertion would overwrite unrelated data");
  }

  const rewritten = rewriteWorkbookForRowEdits(workbook, table, [{
    row: insertionRow,
    count: input.count,
    operation: "insert"
  }]);
  if (rewritten.status === "rejected") return rewritten;
  const resized = reduceStructuredTableCommand(rewritten.workbook, {
    type: "table.resize",
    tableId,
    range: {
      start: { ...table.range.start },
      end: { ...table.range.end, row: newBottom }
    }
  }, services);
  if (resized.status === "rejected") return resized;
  const resizedTable = getStructuredTable(resized.workbook, tableId)!;
  const generatedIds = resizedTable.rowIds.slice(table.rowIds.length);
  const mappings: RowMapping[] = [];
  for (let row = insertionRow; row <= table.range.end.row; row += 1) {
    mappings.push({ sourceRow: row, targetRow: row + input.count });
  }
  let next = remapTableRows(
    resized.workbook,
    table,
    mappings,
    insertionRow,
    newBottom
  );
  const rowIds = [
    ...table.rowIds.slice(0, insertionIndex),
    ...generatedIds,
    ...table.rowIds.slice(insertionIndex)
  ];
  next = replaceTable(next, { ...resizedTable, rowIds });
  const nextTable = getStructuredTable(next, tableId)!;
  next = regenerateCalculatedColumns(next, nextTable);
  next = regenerateStructuredTableTotals(next, nextTable);
  return { status: "committed", workbook: next };
}

export function deleteStructuredTableRows(
  workbook: WorkbookModel,
  tableId: string,
  rowIds: readonly string[],
  _services: StructuredTableCommandServices
): StructuredTableReduction {
  const table = getStructuredTable(workbook, tableId);
  if (!table) return rejected(workbook, "TABLE_NOT_FOUND", "Structured table does not exist");
  if (rowIds.length === 0) return { status: "unchanged", workbook };
  if (new Set(rowIds).size !== rowIds.length || rowIds.some((rowId) => !table.rowIds.includes(rowId))) {
    return rejected(workbook, "TABLE_ROW_ID_INVALID", "Every deleted row ID must exist exactly once");
  }
  const body = getStructuredTableBodyRange(table);
  if (!body) return rejected(workbook, "TABLE_ROW_ID_INVALID", "Structured table has no body rows");
  const deleting = new Set(rowIds);
  const deletedPhysicalRows = table.rowIds
    .map((rowId, index) => ({ rowId, row: body.start.row + index }))
    .filter(({ rowId }) => deleting.has(rowId))
    .map(({ row }) => row);
  const edits = contiguousDeletionEdits(deletedPhysicalRows);
  const rewritten = rewriteWorkbookForRowEdits(workbook, table, edits);
  if (rewritten.status === "rejected") return rewritten;
  const survivors = table.rowIds
    .map((rowId, index) => ({ rowId, sourceRow: body.start.row + index }))
    .filter(({ rowId }) => !deleting.has(rowId));
  const newBottom = table.range.end.row - rowIds.length;
  if (newBottom < table.range.start.row) {
    return rejected(workbook, "TABLE_RANGE_BLOCKED", "A table must retain at least one physical row");
  }
  const mappings: RowMapping[] = survivors.map((entry, index) => ({
    sourceRow: entry.sourceRow,
    targetRow: body.start.row + index
  }));
  if (table.totalsRow) {
    mappings.push({ sourceRow: table.range.end.row, targetRow: newBottom });
  }
  let next = remapTableRows(rewritten.workbook, table, mappings, body.start.row, table.range.end.row);
  const nextTable: StructuredTable = {
    ...table,
    range: { start: { ...table.range.start }, end: { ...table.range.end, row: newBottom } },
    rowIds: survivors.map(({ rowId }) => rowId)
  };
  next = replaceTable(next, nextTable);
  next = regenerateCalculatedColumns(next, nextTable);
  return { status: "committed", workbook: next };
}

export function setStructuredTableCalculatedColumn(
  workbook: WorkbookModel,
  tableId: string,
  columnId: string,
  formula: string | undefined,
  _services: StructuredTableCommandServices
): StructuredTableReduction {
  const table = getStructuredTable(workbook, tableId);
  if (!table) return rejected(workbook, "TABLE_NOT_FOUND", "Structured table does not exist");
  const columnIndex = table.columns.findIndex((column) => column.id === columnId);
  if (columnIndex < 0) return rejected(workbook, "TABLE_COLUMN_NOT_FOUND", "Structured table column does not exist");
  if (formula !== undefined && !formula.startsWith("=")) {
    return rejected(workbook, "TABLE_FORMULA_INVALID", "Calculated column formulas must begin with equals");
  }
  const current = table.columns[columnIndex].calculatedFormula;
  if (current === formula) return { status: "unchanged", workbook };
  const columns = table.columns.map((column, index) => {
    if (index !== columnIndex) return column;
    const { calculatedFormula: _current, ...base } = column;
    return formula === undefined ? base : { ...base, calculatedFormula: formula };
  });
  const nextTable = { ...table, columns };
  let next = replaceTable(workbook, nextTable);
  if (formula !== undefined) next = regenerateCalculatedColumns(next, nextTable);
  return { status: "committed", workbook: next };
}

export function sortStructuredTableRows(
  workbook: WorkbookModel,
  tableId: string,
  sorting: readonly TableSort[],
  services: StructuredTableCommandServices
): StructuredTableReduction {
  const table = getStructuredTable(workbook, tableId);
  if (!table) return rejected(workbook, "TABLE_NOT_FOUND", "Structured table does not exist");
  const migratedSorting = migrateTableSort(
    sorting,
    new Set(table.columns.map((column) => column.id))
  );
  if (!migratedSorting) {
    return rejected(workbook, "TABLE_SORT_INVALID", "Structured table sort is invalid");
  }
  const body = getStructuredTableBodyRange(table);
  if (!body || table.rowIds.length < 2) {
    const nextTable = { ...table, sort: migratedSorting };
    return { status: "committed", workbook: replaceTable(workbook, nextTable) };
  }
  const records = table.rowIds.map((rowId, index) => ({
    rowId,
    sourceRow: body.start.row + index,
    originalIndex: index
  }));
  records.sort((left, right) => {
    for (const sort of migratedSorting) {
      const column = table.columns.find((candidate) => candidate.id === sort.columnId)!;
      const leftValue = services.getCellEvaluation(
        table.sheetId,
        formatCellAddress({ row: left.sourceRow, column: column.sheetColumn })
      );
      const rightValue = services.getCellEvaluation(
        table.sheetId,
        formatCellAddress({ row: right.sourceRow, column: column.sheetColumn })
      );
      const comparison = compareComputedValues(leftValue, rightValue, sort.nulls ?? "last");
      if (comparison !== 0) {
        const hasSpecialValue = computedCategory(leftValue) !== 0 || computedCategory(rightValue) !== 0;
        return hasSpecialValue || sort.direction === "asc" ? comparison : -comparison;
      }
    }
    return left.originalIndex - right.originalIndex;
  });
  const mappings = records.map((record, index) => ({
    sourceRow: record.sourceRow,
    targetRow: body.start.row + index
  }));
  let next = remapTableRows(workbook, table, mappings, body.start.row, body.end.row, true);
  const nextTable = {
    ...table,
    rowIds: records.map(({ rowId }) => rowId),
    sort: migratedSorting
  };
  next = replaceTable(next, nextTable);
  next = regenerateCalculatedColumns(next, nextTable);
  return { status: "committed", workbook: next };
}

function regenerateCalculatedColumns(workbook: WorkbookModel, table: StructuredTable): WorkbookModel {
  const body = getStructuredTableBodyRange(table);
  if (!body) return workbook;
  let next = workbook;
  for (const column of table.columns) {
    if (!column.calculatedFormula) continue;
    for (let row = body.start.row; row <= body.end.row; row += 1) {
      next = setCellPlaneValue(
        next,
        table.sheetId,
        row,
        column.sheetColumn,
        translateFormulaReferences(column.calculatedFormula, {
          rowOffset: row - body.start.row,
          columnOffset: 0
        })
      );
    }
  }
  return next;
}

type RowMapping = { sourceRow: number; targetRow: number };
type RectangularEdit = { row: number; count: number; operation: "insert" | "delete" };
type RewriteOutcome =
  | { status: "committed"; workbook: WorkbookModel }
  | { status: "rejected"; workbook: WorkbookModel; issues: readonly { code: string; message: string }[] };

function rewriteWorkbookForRowEdits(
  workbook: WorkbookModel,
  table: StructuredTable,
  edits: readonly RectangularEdit[]
): RewriteOutcome {
  let candidate = workbook;
  let tableRowEnd = table.range.end.row;
  const editedSheet = workbook.sheets.find((sheet) => sheet.id === table.sheetId)!;
  for (const edit of edits) {
    const sheets: SheetModel[] = [];
    for (const sheet of candidate.sheets) {
      let cells: SheetModel["cells"] | undefined;
      for (const [address, value] of Object.entries(sheet.cells)) {
        if (typeof value !== "string" || !value.startsWith("=")) continue;
        const rewritten = rewriteFormulaForRectangularRowEdit(value, {
          formulaSheetId: sheet.name,
          editedSheetId: editedSheet.name,
          tableColumnStart: table.range.start.column,
          tableColumnEnd: table.range.end.column,
          tableRowEnd,
          row: edit.row,
          count: edit.count,
          operation: edit.operation,
          sheetBounds: { rowCount: editedSheet.rowCount, columnCount: editedSheet.columnCount }
        });
        if (!rewritten.ok) {
          return { status: "rejected", workbook, issues: [rewritten.issue] };
        }
        if (rewritten.formula !== value) {
          cells ??= { ...sheet.cells };
          cells[address] = rewritten.formula;
        }
      }
      sheets.push(cells ? { ...sheet, cells } : sheet);
    }
    let namedRanges: WorkbookModel["namedRanges"] | undefined;
    for (let index = 0; index < candidate.namedRanges.length; index += 1) {
      const namedRange = candidate.namedRanges[index];
      const rewritten = rewriteNamedRange(namedRange, table, edit, tableRowEnd);
      if (rewritten === "unsupported") {
        return {
          status: "rejected",
          workbook,
          issues: [{
            code: "TABLE_FORMULA_REFERENCE_UNSUPPORTED",
            message: `Named range ${namedRange.name} cannot represent this rectangular row edit`
          }]
        };
      }
      if (rewritten !== namedRange && !namedRanges) {
        namedRanges = candidate.namedRanges.slice(0, index);
      }
      namedRanges?.push(rewritten);
    }
    candidate = { ...candidate, sheets, ...(namedRanges ? { namedRanges } : {}) };
    tableRowEnd += edit.operation === "insert" ? edit.count : -edit.count;
  }
  return { status: "committed", workbook: candidate };
}

function rewriteNamedRange(
  namedRange: WorkbookModel["namedRanges"][number],
  table: StructuredTable,
  edit: RectangularEdit,
  tableRowEnd: number
): WorkbookModel["namedRanges"][number] | "unsupported" {
  if (namedRange.sheetId !== table.sheetId) return namedRange;
  const range = namedRange.range;
  const overlapsColumns = range.start.column <= table.range.end.column
    && range.end.column >= table.range.start.column;
  if (!overlapsColumns) return namedRange;
  const interval = rewriteInterval(
    range.start.row,
    range.end.row,
    edit,
    tableRowEnd
  );
  if (interval === "unchanged") return namedRange;
  const whollyInside = range.start.column >= table.range.start.column
    && range.end.column <= table.range.end.column;
  if (!whollyInside || interval === null) return "unsupported";
  return {
    ...namedRange,
    range: {
      start: { ...range.start, row: interval.start },
      end: { ...range.end, row: interval.end }
    }
  };
}

function rewriteInterval(
  start: number,
  end: number,
  edit: RectangularEdit,
  tableRowEnd: number
): { start: number; end: number } | null | "unchanged" {
  const interval = rewriteRowIntervalForRectangularEdit(start, end, {
    ...edit,
    tableRowEnd
  });
  if (interval?.start === start && interval.end === end) return "unchanged";
  return interval;
}

function contiguousDeletionEdits(rows: readonly number[]): RectangularEdit[] {
  const sorted = [...rows].sort((left, right) => left - right);
  const groups: RectangularEdit[] = [];
  for (const row of sorted) {
    const last = groups.at(-1);
    if (last && last.row + last.count === row) last.count += 1;
    else groups.push({ row, count: 1, operation: "delete" });
  }
  return groups.reverse();
}

function remapTableRows(
  workbook: WorkbookModel,
  table: StructuredTable,
  mappings: readonly RowMapping[],
  clearStartRow: number,
  clearEndRow: number,
  translateMovedFormulas = false
): WorkbookModel {
  const sheetIndex = workbook.sheets.findIndex((sheet) => sheet.id === table.sheetId);
  const sheet = workbook.sheets[sheetIndex];
  const remap = <T>(record: Readonly<Record<string, T>>, translateFormulas = false): Record<string, T> => {
    const next = { ...record };
    for (let row = clearStartRow; row <= clearEndRow; row += 1) {
      for (let column = table.range.start.column; column <= table.range.end.column; column += 1) {
        delete next[formatCellAddress({ row, column })];
      }
    }
    for (const mapping of mappings) {
      for (let column = table.range.start.column; column <= table.range.end.column; column += 1) {
        const source = formatCellAddress({ row: mapping.sourceRow, column });
        if (!Object.prototype.hasOwnProperty.call(record, source)) continue;
        let value = record[source];
        if (translateFormulas && typeof value === "string" && value.startsWith("=")) {
          value = translateFormulaRowsWithinColumns(value, {
            rowOffset: mapping.targetRow - mapping.sourceRow,
            formulaSheetName: sheet.name,
            columnStart: table.range.start.column,
            columnEnd: table.range.end.column
          }) as T;
        }
        next[formatCellAddress({ row: mapping.targetRow, column })] = value;
      }
    }
    return next;
  };
  const nextSheet: SheetModel = {
    ...sheet,
    cells: remap(sheet.cells, translateMovedFormulas),
    formats: remap(sheet.formats),
    validations: remap(sheet.validations),
    comments: remap(sheet.comments),
    hyperlinks: remap(sheet.hyperlinks)
  };
  return {
    ...workbook,
    sheets: workbook.sheets.map((candidate, index) => index === sheetIndex ? nextSheet : candidate)
  };
}

function rowsAreEmpty(
  sheet: SheetModel,
  startColumn: number,
  endColumn: number,
  startRow: number,
  endRow: number
): boolean {
  const planes: Readonly<Record<string, unknown>>[] = [
    sheet.cells,
    sheet.formats,
    sheet.validations,
    sheet.comments,
    sheet.hyperlinks
  ];
  for (let row = startRow; row <= endRow; row += 1) {
    for (let column = startColumn; column <= endColumn; column += 1) {
      const address = formatCellAddress({ row, column });
      if (planes.some((plane) => Object.prototype.hasOwnProperty.call(plane, address))) return false;
    }
  }
  return true;
}

function replaceTable(workbook: WorkbookModel, table: StructuredTable): WorkbookModel {
  return {
    ...workbook,
    tables: workbook.tables.map((candidate) => candidate.id === table.id ? table : candidate)
  };
}

function setCellPlaneValue(
  workbook: WorkbookModel,
  sheetId: string,
  row: number,
  column: number,
  value: CellContent
): WorkbookModel {
  const sheetIndex = workbook.sheets.findIndex((sheet) => sheet.id === sheetId);
  const sheet = workbook.sheets[sheetIndex];
  const address = formatCellAddress({ row, column });
  const cells = { ...sheet.cells, [address]: value };
  return {
    ...workbook,
    sheets: workbook.sheets.map((candidate, index) => index === sheetIndex ? { ...sheet, cells } : candidate)
  };
}

function compareComputedValues(
  left: ComputedCellValue,
  right: ComputedCellValue,
  nulls: "first" | "last"
): number {
  const leftCategory = computedCategory(left);
  const rightCategory = computedCategory(right);
  if (leftCategory === 2 || rightCategory === 2) {
    if (leftCategory === rightCategory) return 0;
    return leftCategory === 2 ? (nulls === "first" ? -1 : 1) : (nulls === "first" ? 1 : -1);
  }
  if (leftCategory !== rightCategory) return leftCategory - rightCategory;
  if (leftCategory === 1) return 0;
  if (typeof left === "number" && typeof right === "number") return left - right;
  if (typeof left === "boolean" && typeof right === "boolean") return Number(left) - Number(right);
  if (typeof left === "string" && typeof right === "string") return compareDeterministicText(left, right);
  return ordinaryTypeRank(left) - ordinaryTypeRank(right);
}

function computedCategory(value: ComputedCellValue): 0 | 1 | 2 {
  if (value === null || value === "") return 2;
  if (typeof value === "object") return 1;
  return 0;
}

function ordinaryTypeRank(value: ComputedCellValue): number {
  if (typeof value === "number") return 0;
  if (typeof value === "boolean") return 1;
  if (typeof value === "string") return 2;
  return 3;
}

function rejected(
  workbook: WorkbookModel,
  code: string,
  message: string
): StructuredTableReduction {
  return { status: "rejected", workbook, issues: [{ code, message }] };
}
