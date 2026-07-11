import type { CellRange, StructuredTable, StructuredTableColumn, WorkbookModel } from "../../types";
import { formatCellAddress } from "../../lib/addressing";
import { rewriteFormulaForStructure } from "../../lib/formulaReferences";
import { shiftSheetStructurePlanes } from "../../lib/workbook";
import type { FilterExpression } from "../../table/core/query";
import type { TableIssue } from "../commands/types";
import type { IdGenerator } from "../ids";
import {
  nextStructuredTableColumnName,
  normalizeStructuredTableHeader
} from "./structuredTables";

export type WorksheetStructureCommand =
  | { type: "rows.insert" | "rows.delete"; sheetId: string; index: number; count: number }
  | {
      type: "columns.insert";
      sheetId: string;
      index: number;
      count: number;
      expandTableIds?: readonly string[];
    }
  | { type: "columns.delete"; sheetId: string; index: number; count: number };

export type WorksheetStructureServices = Readonly<{ createId: IdGenerator }>;

export type WorksheetStructureReduction =
  | { status: "committed"; workbook: WorkbookModel }
  | {
      status: "rejected";
      reason: "validation" | "permission";
      workbook: WorkbookModel;
      issues: readonly TableIssue[];
    };

type WorksheetStructureIssueCode =
  | "SHEET_STRUCTURE_INDEX_INVALID"
  | "SHEET_STRUCTURE_COUNT_INVALID"
  | "SHEET_STRUCTURE_OUT_OF_BOUNDS"
  | "TABLE_PROTECTED"
  | "TABLE_PARTIAL_STRUCTURAL_EDIT"
  | "TABLE_EXPANSION_CONTEXT_INVALID"
  | "TABLE_RANGE_OVERLAP"
  | "TABLE_COLUMN_ID_CONFLICT";

type HeaderWrite = Readonly<{
  sheetId: string;
  row: number;
  column: number;
  value: string;
}>;

type StructureProjection = Readonly<{
  tables: readonly StructuredTable[];
  headerWrites: readonly HeaderWrite[];
}>;

type ColumnInsertionPlan = Readonly<{
  table: StructuredTable;
  range: CellRange;
  existingColumns: readonly StructuredTableColumn[];
  insertedSheetColumns: readonly number[];
  changed: boolean;
}>;

export function reduceWorksheetStructureCommand(
  workbook: WorkbookModel,
  command: WorksheetStructureCommand,
  services: WorksheetStructureServices
): WorksheetStructureReduction {
  const sheet = workbook.sheets.find((candidate) => candidate.id === command.sheetId);
  if (!sheet) {
    return rejected(
      workbook,
      "validation",
      "SHEET_STRUCTURE_OUT_OF_BOUNDS",
      "Worksheet does not exist"
    );
  }
  if (!Number.isInteger(command.index) || command.index < 0) {
    return rejected(
      workbook,
      "validation",
      "SHEET_STRUCTURE_INDEX_INVALID",
      "Structure index must be a non-negative integer"
    );
  }
  if (!Number.isInteger(command.count) || command.count <= 0) {
    return rejected(
      workbook,
      "validation",
      "SHEET_STRUCTURE_COUNT_INVALID",
      "Structure count must be a positive integer"
    );
  }
  const dimension = command.type.startsWith("rows.") ? sheet.rowCount : sheet.columnCount;
  const withinBounds = command.type.endsWith(".insert")
    ? command.index <= dimension
    : command.index < dimension && command.index + command.count <= dimension;
  if (!withinBounds) {
    return rejected(
      workbook,
      "validation",
      "SHEET_STRUCTURE_OUT_OF_BOUNDS",
      "Structure edit is outside the worksheet"
    );
  }
  if (sheet.protection.isProtected) {
    return rejected(
      workbook,
      "permission",
      "TABLE_PROTECTED",
      "Protected sheets cannot change worksheet structure"
    );
  }
  const projection = command.type === "columns.insert"
    ? projectColumnInsertion(workbook, command, services)
    : command.type === "columns.delete"
      ? projectColumnDeletion(workbook, command)
      : projectRowOperation(workbook, command);
  if (Array.isArray(projection)) {
    return {
      status: "rejected",
      reason: "validation",
      workbook,
      issues: projection
    };
  }
  const committedProjection = projection as StructureProjection;
  const operation = operationFor(command);
  const shifted = shiftSheetStructurePlanes(workbook, command.sheetId, operation);
  const withProjectedTables = rewriteCalculatedFormulaMetadata(
    { ...shifted, tables: [...committedProjection.tables] },
    sheet.name,
    operation
  );
  return {
    status: "committed",
    workbook: applyHeaderWrites(withProjectedTables, committedProjection.headerWrites)
  };
}

export function isWorksheetStructureCommand(command: unknown): command is WorksheetStructureCommand {
  if (!command || typeof command !== "object" || !("type" in command)) {
    return false;
  }
  const type = command.type;
  return type === "rows.insert"
    || type === "rows.delete"
    || type === "columns.insert"
    || type === "columns.delete";
}

function rejected(
  workbook: WorkbookModel,
  reason: "validation" | "permission",
  code: WorksheetStructureIssueCode,
  message: string
): WorksheetStructureReduction {
  return {
    status: "rejected",
    reason,
    workbook,
    issues: [{ code, message }]
  };
}

function operationFor(command: WorksheetStructureCommand) {
  return {
    axis: command.type.startsWith("rows.") ? "row" as const : "column" as const,
    mode: command.type.endsWith(".insert") ? "insert" as const : "delete" as const,
    index: command.index,
    count: command.count
  };
}

function projectColumnInsertion(
  workbook: WorkbookModel,
  command: Extract<WorksheetStructureCommand, { type: "columns.insert" }>,
  services: WorksheetStructureServices
): StructureProjection | readonly TableIssue[] {
  const expandTableIds = command.expandTableIds ?? [];
  const expandedIds = new Set(expandTableIds);
  if (expandedIds.size !== expandTableIds.length) {
    return [structureIssue(
      "TABLE_EXPANSION_CONTEXT_INVALID",
      "Column insertion expansion context contains duplicate table ids"
    )];
  }
  for (const tableId of expandTableIds) {
    const table = workbook.tables.find((candidate) => candidate.id === tableId);
    if (
      !table
      || table.sheetId !== command.sheetId
      || (command.index !== table.range.start.column && command.index !== table.range.end.column + 1)
    ) {
      return [structureIssue(
        "TABLE_EXPANSION_CONTEXT_INVALID",
        "Column insertion expansion context is stale, cross-sheet, or not at a table boundary"
      )];
    }
  }

  const plans = workbook.tables.map((table) => planColumnInsertion(table, command, expandedIds));
  if (projectedTablesOverlap(plans)) {
    return [structureIssue("TABLE_RANGE_OVERLAP", "Projected structured table ranges cannot overlap")];
  }

  const reservedColumnIds = new Set(
    workbook.tables.flatMap((table) => table.columns.map((column) => column.id))
  );
  const headerWrites: HeaderWrite[] = [];
  const tables: StructuredTable[] = [];
  for (const plan of plans) {
    if (plan.insertedSheetColumns.length === 0) {
      tables.push(plan.changed
        ? { ...plan.table, range: plan.range, columns: plan.existingColumns }
        : plan.table);
      continue;
    }

    const normalizedNames = new Set(
      plan.table.columns.map((column) => normalizeStructuredTableHeader(column.name))
    );
    const insertedColumns: StructuredTableColumn[] = [];
    for (const sheetColumn of plan.insertedSheetColumns) {
      const ordinal = sheetColumn - plan.range.start.column + 1;
      const name = nextStructuredTableColumnName(ordinal, normalizedNames);
      normalizedNames.add(normalizeStructuredTableHeader(name));
      const id = services.createId("table-column");
      if (typeof id !== "string" || id.trim().length === 0 || reservedColumnIds.has(id)) {
        return [structureIssue(
          "TABLE_COLUMN_ID_CONFLICT",
          "Worksheet structure id service returned a duplicate table-column id"
        )];
      }
      reservedColumnIds.add(id);
      insertedColumns.push({ id, name, sheetColumn });
      if (plan.table.headerRow) {
        headerWrites.push({
          sheetId: plan.table.sheetId,
          row: plan.range.start.row,
          column: sheetColumn,
          value: name
        });
      }
    }
    const columns = [...plan.existingColumns, ...insertedColumns]
      .sort((left, right) => left.sheetColumn - right.sheetColumn);
    tables.push({ ...plan.table, range: plan.range, columns });
  }
  return { tables, headerWrites };
}

function projectColumnDeletion(
  workbook: WorkbookModel,
  command: Extract<WorksheetStructureCommand, { type: "columns.delete" }>
): StructureProjection | readonly TableIssue[] {
  const deleteStart = command.index;
  const deleteEnd = command.index + command.count;
  const tables: StructuredTable[] = [];

  for (const table of workbook.tables) {
    if (table.sheetId !== command.sheetId) {
      tables.push(table);
      continue;
    }
    const tableStart = table.range.start.column;
    const tableEnd = table.range.end.column + 1;
    if (deleteEnd <= tableStart) {
      tables.push({
        ...table,
        range: shiftRangeColumns(table.range, -command.count),
        columns: table.columns.map((column) => ({
          ...column,
          sheetColumn: column.sheetColumn - command.count
        }))
      });
      continue;
    }
    if (deleteStart >= tableEnd) {
      tables.push(table);
      continue;
    }

    const columns = table.columns
      .filter((column) => column.sheetColumn < deleteStart || column.sheetColumn >= deleteEnd)
      .map((column) => ({
        ...column,
        sheetColumn: column.sheetColumn >= deleteEnd
          ? column.sheetColumn - command.count
          : column.sheetColumn
      }));
    if (columns.length === 0) {
      return [structureIssue(
        "TABLE_PARTIAL_STRUCTURAL_EDIT",
        "Worksheet column deletion cannot remove every structured table column"
      )];
    }

    const survivingIds = new Set(columns.map((column) => column.id));
    const keyColumnId = table.keyColumnId && survivingIds.has(table.keyColumnId)
      ? table.keyColumnId
      : undefined;
    const sort = table.sort?.filter((entry) => survivingIds.has(entry.columnId));
    const filter = table.filter && filterUsesOnlyColumns(table.filter, survivingIds)
      ? table.filter
      : undefined;
    const startColumn = Math.min(...columns.map((column) => column.sheetColumn));
    const endColumn = Math.max(...columns.map((column) => column.sheetColumn));
    const { keyColumnId: _key, sort: _sort, filter: _filter, ...base } = table;
    tables.push({
      ...base,
      range: {
        start: { ...table.range.start, column: startColumn },
        end: { ...table.range.end, column: endColumn }
      },
      columns,
      ...(keyColumnId === undefined ? {} : { keyColumnId }),
      ...(sort === undefined ? {} : { sort }),
      ...(filter === undefined ? {} : { filter })
    });
  }

  if (tablesOverlap(tables)) {
    return [structureIssue("TABLE_RANGE_OVERLAP", "Projected structured table ranges cannot overlap")];
  }
  return { tables, headerWrites: [] };
}

function projectRowOperation(
  workbook: WorkbookModel,
  command: Extract<WorksheetStructureCommand, { type: "rows.insert" | "rows.delete" }>
): StructureProjection | readonly TableIssue[] {
  const tables: StructuredTable[] = [];
  for (const table of workbook.tables) {
    if (table.sheetId !== command.sheetId) {
      tables.push(table);
      continue;
    }

    const tableStart = table.range.start.row;
    const tableEnd = table.range.end.row;
    if (command.type === "rows.insert") {
      if (command.index <= tableStart) {
        tables.push({ ...table, range: shiftRangeRows(table.range, command.count) });
        continue;
      }
      if (command.index >= tableEnd + 1) {
        tables.push(table);
        continue;
      }
    } else {
      const deleteEnd = command.index + command.count;
      if (deleteEnd <= tableStart) {
        tables.push({ ...table, range: shiftRangeRows(table.range, -command.count) });
        continue;
      }
      if (command.index > tableEnd) {
        tables.push(table);
        continue;
      }
    }
    return [structureIssue(
      "TABLE_PARTIAL_STRUCTURAL_EDIT",
      "Worksheet row edits cannot change a structured table body"
    )];
  }

  if (tablesOverlap(tables)) {
    return [structureIssue("TABLE_RANGE_OVERLAP", "Projected structured table ranges cannot overlap")];
  }
  return { tables, headerWrites: [] };
}

function planColumnInsertion(
  table: StructuredTable,
  command: Extract<WorksheetStructureCommand, { type: "columns.insert" }>,
  expandedIds: ReadonlySet<string>
): ColumnInsertionPlan {
  if (table.sheetId !== command.sheetId) {
    return unchangedInsertionPlan(table);
  }
  const start = table.range.start.column;
  const end = table.range.end.column;
  const expandAtBoundary = expandedIds.has(table.id);

  if (command.index < start || (command.index === start && !expandAtBoundary)) {
    return {
      table,
      range: shiftRangeColumns(table.range, command.count),
      existingColumns: table.columns.map((column) => ({
        ...column,
        sheetColumn: column.sheetColumn + command.count
      })),
      insertedSheetColumns: [],
      changed: true
    };
  }

  const expandsInternally = start < command.index && command.index <= end;
  const expandsAtLeftBoundary = command.index === start && expandAtBoundary;
  const expandsAtRightBoundary = command.index === end + 1 && expandAtBoundary;
  if (expandsInternally || expandsAtLeftBoundary || expandsAtRightBoundary) {
    return {
      table,
      range: {
        start: { ...table.range.start },
        end: { ...table.range.end, column: end + command.count }
      },
      existingColumns: table.columns.map((column) => ({
        ...column,
        sheetColumn: column.sheetColumn >= command.index
          ? column.sheetColumn + command.count
          : column.sheetColumn
      })),
      insertedSheetColumns: Array.from(
        { length: command.count },
        (_, offset) => command.index + offset
      ),
      changed: true
    };
  }

  return unchangedInsertionPlan(table);
}

function unchangedInsertionPlan(table: StructuredTable): ColumnInsertionPlan {
  return {
    table,
    range: table.range,
    existingColumns: table.columns,
    insertedSheetColumns: [],
    changed: false
  };
}

function shiftRangeColumns(range: CellRange, count: number): CellRange {
  return {
    start: { ...range.start, column: range.start.column + count },
    end: { ...range.end, column: range.end.column + count }
  };
}

function shiftRangeRows(range: CellRange, count: number): CellRange {
  return {
    start: { ...range.start, row: range.start.row + count },
    end: { ...range.end, row: range.end.row + count }
  };
}

function projectedTablesOverlap(plans: readonly ColumnInsertionPlan[]): boolean {
  for (let leftIndex = 0; leftIndex < plans.length; leftIndex += 1) {
    const left = plans[leftIndex];
    for (let rightIndex = leftIndex + 1; rightIndex < plans.length; rightIndex += 1) {
      const right = plans[rightIndex];
      if (left.table.sheetId === right.table.sheetId && rangesIntersect(left.range, right.range)) {
        return true;
      }
    }
  }
  return false;
}

function tablesOverlap(tables: readonly StructuredTable[]): boolean {
  for (let leftIndex = 0; leftIndex < tables.length; leftIndex += 1) {
    const left = tables[leftIndex];
    for (let rightIndex = leftIndex + 1; rightIndex < tables.length; rightIndex += 1) {
      const right = tables[rightIndex];
      if (left.sheetId === right.sheetId && rangesIntersect(left.range, right.range)) {
        return true;
      }
    }
  }
  return false;
}

function rangesIntersect(left: CellRange, right: CellRange): boolean {
  return left.start.row <= right.end.row
    && left.end.row >= right.start.row
    && left.start.column <= right.end.column
    && left.end.column >= right.start.column;
}

function applyHeaderWrites(workbook: WorkbookModel, writes: readonly HeaderWrite[]): WorkbookModel {
  if (writes.length === 0) return workbook;
  const writesBySheet = new Map<string, HeaderWrite[]>();
  for (const write of writes) {
    const sheetWrites = writesBySheet.get(write.sheetId) ?? [];
    sheetWrites.push(write);
    writesBySheet.set(write.sheetId, sheetWrites);
  }
  return {
    ...workbook,
    sheets: workbook.sheets.map((sheet) => {
      const sheetWrites = writesBySheet.get(sheet.id);
      if (!sheetWrites) return sheet;
      const cells = { ...sheet.cells };
      for (const write of sheetWrites) {
        cells[formatCellAddress({ row: write.row, column: write.column })] = write.value;
      }
      return { ...sheet, cells };
    })
  };
}

function filterUsesOnlyColumns(
  filter: FilterExpression,
  survivingIds: ReadonlySet<string>
): boolean {
  if (filter.kind === "logical") {
    return filter.operands.every((operand) => filterUsesOnlyColumns(operand, survivingIds));
  }
  if (filter.kind === "not") {
    return filterUsesOnlyColumns(filter.operand, survivingIds);
  }
  return survivingIds.has(filter.columnId);
}

function rewriteCalculatedFormulaMetadata(
  workbook: WorkbookModel,
  editedSheetName: string,
  operation: ReturnType<typeof operationFor>
): WorkbookModel {
  let changed = false;
  const tables = workbook.tables.map((table) => {
    const owningSheet = workbook.sheets.find((sheet) => sheet.id === table.sheetId);
    if (!owningSheet) return table;
    let tableChanged = false;
    const columns = table.columns.map((column) => {
      if (column.calculatedFormula === undefined) return column;
      const calculatedFormula = rewriteFormulaForStructure(column.calculatedFormula, {
        formulaSheetName: owningSheet.name,
        editedSheetName,
        ...operation
      });
      if (calculatedFormula === column.calculatedFormula) return column;
      tableChanged = true;
      return { ...column, calculatedFormula };
    });
    if (!tableChanged) return table;
    changed = true;
    return { ...table, columns };
  });
  return changed ? { ...workbook, tables } : workbook;
}

function structureIssue(code: WorksheetStructureIssueCode, message: string): TableIssue {
  return { code, message };
}
