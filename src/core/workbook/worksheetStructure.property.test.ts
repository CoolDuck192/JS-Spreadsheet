import fc from "fast-check";
import { describe, expect, it } from "vitest";
import type { CellRange, StructuredTable, WorkbookModel } from "../../types";
import { formatCellAddress } from "../../lib/addressing";
import { createBlankWorkbook } from "../../lib/workbook";
import type { FilterExpression } from "../../table/core/query";
import { migrateWorkbookModel } from "./migrateWorkbook";
import { reduceWorksheetStructureCommand } from "./worksheetStructure";

describe("worksheet structure properties", () => {
  it("preserves migrated workbook invariants for bounded multi-table column edits", () => {
    fc.assert(fc.property(
      fc.record({
        gap: fc.integer({ min: 1, max: 5 }),
        count: fc.integer({ min: 1, max: 3 }),
        direction: fc.constantFrom("before", "inside", "after")
      }),
      ({ gap, count, direction }) => {
        const workbook = generatedTwoTableWorkbook(gap);
        expect(migrateWorkbookModel(workbook)).not.toBeNull();
        assertTableStructureInvariants(workbook);
        const first = workbook.tables[0];
        const index = direction === "before"
          ? first.range.start.column
          : direction === "inside"
            ? first.range.end.column
            : first.range.end.column + 1;
        const expandTableIds = direction === "after" ? [first.id] : undefined;
        const result = reduceWorksheetStructureCommand(workbook, {
          type: "columns.insert",
          sheetId: workbook.activeSheetId,
          index,
          count,
          expandTableIds
        }, sequentialServices());

        expect(result.status).toBe("committed");
        if (result.status !== "committed") return;
        expect(migrateWorkbookModel(result.workbook)).not.toBeNull();
        assertTableStructureInvariants(result.workbook);
        assertSurvivingMetadata(workbook, result.workbook);
      }
    ), { numRuns: 100, seed: 20260710 });
  });
});

function generatedTwoTableWorkbook(gap: number): WorkbookModel {
  const workbook = createBlankWorkbook();
  const sheetId = workbook.activeSheetId;
  const firstEnd = 2;
  const secondStart = firstEnd + gap + 1;
  const tables: StructuredTable[] = [
    {
      id: "table-generated-sales",
      name: "GeneratedSales",
      sheetId,
      range: range(0, 0, 3, firstEnd),
      headerRow: true,
      totalsRow: false,
      columns: [
        { id: "sales-region", name: "Region", sheetColumn: 0, dataType: "text" },
        { id: "sales-amount", name: "Amount", sheetColumn: 1, dataType: "number" },
        { id: "sales-variance", name: "Variance", sheetColumn: 2, dataType: "number" }
      ],
      rowIds: ["sales-row-1", "sales-row-2", "sales-row-3"],
      keyColumnId: "sales-region",
      style: { theme: "TableStyleMedium2", showFirstColumn: true, showRowStripes: true },
      sort: [{ columnId: "sales-amount", direction: "desc", nulls: "last" }],
      filter: {
        kind: "comparison",
        columnId: "sales-variance",
        operator: "gt",
        value: { type: "number", value: 0 }
      }
    },
    {
      id: "table-generated-stock",
      name: "GeneratedStock",
      sheetId,
      range: range(0, secondStart, 3, secondStart + 1),
      headerRow: true,
      totalsRow: false,
      columns: [
        { id: "stock-product", name: "Product", sheetColumn: secondStart, dataType: "text" },
        { id: "stock-units", name: "Units", sheetColumn: secondStart + 1, dataType: "number" }
      ],
      rowIds: ["stock-row-1", "stock-row-2", "stock-row-3"],
      keyColumnId: "stock-product",
      style: { theme: "TableStyleLight1", showColumnStripes: true },
      sort: [{ columnId: "stock-units", direction: "asc" }],
      filter: { kind: "blank", columnId: "stock-units", operator: "isNotBlank" }
    }
  ];
  const cells = { ...workbook.sheets[0].cells };
  for (const table of tables) {
    for (const column of table.columns) {
      cells[formatCellAddress({ row: table.range.start.row, column: column.sheetColumn })] = column.name;
    }
  }
  return {
    ...workbook,
    sheets: [{ ...workbook.sheets[0], cells }],
    tables
  };
}

function assertTableStructureInvariants(workbook: WorkbookModel): void {
  const tableIds = new Set<string>();

  for (const table of workbook.tables) {
    expect(table.id.trim()).not.toBe("");
    expect(tableIds.has(table.id)).toBe(false);
    tableIds.add(table.id);

    const sheet = workbook.sheets.find((candidate) => candidate.id === table.sheetId);
    expect(sheet).toBeDefined();
    if (!sheet) continue;

    expect(table.range.start.row).toBeGreaterThanOrEqual(0);
    expect(table.range.start.column).toBeGreaterThanOrEqual(0);
    expect(table.range.end.row).toBeGreaterThanOrEqual(table.range.start.row);
    expect(table.range.end.column).toBeGreaterThanOrEqual(table.range.start.column);
    expect(table.range.end.row).toBeLessThan(sheet.rowCount);
    expect(table.range.end.column).toBeLessThan(sheet.columnCount);

    const width = table.range.end.column - table.range.start.column + 1;
    expect(table.columns).toHaveLength(width);
    expect(table.columns.map((column) => column.sheetColumn)).toEqual(
      Array.from({ length: width }, (_, offset) => table.range.start.column + offset)
    );

    const normalizedHeaders = table.columns.map((column) => {
      expect(column.name.trim()).not.toBe("");
      return column.name.normalize("NFKC").toLowerCase();
    });
    expect(new Set(normalizedHeaders).size).toBe(normalizedHeaders.length);

    const columnIds = new Set<string>();
    for (const column of table.columns) {
      expect(column.id.trim()).not.toBe("");
      expect(columnIds.has(column.id)).toBe(false);
      columnIds.add(column.id);
    }

    const bodyHeight = table.range.end.row - table.range.start.row + 1
      - Number(table.headerRow)
      - Number(table.totalsRow);
    expect(table.rowIds).toHaveLength(bodyHeight);
    expect(table.rowIds.every((rowId) => rowId.trim().length > 0)).toBe(true);
    expect(new Set(table.rowIds).size).toBe(table.rowIds.length);

    if (table.keyColumnId !== undefined) {
      expect(columnIds.has(table.keyColumnId)).toBe(true);
    }
    for (const sort of table.sort ?? []) {
      expect(columnIds.has(sort.columnId)).toBe(true);
    }
    for (const filterColumnId of filterColumnIds(table.filter)) {
      expect(columnIds.has(filterColumnId)).toBe(true);
    }
  }

  expect(tableIds.size).toBe(workbook.tables.length);
  for (let leftIndex = 0; leftIndex < workbook.tables.length; leftIndex += 1) {
    for (let rightIndex = leftIndex + 1; rightIndex < workbook.tables.length; rightIndex += 1) {
      const left = workbook.tables[leftIndex];
      const right = workbook.tables[rightIndex];
      if (left.sheetId === right.sheetId) {
        expect(rangesIntersect(left.range, right.range)).toBe(false);
      }
    }
  }
}

function assertSurvivingMetadata(before: WorkbookModel, after: WorkbookModel): void {
  expect(after.tables.map((table) => table.id)).toEqual(before.tables.map((table) => table.id));
  for (const previousTable of before.tables) {
    const nextTable = after.tables.find((table) => table.id === previousTable.id);
    expect(nextTable).toBeDefined();
    if (!nextTable) continue;
    expect({
      id: nextTable.id,
      name: nextTable.name,
      sheetId: nextTable.sheetId,
      headerRow: nextTable.headerRow,
      totalsRow: nextTable.totalsRow,
      rowIds: nextTable.rowIds,
      keyColumnId: nextTable.keyColumnId,
      style: nextTable.style,
      sort: nextTable.sort,
      filter: nextTable.filter
    }).toEqual({
      id: previousTable.id,
      name: previousTable.name,
      sheetId: previousTable.sheetId,
      headerRow: previousTable.headerRow,
      totalsRow: previousTable.totalsRow,
      rowIds: previousTable.rowIds,
      keyColumnId: previousTable.keyColumnId,
      style: previousTable.style,
      sort: previousTable.sort,
      filter: previousTable.filter
    });
    for (const previousColumn of previousTable.columns) {
      const nextColumn = nextTable.columns.find((column) => column.id === previousColumn.id);
      expect(nextColumn).toBeDefined();
      if (!nextColumn) continue;
      const { sheetColumn: _previousSheetColumn, ...previousMetadata } = previousColumn;
      const { sheetColumn: _nextSheetColumn, ...nextMetadata } = nextColumn;
      expect(nextMetadata).toEqual(previousMetadata);
    }
  }
}

function filterColumnIds(filter: FilterExpression | undefined): readonly string[] {
  if (!filter) return [];
  if (filter.kind === "logical") return filter.operands.flatMap(filterColumnIds);
  if (filter.kind === "not") return filterColumnIds(filter.operand);
  return [filter.columnId];
}

function rangesIntersect(left: CellRange, right: CellRange): boolean {
  return left.start.row <= right.end.row
    && left.end.row >= right.start.row
    && left.start.column <= right.end.column
    && left.end.column >= right.start.column;
}

function range(startRow: number, startColumn: number, endRow: number, endColumn: number): CellRange {
  return {
    start: { row: startRow, column: startColumn },
    end: { row: endRow, column: endColumn }
  };
}

function sequentialServices() {
  let sequence = 0;
  return {
    createId(kind: string) {
      sequence += 1;
      return `${kind}-generated-${sequence}`;
    }
  };
}
