import { describe, expect, it } from "vitest";
import type { CellRange, WorkbookModel } from "../../types";
import { createBlankWorkbook, deleteSheet, duplicateSheet, getCellContent, setCellContent } from "../../lib/workbook";
import {
  getStructuredTable,
  getStructuredTableAtCell,
  getStructuredTableBodyRange,
  getStructuredTableForSelection,
  reduceStructuredTableCommand,
  type StructuredTableCommand,
  type StructuredTableCommandServices
} from "./structuredTables";

describe("structured table metadata", () => {
  it("creates an A1:C4 table with stable lookup and body geometry", () => {
    const { workbook, services } = tableFixture(3, 4);
    const result = reduceStructuredTableCommand(workbook, {
      type: "table.create",
      sheetId: workbook.activeSheetId,
      range: range(0, 0, 3, 2),
      name: "Sales",
      headerRow: true,
      totalsRow: false
    }, services);
    expect(result.status).toBe("committed");
    const next = result.workbook;
    const table = next.tables[0];
    expect(table).toMatchObject({
      name: "Sales",
      range: range(0, 0, 3, 2),
      headerRow: true,
      columns: [{ name: "Column 1" }, { name: "Column 2" }, { name: "Column 3" }]
    });
    expect(table.rowIds).toHaveLength(3);
    expect(getStructuredTable(next, table.id)).toBe(table);
    expect(getStructuredTableAtCell(next, table.sheetId, { row: 2, column: 1 })).toBe(table);
    expect(getStructuredTableForSelection(next, table.sheetId, range(5, 5, 2, 1))).toBe(table);
    expect(getStructuredTableBodyRange(table)).toEqual(range(1, 0, 3, 2));
  });

  it.each(["", " Sales", "R", "A1", "XFD1048576", "R1C1", "x".repeat(256)])(
    "rejects invalid table name %s atomically",
    (name) => {
      const { workbook, services } = tableFixture(2, 3);
      expectRejectedUnchanged(workbook, {
        type: "table.create",
        sheetId: workbook.activeSheetId,
        range: range(0, 0, 2, 1),
        name,
        headerRow: true,
        totalsRow: false
      }, services, "TABLE_NAME_INVALID");
    }
  );

  it("uses the same validator for rename and rejects normalized name conflicts", () => {
    const { workbook, services } = tableFixture(4, 6);
    const first = commit(workbook, {
      type: "table.create", sheetId: workbook.activeSheetId, range: range(0, 0, 2, 1),
      name: "Sales", headerRow: true, totalsRow: false
    }, services);
    const second = commit(first, {
      type: "table.create", sheetId: workbook.activeSheetId, range: range(0, 2, 2, 3),
      name: "Costs", headerRow: true, totalsRow: false
    }, services);
    expectRejectedUnchanged(second, {
      type: "table.rename", tableId: second.tables[1].id, name: "ＳＡＬＥＳ"
    }, services, "TABLE_NAME_CONFLICT");
    expectRejectedUnchanged(second, {
      type: "table.rename", tableId: second.tables[1].id, name: "R1C1"
    }, services, "TABLE_NAME_INVALID");
  });

  it("rejects blank or duplicate headers", () => {
    const { workbook, services } = tableFixture(3, 3);
    const blank = setCellContent(workbook, workbook.activeSheetId, "A1", "");
    expectRejectedUnchanged(blank, {
      type: "table.create", sheetId: blank.activeSheetId, range: range(0, 0, 2, 1),
      name: "BlankHeaders", headerRow: true, totalsRow: false
    }, services, "TABLE_HEADER_INVALID");
    let duplicate = setCellContent(workbook, workbook.activeSheetId, "A1", "Same");
    duplicate = setCellContent(duplicate, duplicate.activeSheetId, "B1", "same");
    expectRejectedUnchanged(duplicate, {
      type: "table.create", sheetId: duplicate.activeSheetId, range: range(0, 0, 2, 1),
      name: "DuplicateHeaders", headerRow: true, totalsRow: false
    }, services, "TABLE_HEADER_INVALID");
  });

  it("rejects overlap, merges, protection, and impossible ranges", () => {
    const { workbook, services } = tableFixture(4, 5);
    const first = commit(workbook, {
      type: "table.create", sheetId: workbook.activeSheetId, range: range(0, 0, 2, 1),
      name: "First", headerRow: true, totalsRow: false
    }, services);
    expectRejectedUnchanged(first, {
      type: "table.create", sheetId: first.activeSheetId, range: range(1, 1, 3, 2),
      name: "Overlap", headerRow: false, totalsRow: false
    }, services, "TABLE_RANGE_OVERLAP");

    const withMerge: WorkbookModel = {
      ...workbook,
      sheets: [{ ...workbook.sheets[0], merges: [{ id: "merge-1", range: range(0, 0, 1, 1) }] }]
    };
    expectRejectedUnchanged(withMerge, {
      type: "table.create", sheetId: withMerge.activeSheetId, range: range(0, 0, 2, 1),
      name: "Merged", headerRow: true, totalsRow: false
    }, services, "TABLE_MERGE_CONFLICT");

    const protectedWorkbook: WorkbookModel = {
      ...workbook,
      sheets: [{
        ...workbook.sheets[0],
        protection: { isProtected: true, lockedCells: {}, unlockedCells: {} }
      }]
    };
    expectRejectedUnchanged(protectedWorkbook, {
      type: "table.create", sheetId: protectedWorkbook.activeSheetId, range: range(0, 0, 2, 1),
      name: "Protected", headerRow: true, totalsRow: false
    }, services, "TABLE_PROTECTED");

    for (const invalidRange of [range(-1, 0, 2, 1), range(0, 0, 5_000, 1), range(2, 1, 1, 1)]) {
      expectRejectedUnchanged(workbook, {
        type: "table.create", sheetId: workbook.activeSheetId, range: invalidRange,
        name: "Blocked", headerRow: false, totalsRow: false
      }, services, "TABLE_RANGE_BLOCKED");
    }
  });

  it("resizes while preserving surviving IDs and absorbing populated cells", () => {
    const { workbook, services } = tableFixture(4, 5);
    const created = commit(workbook, {
      type: "table.create", sheetId: workbook.activeSheetId, range: range(0, 0, 2, 1),
      name: "Resizable", headerRow: true, totalsRow: false
    }, services);
    const initial = created.tables[0];
    const expanded = commit(created, {
      type: "table.resize", tableId: initial.id, range: range(0, 0, 3, 2)
    }, services);
    const table = expanded.tables[0];
    expect(table.columns.slice(0, 2).map((column) => column.id)).toEqual(initial.columns.map((column) => column.id));
    expect(table.columns).toHaveLength(3);
    expect(table.rowIds.slice(0, 2)).toEqual(initial.rowIds);
    expect(table.rowIds).toHaveLength(3);
    expectRejectedUnchanged(expanded, {
      type: "table.resize", tableId: table.id, range: range(1, 0, 3, 2)
    }, services, "TABLE_RANGE_BLOCKED");
  });

  it("toggles headers and totals without replacing body row IDs", () => {
    let workbook = createBlankWorkbook();
    workbook = setCellContent(workbook, workbook.activeSheetId, "A1", "Ada");
    workbook = setCellContent(workbook, workbook.activeSheetId, "A2", "Grace");
    const services = deterministicServices();
    const created = commit(workbook, {
      type: "table.create", sheetId: workbook.activeSheetId, range: range(0, 0, 1, 0),
      name: "People", headerRow: false, totalsRow: false
    }, services);
    const tableId = created.tables[0].id;
    const rowIds = created.tables[0].rowIds;
    const withHeader = commit(created, { type: "table.setHeaderRow", tableId, enabled: true }, services);
    expect(withHeader.tables[0].rowIds).toEqual(rowIds);
    expect(getCellContent(withHeader, workbook.activeSheetId, "A1")).toBe("Column1");
    expect(getCellContent(withHeader, workbook.activeSheetId, "A2")).toBe("Ada");
    const withTotals = commit(withHeader, { type: "table.setTotalsRow", tableId, enabled: true }, services);
    expect(withTotals.tables[0].totalsRow).toBe(true);
    expect(withTotals.tables[0].rowIds).toEqual(rowIds);
    const restored = commit(withTotals, { type: "table.setTotalsRow", tableId, enabled: false }, services);
    expect(restored.tables[0].totalsRow).toBe(false);
    expect(restored.tables[0].rowIds).toEqual(rowIds);
  });

  it("renames columns, applies metadata, and converts to a formatted range", () => {
    const { workbook, services } = tableFixture(3, 3);
    const created = commit(workbook, {
      type: "table.create", sheetId: workbook.activeSheetId, range: range(0, 0, 2, 1),
      name: "Styled", headerRow: true, totalsRow: false
    }, services);
    const table = created.tables[0];
    const renamed = commit(created, {
      type: "table.renameColumn", tableId: table.id, columnId: table.columns[0].id, name: "Account"
    }, services);
    expect(getCellContent(renamed, workbook.activeSheetId, "A1")).toBe("Account");
    const keyed = commit(renamed, {
      type: "table.setKeyColumn", tableId: table.id, columnId: table.columns[0].id
    }, services);
    const styled = commit(keyed, {
      type: "table.setStyle", tableId: table.id,
      style: { theme: "TableStyleLight2", showRowStripes: true }
    }, services);
    const converted = commit(styled, { type: "table.convertToRange", tableId: table.id }, services);
    expect(converted.tables).toEqual([]);
    expect(converted.sheets[0].formats.A1).toMatchObject({ bold: true });
    expect(getCellContent(converted, workbook.activeSheetId, "A2")).not.toBeNull();
  });

  it("duplicates and deletes sheet-owned tables with fresh IDs", () => {
    const { workbook, services } = tableFixture(3, 3);
    const created = commit(workbook, {
      type: "table.create", sheetId: workbook.activeSheetId, range: range(0, 0, 2, 1),
      name: "Source", headerRow: true, totalsRow: false
    }, services);
    const duplicated = duplicateSheet(created, created.activeSheetId, services.createId);
    expect(duplicated.tables).toHaveLength(2);
    expect(duplicated.tables[1].sheetId).toBe(duplicated.activeSheetId);
    expect(duplicated.tables[1].id).not.toBe(duplicated.tables[0].id);
    expect(duplicated.tables[1].columns[0].id).not.toBe(duplicated.tables[0].columns[0].id);
    const deleted = deleteSheet(duplicated, created.activeSheetId);
    expect(deleted.tables).toHaveLength(1);
    expect(deleted.tables[0].sheetId).toBe(duplicated.activeSheetId);
  });
});

function tableFixture(columnCount: number, rowCount: number) {
  let workbook = createBlankWorkbook();
  for (let column = 0; column < columnCount; column += 1) {
    const letter = String.fromCharCode(65 + column);
    workbook = setCellContent(workbook, workbook.activeSheetId, `${letter}1`, `Column ${column + 1}`);
    for (let row = 2; row <= rowCount; row += 1) {
      workbook = setCellContent(workbook, workbook.activeSheetId, `${letter}${row}`, `${letter}-${row}`);
    }
  }
  return { workbook, services: deterministicServices() };
}

function deterministicServices(): StructuredTableCommandServices {
  let next = 0;
  return {
    createId(kind) {
      next += 1;
      return `${kind}-${next}`;
    },
    getCellEvaluation() {
      return null;
    }
  };
}

function commit(
  workbook: WorkbookModel,
  command: StructuredTableCommand,
  services: StructuredTableCommandServices
): WorkbookModel {
  const result = reduceStructuredTableCommand(workbook, command, services);
  expect(result.status).toBe("committed");
  return result.workbook;
}

function expectRejectedUnchanged(
  workbook: WorkbookModel,
  command: StructuredTableCommand,
  services: StructuredTableCommandServices,
  code: string
): void {
  const before = JSON.stringify(workbook);
  const result = reduceStructuredTableCommand(workbook, command, services);
  expect(result).toMatchObject({ status: "rejected", workbook, issues: [{ code }] });
  expect(result.workbook).toBe(workbook);
  expect(JSON.stringify(workbook)).toBe(before);
}

function range(startRow: number, startColumn: number, endRow: number, endColumn: number): CellRange {
  return {
    start: { row: startRow, column: startColumn },
    end: { row: endRow, column: endColumn }
  };
}
