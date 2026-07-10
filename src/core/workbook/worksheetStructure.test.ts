import { describe, expect, it, vi } from "vitest";
import type { StructuredTable } from "../../types";
import { createBlankWorkbook, getCellContent, setCellContent, setSheetProtection } from "../../lib/workbook";
import { isWorksheetStructureCommand, reduceWorksheetStructureCommand } from "./worksheetStructure";

const services = { createId: vi.fn((kind: string) => `${kind}-generated`) };

describe("worksheet structure reducer", () => {
  it.each([
    { type: "columns.insert", index: 27, count: 1 },
    { type: "columns.delete", index: 25, count: 2 },
    { type: "rows.insert", index: -1, count: 1 },
    { type: "rows.delete", index: 0, count: 0 }
  ] as const)("rejects strict structural bounds: $type", (command) => {
    const workbook = createBlankWorkbook();
    const result = reduceWorksheetStructureCommand(workbook, {
      ...command,
      sheetId: workbook.activeSheetId
    }, services);
    expect(result).toMatchObject({
      status: "rejected",
      reason: "validation",
      workbook,
      issues: [{ code: expect.stringMatching(/^SHEET_STRUCTURE_/) }]
    });
    expect(services.createId).not.toHaveBeenCalled();
  });

  it("rejects protected sheets without mutation", () => {
    let workbook = createBlankWorkbook();
    workbook = setSheetProtection(workbook, workbook.activeSheetId, true);
    const result = reduceWorksheetStructureCommand(workbook, {
      type: "columns.insert",
      sheetId: workbook.activeSheetId,
      index: 0,
      count: 1
    }, services);
    expect(result).toEqual({
      status: "rejected",
      reason: "permission",
      workbook,
      issues: [{ code: "TABLE_PROTECTED", message: "Protected sheets cannot change worksheet structure" }]
    });
  });

  it("commits a valid table-free plane shift", () => {
    let workbook = createBlankWorkbook();
    const sheetId = workbook.activeSheetId;
    workbook = setCellContent(workbook, sheetId, "A1", "Name");
    const result = reduceWorksheetStructureCommand(workbook, {
      type: "columns.insert", sheetId, index: 0, count: 1
    }, services);
    expect(result.status).toBe("committed");
    if (result.status !== "committed") return;
    expect(getCellContent(result.workbook, sheetId, "B1")).toBe("Name");
  });

  it("rejects table-owned sheets until table-aware editing is implemented", () => {
    let workbook = createBlankWorkbook();
    const sheetId = workbook.activeSheetId;
    const table: StructuredTable = {
      id: "table-1",
      name: "TableOne",
      sheetId,
      range: { start: { row: 0, column: 0 }, end: { row: 1, column: 0 } },
      headerRow: true,
      totalsRow: false,
      columns: [{ id: "table-column-1", name: "Name", sheetColumn: 0 }],
      rowIds: ["table-row-1"]
    };
    workbook = { ...workbook, tables: [table] };

    const result = reduceWorksheetStructureCommand(workbook, {
      type: "rows.insert",
      sheetId,
      index: 0,
      count: 1
    }, services);

    expect(result).toEqual({
      status: "rejected",
      reason: "validation",
      workbook,
      issues: [{
        code: "TABLE_PARTIAL_STRUCTURAL_EDIT",
        message: "Structured tables require table-aware structure editing"
      }]
    });
    expect(services.createId).not.toHaveBeenCalled();
  });

  it("recognizes only worksheet structure command types", () => {
    expect(isWorksheetStructureCommand({
      type: "columns.insert",
      sheetId: "sheet-1",
      index: 0,
      count: 1,
      expandTableIds: ["table-1"]
    })).toBe(true);
    expect(isWorksheetStructureCommand({ type: "columns.resize" })).toBe(false);
  });
});
