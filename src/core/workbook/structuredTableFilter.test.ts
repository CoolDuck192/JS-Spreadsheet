import { describe, expect, it } from "vitest";
import type { QueryScalar } from "../../table/core/query";
import { createBlankWorkbook, setCellContent } from "../../lib/workbook";
import { matchesStructuredTableFilter, isStructuredTableRowVisible } from "./structuredTableFilter";
import type { StructuredTable } from "../../types";

const values: Record<string, QueryScalar> = {
  name: { type: "string", value: "Ada Lovelace" },
  salary: { type: "number", value: 100 },
  active: { type: "boolean", value: false },
  empty: { type: "string", value: "" },
  missing: { type: "null" }
};

describe("structured table filters", () => {
  it("evaluates comparison, set, range, logical, and not expressions", () => {
    const getValue = (columnId: string) => values[columnId] ?? { type: "null" as const };
    expect(matchesStructuredTableFilter({
      kind: "comparison", columnId: "name", operator: "contains", value: { type: "string", value: "Love" }
    }, getValue)).toBe(true);
    expect(matchesStructuredTableFilter({
      kind: "set", columnId: "active", operator: "in",
      values: [{ type: "boolean", value: false }, { type: "boolean", value: true }]
    }, getValue)).toBe(true);
    expect(matchesStructuredTableFilter({
      kind: "range", columnId: "salary", operator: "between",
      lower: { type: "number", value: 90 }, upper: { type: "number", value: 110 }
    }, getValue)).toBe(true);
    expect(matchesStructuredTableFilter({
      kind: "logical", operator: "and", operands: [
        { kind: "comparison", columnId: "salary", operator: "gte", value: { type: "number", value: 100 } },
        { kind: "not", operand: { kind: "comparison", columnId: "active", operator: "eq", value: { type: "boolean", value: true } } }
      ]
    }, getValue)).toBe(true);
  });

  it("treats only null and the empty string as blank", () => {
    const getValue = (columnId: string) => values[columnId];
    expect(matchesStructuredTableFilter({ kind: "blank", columnId: "missing", operator: "isBlank" }, getValue)).toBe(true);
    expect(matchesStructuredTableFilter({ kind: "blank", columnId: "empty", operator: "isBlank" }, getValue)).toBe(true);
    expect(matchesStructuredTableFilter({ kind: "blank", columnId: "active", operator: "isBlank" }, getValue)).toBe(false);
    expect(matchesStructuredTableFilter({ kind: "blank", columnId: "salary", operator: "isBlank" }, getValue)).toBe(false);
  });

  it("intersects structured, legacy, and explicit hidden-row masks while retaining headers and totals", () => {
    let workbook = createBlankWorkbook();
    const sheetId = workbook.activeSheetId;
    workbook = setCellContent(workbook, sheetId, "A1", "Name");
    workbook = setCellContent(workbook, sheetId, "A2", "Ada");
    workbook = setCellContent(workbook, sheetId, "A3", "Grace");
    workbook = {
      ...workbook,
      sheets: [{
        ...workbook.sheets[0],
        hiddenRows: { "2": true },
        filters: [{
          id: "legacy", range: { start: { row: 0, column: 0 }, end: { row: 2, column: 0 } },
          column: 0, operator: "equals", value: "Ada", hasHeader: true
        }]
      }]
    };
    const table: StructuredTable = {
      id: "table-1", name: "People", sheetId,
      range: { start: { row: 0, column: 0 }, end: { row: 3, column: 0 } },
      headerRow: true, totalsRow: true,
      columns: [{ id: "name", name: "Name", sheetColumn: 0 }],
      rowIds: ["row-1", "row-2"],
      filter: { kind: "comparison", columnId: "name", operator: "startsWith", value: { type: "string", value: "A" } }
    };
    workbook = { ...workbook, tables: [table] };
    const evaluate = (_sheetId: string, address: string) => workbook.sheets[0].cells[address] ?? null;

    expect(isStructuredTableRowVisible(workbook, table, 0, evaluate)).toBe(true);
    expect(isStructuredTableRowVisible(workbook, table, 1, evaluate)).toBe(true);
    expect(isStructuredTableRowVisible(workbook, table, 2, evaluate)).toBe(false);
    expect(isStructuredTableRowVisible(workbook, table, 3, evaluate)).toBe(true);
    expect(workbook.sheets[0].hiddenRows).toEqual({ "2": true });
  });
});
