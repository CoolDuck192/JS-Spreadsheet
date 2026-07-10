import { describe, expect, expectTypeOf, it } from "vitest";
import {
  addConditionalFormatRule,
  addSheet,
  createBlankWorkbook,
  getCellConditionalFormatRules,
  getCellContent,
  getCellFormat,
  getColumnWidth,
  setActiveSheet,
  setCellContent,
  setCellFormat,
  setCellValidation
} from "../../lib/workbook";
import type { CommandEnvelope, CommandResult } from "../commands/types";
import { applyWorkbookMutation, type WorkbookCommand } from "./commands";

const cell = { row: 0, column: 0 } as const;
const range = { start: cell, end: { row: 1, column: 1 } } as const;
const clipboard = {
  range,
  cells: [[{
    sourceAddress: "A1",
    content: "copied",
    displayContent: "copied",
    format: { bold: true },
    validation: null,
    comment: "note",
    hyperlink: "https://example.test"
  }]]
} as const;

const completePreTableCommands = [
  { type: "transaction", commands: [{ type: "selection.set", selection: range }] },
  { type: "selection.set", selection: range },
  { type: "cell.set", sheetId: "sheet-1", address: "A1", input: "42" },
  { type: "cell.comment.set", sheetId: "sheet-1", address: "A1", comment: "note" },
  { type: "cell.hyperlink.set", sheetId: "sheet-1", address: "A1", hyperlink: null },
  ...(["contents", "formats", "comments", "hyperlinks", "all"] as const).map((mode) => ({
    type: "range.clear" as const,
    sheetId: "sheet-1",
    range,
    mode
  })),
  { type: "range.format", sheetId: "sheet-1", range, format: { bold: true } },
  { type: "range.directFormat.clear", sheetId: "sheet-1", range },
  { type: "range.format.replace", sheetId: "sheet-1", range, format: { italic: true } },
  { type: "range.borders", sheetId: "sheet-1", range, preset: "outer" },
  { type: "range.validation.set", sheetId: "sheet-1", range, rule: { type: "number", min: 1 } },
  { type: "range.validation.clear", sheetId: "sheet-1", range },
  {
    type: "range.conditionalFormat.add",
    sheetId: "sheet-1",
    range,
    rule: {
      id: "conditional-1",
      range,
      condition: { type: "greaterThan", value: "1" },
      format: { backgroundColor: "#ffffff" }
    }
  },
  { type: "range.conditionalFormat.remove", sheetId: "sheet-1", ruleId: "conditional-1" },
  { type: "range.conditionalFormat.clear", sheetId: "sheet-1", range },
  { type: "range.readOnly.set", sheetId: "sheet-1", range, readOnly: true },
  { type: "range.merge", sheetId: "sheet-1", range },
  { type: "range.unmerge", sheetId: "sheet-1", range },
  { type: "range.fill", sheetId: "sheet-1", range, direction: "down" },
  { type: "range.fill", sheetId: "sheet-1", range, direction: "right" },
  { type: "range.autoFill", sheetId: "sheet-1", source: range, target: range },
  { type: "range.sort", sheetId: "sheet-1", range, direction: "asc", sortColumn: 0 },
  { type: "range.removeDuplicates", sheetId: "sheet-1", range },
  { type: "clipboard.paste", sheetId: "sheet-1", target: cell, payload: clipboard, mode: "all" },
  { type: "clipboard.pasteMatrix", sheetId: "sheet-1", target: cell, matrix: [["plain", "=A1"]] },
  {
    type: "clipboard.move",
    sourceSheetId: "sheet-1",
    source: range,
    targetSheetId: "sheet-2",
    target: cell
  },
  { type: "rows.insert", sheetId: "sheet-1", index: 0, count: 2 },
  { type: "rows.delete", sheetId: "sheet-1", index: 0, count: 2 },
  { type: "columns.insert", sheetId: "sheet-1", index: 0, count: 2 },
  { type: "columns.delete", sheetId: "sheet-1", index: 0, count: 2 },
  { type: "rows.resize", sheetId: "sheet-1", rows: [0, 2], height: 32 },
  { type: "columns.resize", sheetId: "sheet-1", columns: [0, 2], width: 120 },
  { type: "rows.hidden.set", sheetId: "sheet-1", rows: [0, 2], hidden: true },
  { type: "columns.hidden.set", sheetId: "sheet-1", columns: [0, 2], hidden: true },
  { type: "sheet.add", name: "Second" },
  {
    type: "sheet.createFromMatrix",
    sheetId: "sheet-2",
    name: "Generated",
    rows: [["Heading"], ["Value"]],
    formats: [{ range: { start: cell, end: cell }, format: { bold: true } }],
    columnWidths: [120],
    freeze: { rows: 1, columns: 0 }
  },
  { type: "sheet.replaceWithRows", sheetId: "sheet-1", rows: [["A", "B"]] },
  { type: "sheet.rename", sheetId: "sheet-1", name: "Renamed" },
  { type: "sheet.duplicate", sheetId: "sheet-1" },
  { type: "sheet.delete", sheetId: "sheet-1" },
  { type: "sheet.activate", sheetId: "sheet-1" },
  { type: "sheet.move", sheetId: "sheet-1", targetIndex: 1 },
  { type: "sheet.hidden.set", sheetId: "sheet-1", hidden: true },
  { type: "sheet.tabColor.set", sheetId: "sheet-1", color: "#123456" },
  { type: "sheet.freeze.set", sheetId: "sheet-1", rows: 1, columns: 1 },
  { type: "sheet.protection.set", sheetId: "sheet-1", protected: true },
  {
    type: "sheet.filter.set",
    sheetId: "sheet-1",
    filter: {
      id: "filter-1",
      range,
      column: 0,
      operator: "equals",
      value: "Open"
    }
  },
  { type: "sheet.filter.clear", sheetId: "sheet-1", column: 0 },
  {
    type: "sheet.chart.add",
    sheetId: "sheet-1",
    chart: { id: "chart-1", title: "Chart", type: "bar", range, anchor: cell }
  },
  { type: "sheet.chart.delete", sheetId: "sheet-1", chartId: "chart-1" },
  { type: "namedRange.define", namedRange: { name: "Sales", sheetId: "sheet-1", range } },
  { type: "namedRange.remove", name: "Sales" },
  { type: "history.undo" },
  { type: "history.redo" },
  { type: "persistence.status", status: "failed", message: "storage unavailable" },
  { type: "workbook.replace", workbook: createBlankWorkbook(), history: "reset" },
  { type: "table.create", sheetId: "sheet-1", range, name: "TableOne", headerRow: true, totalsRow: false },
  { type: "table.rename", tableId: "table-1", name: "RenamedTable" },
  { type: "table.renameColumn", tableId: "table-1", columnId: "column-1", name: "Renamed" },
  { type: "table.resize", tableId: "table-1", range },
  { type: "table.setHeaderRow", tableId: "table-1", enabled: true },
  { type: "table.setTotalsRow", tableId: "table-1", enabled: true },
  { type: "table.setTotalsFunction", tableId: "table-1", columnId: "column-1", aggregate: "sum" },
  { type: "table.setStyle", tableId: "table-1", style: { theme: "TableStyleLight1" } },
  { type: "table.setKeyColumn", tableId: "table-1", columnId: "column-1" },
  { type: "table.setCalculatedColumn", tableId: "table-1", columnId: "column-1", formula: "=A2" },
  { type: "table.setFilter", tableId: "table-1", filter: { kind: "blank", columnId: "column-1", operator: "isBlank" } },
  { type: "table.sort", tableId: "table-1", sorting: [{ columnId: "column-1", direction: "asc" }] },
  { type: "table.insertRows", tableId: "table-1", count: 2, beforeRowId: "row-1" },
  { type: "table.deleteRows", tableId: "table-1", rowIds: ["row-1"] },
  { type: "table.editCells", tableId: "table-1", edits: [{ rowId: "row-1", columnId: "column-1", rawText: "value" }] },
  { type: "table.convertToRange", tableId: "table-1" }
] satisfies readonly WorkbookCommand[];

describe("WorkbookCommand", () => {
  it("models the complete pre-table command set as serializable data", () => {
    expect(JSON.parse(JSON.stringify(completePreTableCommands))).toHaveLength(completePreTableCommands.length);
    expect(containsFunction(completePreTableCommands)).toBe(false);
  });

  it("uses shared serializable command lifecycle types", () => {
    const envelope: CommandEnvelope<WorkbookCommand> = {
      id: "command-1",
      intent: { type: "history.undo" },
      transactionId: "transaction-1",
      expectedRevision: "4"
    };
    const result: CommandResult = { status: "committed", revision: "5", changed: true };

    expectTypeOf(envelope.intent).toMatchTypeOf<WorkbookCommand>();
    expectTypeOf(result).toMatchTypeOf<CommandResult>();
    expect(JSON.parse(JSON.stringify(envelope))).toEqual(envelope);
  });

  it("clears and replaces direct formats without removing conditional formats", () => {
    let workbook = createBlankWorkbook();
    workbook = setCellFormat(workbook, "sheet-1", range, { bold: true, backgroundColor: "#ffffff" });
    workbook = addConditionalFormatRule(workbook, "sheet-1", range, {
      condition: { type: "greaterThan", value: "1" },
      format: { textColor: "#ff0000" }
    });

    const cleared = apply(workbook, { type: "range.directFormat.clear", sheetId: "sheet-1", range });
    expect(cleared.status).toBe("applied");
    if (cleared.status !== "applied") return;
    expect(getCellFormat(cleared.workbook, "sheet-1", "A1")).toEqual({});
    expect(getCellConditionalFormatRules(cleared.workbook, "sheet-1", "A1")).toHaveLength(1);

    const replaced = apply(workbook, {
      type: "range.format.replace",
      sheetId: "sheet-1",
      range,
      format: { italic: true }
    });
    expect(replaced.status).toBe("applied");
    if (replaced.status !== "applied") return;
    expect(getCellFormat(replaced.workbook, "sheet-1", "A1")).toEqual({ italic: true });
    expect(getCellConditionalFormatRules(replaced.workbook, "sheet-1", "A1")).toHaveLength(1);
  });

  it("clears only conditional-format rules intersecting the requested range", () => {
    let workbook = createBlankWorkbook();
    workbook = addConditionalFormatRule(workbook, "sheet-1", {
      start: cell,
      end: cell
    }, {
      condition: { type: "blank" },
      format: { bold: true }
    });
    workbook = addConditionalFormatRule(workbook, "sheet-1", {
      start: { row: 0, column: 1 },
      end: { row: 0, column: 1 }
    }, {
      condition: { type: "blank" },
      format: { italic: true }
    });

    const result = apply(workbook, {
      type: "range.conditionalFormat.clear",
      sheetId: "sheet-1",
      range: { start: cell, end: cell }
    });
    expect(result.status).toBe("applied");
    if (result.status !== "applied") return;
    expect(getCellConditionalFormatRules(result.workbook, "sheet-1", "A1")).toEqual([]);
    expect(getCellConditionalFormatRules(result.workbook, "sheet-1", "B1")).toHaveLength(1);
  });

  it("pastes a plain matrix atomically without translating formula text", () => {
    let workbook = createBlankWorkbook();
    workbook = setCellValidation(workbook, "sheet-1", {
      start: { row: 0, column: 1 },
      end: { row: 0, column: 1 }
    }, { type: "number", max: 5 });

    const rejected = apply(workbook, {
      type: "clipboard.pasteMatrix",
      sheetId: "sheet-1",
      target: cell,
      matrix: [["kept", "10"]]
    });
    expect(rejected).toMatchObject({
      status: "rejected",
      reason: "validation",
      issues: [{ message: "Enter a number less than or equal to 5", address: "B1" }]
    });
    expect(getCellContent(workbook, "sheet-1", "A1")).toBeNull();

    const accepted = apply(workbook, {
      type: "clipboard.pasteMatrix",
      sheetId: "sheet-1",
      target: cell,
      matrix: [["=A2", "5"]]
    });
    expect(accepted.status).toBe("applied");
    if (accepted.status !== "applied") return;
    expect(getCellContent(accepted.workbook, "sheet-1", "A1")).toBe("=A2");
    expect(getCellContent(accepted.workbook, "sheet-1", "B1")).toBe("5");
  });

  it("replaces one sheet from CSV rows against the dispatch-time workbook", () => {
    let workbook = addSheet(createBlankWorkbook(), "Other");
    workbook = setActiveSheet(workbook, "sheet-1");
    workbook = setCellContent(workbook, "sheet-1", "A1", "old");
    workbook = setCellContent(workbook, "sheet-2", "A1", "keep");
    workbook = setCellFormat(workbook, "sheet-1", range, { bold: true });

    const result = apply(workbook, {
      type: "sheet.replaceWithRows",
      sheetId: "sheet-1",
      rows: [["Name", "Amount"], ["Rent", "1200"]]
    });
    expect(result.status).toBe("applied");
    if (result.status !== "applied") return;
    expect(getCellContent(result.workbook, "sheet-1", "A1")).toBe("Name");
    expect(getCellContent(result.workbook, "sheet-1", "B2")).toBe("1200");
    expect(getCellFormat(result.workbook, "sheet-1", "A1")).toEqual({});
    expect(getCellContent(result.workbook, "sheet-2", "A1")).toBe("keep");
  });

  it("creates a deterministic generated sheet from a matrix and presentation data", () => {
    const workbook = createBlankWorkbook();
    const result = apply(workbook, {
      type: "sheet.createFromMatrix",
      sheetId: "sheet-2",
      name: "Generated",
      rows: [["Heading"], ["42"]],
      formats: [{ range: { start: cell, end: cell }, format: { bold: true } }],
      columnWidths: [123],
      freeze: { rows: 1, columns: 0 }
    });
    expect(result.status).toBe("applied");
    if (result.status !== "applied") return;
    expect(result.workbook.activeSheetId).toBe("sheet-2");
    expect(getCellContent(result.workbook, "sheet-2", "A1")).toBe("Heading");
    expect(getCellFormat(result.workbook, "sheet-2", "A1")).toMatchObject({ bold: true });
    expect(getColumnWidth(result.workbook, "sheet-2", 0)).toBe(123);
    expect(result.workbook.sheets[1].freezeTopRow).toBe(true);
  });
});

function apply(workbook: ReturnType<typeof createBlankWorkbook>, command: WorkbookCommand) {
  return applyWorkbookMutation(workbook, command, {
    evaluateCell(candidate, sheetId, address) {
      return getCellContent(candidate, sheetId, address);
    }
  });
}

function containsFunction(value: unknown): boolean {
  if (typeof value === "function") {
    return true;
  }
  if (Array.isArray(value)) {
    return value.some(containsFunction);
  }
  if (value && typeof value === "object") {
    return Object.values(value).some(containsFunction);
  }
  return false;
}
