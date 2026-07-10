import { describe, expect, expectTypeOf, it } from "vitest";
import { createBlankWorkbook } from "../../lib/workbook";
import type { CommandEnvelope, CommandResult } from "../commands/types";
import type { WorkbookCommand } from "./commands";

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
  { type: "range.readOnly.set", sheetId: "sheet-1", range, readOnly: true },
  { type: "range.merge", sheetId: "sheet-1", range },
  { type: "range.unmerge", sheetId: "sheet-1", range },
  { type: "range.fill", sheetId: "sheet-1", range, direction: "down" },
  { type: "range.fill", sheetId: "sheet-1", range, direction: "right" },
  { type: "range.autoFill", sheetId: "sheet-1", source: range, target: range },
  { type: "range.sort", sheetId: "sheet-1", range, direction: "asc", sortColumn: 0 },
  { type: "range.removeDuplicates", sheetId: "sheet-1", range },
  { type: "clipboard.paste", sheetId: "sheet-1", target: cell, payload: clipboard, mode: "all" },
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
  { type: "workbook.replace", workbook: createBlankWorkbook(), history: "reset" }
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
});

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
