import { describe, expect, it } from "vitest";
import type { CellRange } from "../types";
import { createPivotTableWithDrilldowns, getPivotDrilldownCell, togglePivotDrilldown } from "./pivot";
import { replaceGeneratedPivotSheetRows } from "./pivotSheet";
import {
  addConditionalFormatRule,
  addSheet,
  addSheetChart,
  createBlankWorkbook,
  getActiveSheet,
  getCellComment,
  getCellContent,
  getCellFormat,
  getCellHyperlink,
  getCellValidation,
  getSheetCharts,
  mergeCells,
  setCellComment,
  setCellContent,
  setCellFormat,
  setCellHyperlink,
  setCellValidation
} from "./workbook";

const range = (start: string, end = start): CellRange => ({
  start: { row: Number(start.slice(1)) - 1, column: start.charCodeAt(0) - 65 },
  end: { row: Number(end.slice(1)) - 1, column: end.charCodeAt(0) - 65 }
});

describe("pivotSheet", () => {
  it("preserves user-owned pivot sheet content and artifacts outside the generated pivot region", () => {
    const sourceRows = [
      ["Region", "Product", "Sales"],
      ["West", "Hardware", "10"],
      ["East", "Hardware", "8"]
    ];
    const pivot = createPivotTableWithDrilldowns(sourceRows, {
      rowFields: ["Region", "Product"],
      valueField: "Sales",
      aggregator: "SUM"
    });

    let workbook = addSheet(createBlankWorkbook(), "Pivot 1");
    const sheetId = workbook.activeSheetId;
    workbook = replaceGeneratedPivotSheetRows(workbook, sheetId, pivot.metadata);

    workbook = setCellContent(workbook, sheetId, "F10", "Keep this note");
    workbook = setCellFormat(workbook, sheetId, range("F10"), { bold: true, backgroundColor: "#eaf7f2" });
    workbook = setCellComment(workbook, sheetId, "F10", "Keep comment");
    workbook = setCellHyperlink(workbook, sheetId, "F10", "https://example.com/keep");
    workbook = setCellValidation(workbook, sheetId, range("F10"), { type: "list", values: ["Keep this note"] });
    workbook = addConditionalFormatRule(workbook, sheetId, range("F10"), {
      condition: { type: "textContains", value: "Keep" },
      format: { italic: true }
    });
    workbook = addSheetChart(workbook, sheetId, range("F10"), {
      anchor: { row: 9, column: 7 },
      title: "Keep chart",
      type: "bar"
    });
    workbook = mergeCells(workbook, sheetId, range("H10", "I10"));

    const drilldown = getPivotDrilldownCell(getActiveSheet(workbook).pivot, 1, 2);
    const expandedPivot = togglePivotDrilldown(getActiveSheet(workbook).pivot!, drilldown!.entry.id);
    const expandedWorkbook = replaceGeneratedPivotSheetRows(workbook, sheetId, expandedPivot);

    expect(getCellContent(expandedWorkbook, sheetId, "F10")).toBe("Keep this note");
    expect(getCellFormat(expandedWorkbook, sheetId, "F10")).toEqual({ bold: true, backgroundColor: "#eaf7f2" });
    expect(getCellComment(expandedWorkbook, sheetId, "F10")).toBe("Keep comment");
    expect(getCellHyperlink(expandedWorkbook, sheetId, "F10")).toBe("https://example.com/keep");
    expect(getCellValidation(expandedWorkbook, sheetId, "F10")).toEqual({ type: "list", values: ["Keep this note"] });
    expect(getActiveSheet(expandedWorkbook).conditionalFormats).toHaveLength(1);
    expect(getSheetCharts(expandedWorkbook, sheetId)).toHaveLength(1);
    expect(getActiveSheet(expandedWorkbook).merges).toHaveLength(1);
  });
});
