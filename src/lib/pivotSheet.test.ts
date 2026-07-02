import { describe, expect, it } from "vitest";
import type { CellRange } from "../types";
import { parseRangeAddress } from "./addressing";
import { createPivotTableWithDrilldowns, getPivotDrilldownCell, togglePivotDrilldown } from "./pivot";
import { replaceGeneratedPivotSheetRows } from "./pivotSheet";
import {
  addConditionalFormatRule,
  addSheet,
  addSheetChart,
  createBlankWorkbook,
  getActiveSheet,
  getCellComment,
  getCellConditionalFormatRules,
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

const range = (start: string, end = start): CellRange => parseRangeAddress(`${start}:${end}`);

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
    workbook = setCellContent(workbook, sheetId, "AA1", "Keep wide note");
    workbook = setCellContent(workbook, sheetId, "F150", "Keep tall note");

    const drilldown = getPivotDrilldownCell(getActiveSheet(workbook).pivot, 1, 2);
    const expandedPivot = togglePivotDrilldown(getActiveSheet(workbook).pivot!, drilldown!.entry.id);
    const expandedWorkbook = replaceGeneratedPivotSheetRows(workbook, sheetId, expandedPivot);

    expect(getActiveSheet(expandedWorkbook).columnCount).toBeGreaterThanOrEqual(27);
    expect(getActiveSheet(expandedWorkbook).rowCount).toBeGreaterThanOrEqual(150);
    expect(getCellContent(expandedWorkbook, sheetId, "F10")).toBe("Keep this note");
    expect(getCellContent(expandedWorkbook, sheetId, "AA1")).toBe("Keep wide note");
    expect(getCellContent(expandedWorkbook, sheetId, "F150")).toBe("Keep tall note");
    expect(getCellFormat(expandedWorkbook, sheetId, "F10")).toEqual({ bold: true, backgroundColor: "#eaf7f2" });
    expect(getCellComment(expandedWorkbook, sheetId, "F10")).toBe("Keep comment");
    expect(getCellHyperlink(expandedWorkbook, sheetId, "F10")).toBe("https://example.com/keep");
    expect(getCellValidation(expandedWorkbook, sheetId, "F10")).toEqual({ type: "list", values: ["Keep this note"] });
    expect(getActiveSheet(expandedWorkbook).conditionalFormats).toHaveLength(1);
    expect(getSheetCharts(expandedWorkbook, sheetId)).toHaveLength(1);
    expect(getActiveSheet(expandedWorkbook).merges).toHaveLength(1);
  });

  it("scrubs user artifacts from regenerated pivot cells without touching cells beside the pivot", () => {
    const pivot = createPivotTableWithDrilldowns(
      [
        ["Region", "Product", "Sales"],
        ["West", "Hardware", "10"],
        ["East", "Hardware", "8"]
      ],
      {
        rowFields: ["Region", "Product"],
        valueField: "Sales",
        aggregator: "SUM"
      }
    );

    let workbook = addSheet(createBlankWorkbook(), "Pivot 1");
    const sheetId = workbook.activeSheetId;
    workbook = replaceGeneratedPivotSheetRows(workbook, sheetId, pivot.metadata);
    workbook = setCellContent(workbook, sheetId, "F2", "Beside pivot");
    workbook = setCellFormat(workbook, sheetId, range("F2"), { italic: true });
    workbook = setCellComment(workbook, sheetId, "B2", "Generated comment");
    workbook = setCellHyperlink(workbook, sheetId, "B2", "https://example.com/generated");
    workbook = setCellValidation(workbook, sheetId, range("B2"), { type: "list", values: ["Generated"] });
    workbook = addConditionalFormatRule(workbook, sheetId, range("B2"), {
      condition: { type: "textContains", value: "Generated" },
      format: { bold: true }
    });
    workbook = addSheetChart(workbook, sheetId, range("B2"), {
      anchor: { row: 1, column: 1 },
      title: "Generated chart",
      type: "bar"
    });
    workbook = mergeCells(workbook, sheetId, range("A3", "B3"));

    const drilldown = getPivotDrilldownCell(getActiveSheet(workbook).pivot, 1, 2);
    const expandedPivot = togglePivotDrilldown(getActiveSheet(workbook).pivot!, drilldown!.entry.id);
    const expandedWorkbook = replaceGeneratedPivotSheetRows(workbook, sheetId, expandedPivot);

    expect(getCellContent(expandedWorkbook, sheetId, "F2")).toBe("Beside pivot");
    expect(getCellFormat(expandedWorkbook, sheetId, "F2")).toEqual({ italic: true });
    expect(getCellComment(expandedWorkbook, sheetId, "B2")).toBeNull();
    expect(getCellHyperlink(expandedWorkbook, sheetId, "B2")).toBeNull();
    expect(getCellValidation(expandedWorkbook, sheetId, "B2")).toBeNull();
    expect(getCellConditionalFormatRules(expandedWorkbook, sheetId, "B2")).toHaveLength(0);
    expect(getSheetCharts(expandedWorkbook, sheetId)).toHaveLength(0);
    expect(getActiveSheet(expandedWorkbook).merges).toHaveLength(0);
  });
});
