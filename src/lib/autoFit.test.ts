import { describe, expect, it } from "vitest";
import { normalizeRange } from "./addressing";
import {
  createAutoFitColumnPlan,
  createAutoFitRowPlan,
  textToAutoFitColumnWidth,
  textToAutoFitRowHeight
} from "./autoFit";
import { createBlankWorkbook, getActiveSheet, setCellContent } from "./workbook";
import {
  DEFAULT_COLUMN_WIDTH,
  DEFAULT_ROW_HEIGHT,
  MAX_COLUMN_WIDTH,
  MAX_ROW_HEIGHT,
  MIN_COLUMN_WIDTH,
  MIN_ROW_HEIGHT
} from "./sheetDimensions";

describe("autoFit", () => {
  it("plans wider columns from visible values and column headers", () => {
    let workbook = createBlankWorkbook();
    workbook = setCellContent(workbook, workbook.activeSheetId, "B1", "A much longer customer segment name");
    const sheet = getActiveSheet(workbook);

    const plan = createAutoFitColumnPlan(sheet, normalizeRange({ start: { row: 0, column: 1 }, end: { row: 2, column: 1 } }), (address) =>
      String(sheet.cells[address] ?? "")
    );

    expect(plan).toHaveLength(1);
    expect(plan[0]).toEqual({ column: 1, width: textToAutoFitColumnWidth("A much longer customer segment name") });
    expect(plan[0].width).toBeGreaterThan(DEFAULT_COLUMN_WIDTH);
  });

  it("plans row heights from multiline values", () => {
    let workbook = createBlankWorkbook();
    workbook = setCellContent(workbook, workbook.activeSheetId, "A2", "Line one\nLine two\nLine three");
    const sheet = getActiveSheet(workbook);

    const plan = createAutoFitRowPlan(sheet, normalizeRange({ start: { row: 1, column: 0 }, end: { row: 1, column: 2 } }), (address) =>
      String(sheet.cells[address] ?? "")
    );

    expect(plan).toEqual([{ row: 1, height: textToAutoFitRowHeight("Line one\nLine two\nLine three") }]);
    expect(plan[0].height).toBeGreaterThan(DEFAULT_ROW_HEIGHT);
  });

  it("clamps calculated sizes to spreadsheet dimension limits", () => {
    expect(textToAutoFitColumnWidth("x")).toBeGreaterThanOrEqual(MIN_COLUMN_WIDTH);
    expect(textToAutoFitColumnWidth("x".repeat(500))).toBe(MAX_COLUMN_WIDTH);
    expect(textToAutoFitRowHeight("single")).toBeGreaterThanOrEqual(MIN_ROW_HEIGHT);
    expect(textToAutoFitRowHeight(Array.from({ length: 20 }, (_, index) => `Line ${index}`).join("\n"))).toBe(MAX_ROW_HEIGHT);
  });
});
