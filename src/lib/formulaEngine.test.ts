import { describe, expect, it } from "vitest";
import {
  createBlankWorkbook,
  defineNamedRange,
  setCellContent
} from "./workbook";
import { createFormulaEngine } from "./formulaEngine";

describe("formulaEngine", () => {
  it("shows raw values and formula results", () => {
    let workbook = createBlankWorkbook();
    workbook = setCellContent(workbook, workbook.activeSheetId, "A1", "10");
    workbook = setCellContent(workbook, workbook.activeSheetId, "A2", "20");
    workbook = setCellContent(workbook, workbook.activeSheetId, "A3", "=SUM(A1:A2)");

    const engine = createFormulaEngine(workbook);

    expect(engine.getDisplayValue(workbook.activeSheetId, "A1")).toBe("10");
    expect(engine.getDisplayValue(workbook.activeSheetId, "A3")).toBe("30");
    expect(engine.getRawContent(workbook.activeSheetId, "A3")).toBe("=SUM(A1:A2)");
  });

  it("supports common Excel-like functions", () => {
    let workbook = createBlankWorkbook();
    workbook = setCellContent(workbook, workbook.activeSheetId, "A1", "2");
    workbook = setCellContent(workbook, workbook.activeSheetId, "A2", "4");
    workbook = setCellContent(workbook, workbook.activeSheetId, "A3", "6");
    workbook = setCellContent(workbook, workbook.activeSheetId, "B1", "=AVERAGE(A1:A3)");
    workbook = setCellContent(workbook, workbook.activeSheetId, "B2", '=IF(B1>3,"yes","no")');

    const engine = createFormulaEngine(workbook);

    expect(engine.getDisplayValue(workbook.activeSheetId, "B1")).toBe("4");
    expect(engine.getDisplayValue(workbook.activeSheetId, "B2")).toBe("yes");
  });

  it("calculates formulas that use workbook named ranges", () => {
    let workbook = createBlankWorkbook();
    const sheetId = workbook.activeSheetId;
    workbook = setCellContent(workbook, sheetId, "A1", "10");
    workbook = setCellContent(workbook, sheetId, "A2", "20");
    workbook = setCellContent(workbook, sheetId, "A3", "30");
    workbook = defineNamedRange(workbook, sheetId, "Sales", {
      start: { row: 0, column: 0 },
      end: { row: 2, column: 0 }
    });
    workbook = setCellContent(workbook, sheetId, "B1", "=SUM(Sales)");

    const engine = createFormulaEngine(workbook);

    expect(engine.getDisplayValue(sheetId, "B1")).toBe("60");
  });

  it("recalculates when rebuilt with updated workbook content", () => {
    let workbook = createBlankWorkbook();
    workbook = setCellContent(workbook, workbook.activeSheetId, "A1", "10");
    workbook = setCellContent(workbook, workbook.activeSheetId, "A2", "=A1*2");

    const engine = createFormulaEngine(workbook);
    expect(engine.getDisplayValue(workbook.activeSheetId, "A2")).toBe("20");

    workbook = setCellContent(workbook, workbook.activeSheetId, "A1", "11");
    engine.rebuild(workbook);
    expect(engine.getDisplayValue(workbook.activeSheetId, "A2")).toBe("22");
  });

  it("displays spreadsheet errors without throwing", () => {
    let workbook = createBlankWorkbook();
    workbook = setCellContent(workbook, workbook.activeSheetId, "A1", "=1/0");
    workbook = setCellContent(workbook, workbook.activeSheetId, "A2", "=MISSINGFUNC(1)");

    const engine = createFormulaEngine(workbook);

    expect(engine.getDisplayValue(workbook.activeSheetId, "A1")).toBe("#DIV/0!");
    expect(engine.getDisplayValue(workbook.activeSheetId, "A2")).toBe("#NAME?");
  });
});
