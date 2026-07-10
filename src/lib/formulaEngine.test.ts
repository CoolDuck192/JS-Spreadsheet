import { describe, expect, it } from "vitest";
import {
  addSheet,
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

  it("returns evaluated values without flattening their spreadsheet types", () => {
    let workbook = createBlankWorkbook();
    const sheetId = workbook.activeSheetId;
    workbook = setCellContent(workbook, sheetId, "A1", 10);
    workbook = setCellContent(workbook, sheetId, "A2", "=A1=10");
    workbook = setCellContent(workbook, sheetId, "A3", '=IF(A2,"ready","")');
    workbook = setCellContent(workbook, sheetId, "A4", '=IF(A2,"",1)');
    workbook = setCellContent(workbook, sheetId, "A5", "=1/0");

    const engine = createFormulaEngine(workbook);

    expect(engine.getComputedValue(sheetId, "A1")).toBe(10);
    expect(engine.getComputedValue(sheetId, "A2")).toBe(true);
    expect(engine.getComputedValue(sheetId, "A3")).toBe("ready");
    expect(engine.getComputedValue(sheetId, "A4")).toBe("");
    expect(engine.getComputedValue(sheetId, "A5")).toEqual({ kind: "error", code: "#DIV/0!" });
    expect(engine.getComputedValue(sheetId, "A6")).toBeNull();
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

  it("projects structured totals through table, legacy, and hidden-row visibility masks", () => {
    let workbook = createBlankWorkbook();
    const sheetId = workbook.activeSheetId;
    workbook = setCellContent(workbook, sheetId, "A1", "Name");
    workbook = setCellContent(workbook, sheetId, "B1", "Amount");
    workbook = setCellContent(workbook, sheetId, "A2", "Ada");
    workbook = setCellContent(workbook, sheetId, "B2", 10);
    workbook = setCellContent(workbook, sheetId, "A3", "Grace");
    workbook = setCellContent(workbook, sheetId, "B3", 20);
    workbook = setCellContent(workbook, sheetId, "A4", "Linus");
    workbook = setCellContent(workbook, sheetId, "B4", 30);
    workbook = setCellContent(workbook, sheetId, "B5", "=SUBTOTAL(109,B2:B4)");
    workbook = {
      ...workbook,
      tables: [{
        id: "table-1", name: "People", sheetId,
        range: { start: { row: 0, column: 0 }, end: { row: 4, column: 1 } },
        headerRow: true, totalsRow: true,
        columns: [
          { id: "name", name: "Name", sheetColumn: 0 },
          { id: "amount", name: "Amount", sheetColumn: 1, dataType: "number", totalsFunction: "sum" }
        ],
        rowIds: ["row-1", "row-2", "row-3"],
        filter: {
          kind: "comparison", columnId: "name", operator: "startsWith",
          value: { type: "string", value: "A" }
        }
      }]
    };
    workbook = addSheet(workbook, "Summary");
    const summaryId = workbook.activeSheetId;
    workbook = setCellContent(workbook, summaryId, "A1", "=Sheet1!B5");
    const engine = createFormulaEngine(workbook);
    try {
      expect(engine.getRawContent(sheetId, "B5")).toBe("=SUBTOTAL(109,B2:B4)");
      expect(engine.getComputedValue(sheetId, "B5")).toBe(10);
      expect(engine.getComputedValue(summaryId, "A1")).toBe(10);

      workbook = { ...workbook, tables: [{ ...workbook.tables[0], filter: undefined }] };
      engine.update(workbook);
      expect(engine.getComputedValue(sheetId, "B5")).toBe(60);
      expect(engine.getComputedValue(summaryId, "A1")).toBe(60);

      workbook = {
        ...workbook,
        sheets: workbook.sheets.map((sheet) => sheet.id === sheetId
          ? { ...sheet, hiddenRows: { "2": true } }
          : sheet)
      };
      engine.update(workbook);
      expect(engine.getComputedValue(sheetId, "B5")).toBe(40);

      workbook = {
        ...workbook,
        sheets: workbook.sheets.map((sheet) => sheet.id === sheetId
          ? {
              ...sheet,
              filters: [{
                id: "legacy", range: { start: { row: 0, column: 0 }, end: { row: 3, column: 1 } },
                column: 0, operator: "equals", value: "Ada", hasHeader: true
              }]
            }
          : sheet)
      };
      engine.update(workbook);
      expect(engine.getComputedValue(sheetId, "B5")).toBe(10);
      expect(engine.getComputedValue(summaryId, "A1")).toBe(10);
      expect(workbook.sheets[0].hiddenRows).toEqual({ "2": true });
    } finally {
      engine.destroy();
    }
  });

  it("uses Excel-compatible empty, error, count, and sample aggregate semantics", () => {
    let workbook = createBlankWorkbook();
    const sheetId = workbook.activeSheetId;
    workbook = setCellContent(workbook, sheetId, "A1", "Value");
    workbook = setCellContent(workbook, sheetId, "A2", "=1/0");
    workbook = setCellContent(workbook, sheetId, "A3", "text");
    workbook = setCellContent(workbook, sheetId, "A4", null);
    workbook = setCellContent(workbook, sheetId, "A5", "=SUBTOTAL(109,A2:A4)");
    workbook = {
      ...workbook,
      tables: [{
        id: "table-1", name: "Values", sheetId,
        range: { start: { row: 0, column: 0 }, end: { row: 4, column: 0 } },
        headerRow: true, totalsRow: true,
        columns: [{ id: "value", name: "Value", sheetColumn: 0, totalsFunction: "sum" }],
        rowIds: ["row-1", "row-2", "row-3"]
      }]
    };
    const engine = createFormulaEngine(workbook);
    try {
      expect(engine.getComputedValue(sheetId, "A5")).toEqual({ kind: "error", code: "#DIV/0!" });

      workbook = setCellContent(workbook, sheetId, "A5", "=SUBTOTAL(103,A2:A4)");
      workbook = {
        ...workbook,
        tables: [{
          ...workbook.tables[0],
          columns: [{ ...workbook.tables[0].columns[0], totalsFunction: "count" }]
        }]
      };
      engine.update(workbook);
      expect(engine.getComputedValue(sheetId, "A5")).toBe(2);

      workbook = setCellContent(workbook, sheetId, "A2", 10);
      workbook = setCellContent(workbook, sheetId, "A5", "=SUBTOTAL(107,A2:A4)");
      workbook = {
        ...workbook,
        tables: [{
          ...workbook.tables[0],
          columns: [{ ...workbook.tables[0].columns[0], totalsFunction: "standardDeviation" }]
        }]
      };
      engine.update(workbook);
      expect(engine.getComputedValue(sheetId, "A5")).toEqual({ kind: "error", code: "#DIV/0!" });
    } finally {
      engine.destroy();
    }
  });

  it("evaluates every structured aggregate with typed range semantics", () => {
    let workbook = createBlankWorkbook();
    const sheetId = workbook.activeSheetId;
    workbook = setCellContent(workbook, sheetId, "A1", "Value");
    workbook = setCellContent(workbook, sheetId, "A2", 10);
    workbook = setCellContent(workbook, sheetId, "A3", 20);
    workbook = setCellContent(workbook, sheetId, "A4", "text");
    workbook = setCellContent(workbook, sheetId, "A5", "=0");
    workbook = {
      ...workbook,
      tables: [{
        id: "table-1", name: "Values", sheetId,
        range: { start: { row: 0, column: 0 }, end: { row: 4, column: 0 } },
        headerRow: true, totalsRow: true,
        columns: [{ id: "value", name: "Value", sheetColumn: 0, totalsFunction: "sum" }],
        rowIds: ["row-1", "row-2", "row-3"]
      }]
    };
    const engine = createFormulaEngine(workbook);
    try {
      const cases = [
        ["sum", 30],
        ["average", 15],
        ["count", 3],
        ["countNumbers", 2],
        ["min", 10],
        ["max", 20],
        ["standardDeviation", Math.sqrt(50)],
        ["variance", 50]
      ] as const;
      for (const [aggregate, expected] of cases) {
        workbook = {
          ...workbook,
          tables: [{
            ...workbook.tables[0],
            columns: [{ ...workbook.tables[0].columns[0], totalsFunction: aggregate }]
          }]
        };
        engine.update(workbook);
        expect(engine.getComputedValue(sheetId, "A5")).toBeCloseTo(expected);
      }

      workbook = setCellContent(workbook, sheetId, "A5", null);
      workbook = {
        ...workbook,
        tables: [{
          ...workbook.tables[0],
          columns: [{ ...workbook.tables[0].columns[0], totalsFunction: "none" }]
        }]
      };
      engine.update(workbook);
      expect(engine.getComputedValue(sheetId, "A5")).toBeNull();
    } finally {
      engine.destroy();
    }
  });
});
