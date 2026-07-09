import { HyperFormula } from "hyperformula";
import { describe, expect, it, vi } from "vitest";
import { createFormulaEngine } from "./formulaEngine";
import { createBlankWorkbook, setCellContent } from "./workbook";
import type { WorkbookModel } from "../types";

describe("formula engine at scale", () => {
  it("builds 100k rows without crashing and updates single cells incrementally", () => {
    // Build a 100k-row x 3-col sheet directly on the model (setCellContent per cell
    // copies the record each time, so construct the cells object in one shot).
    const blank = createBlankWorkbook();
    const cells: Record<string, string | number> = {};
    for (let row = 1; row <= 100_000; row += 1) {
      cells[`A${row}`] = row;
      cells[`B${row}`] = row * 2;
    }
    cells["C1"] = "=SUM(A1:A100000)";
    const workbook: WorkbookModel = {
      ...blank,
      sheets: [{ ...blank.sheets[0], rowCount: 100_000, columnCount: 26, cells }]
    };

    const t0 = performance.now();
    const engine = createFormulaEngine(workbook);
    const buildMs = performance.now() - t0;
    expect(engine.getDisplayValue(workbook.sheets[0].id, "C1")).toBe("5000050000");

    // Single-cell edit must be an incremental diff, not a rebuild.
    const setCellContents = vi.spyOn(HyperFormula.prototype, "setCellContents");
    const destroy = vi.spyOn(HyperFormula.prototype, "destroy");
    try {
      const t1 = performance.now();
      const next = setCellContent(workbook, workbook.sheets[0].id, "A1", 101);
      engine.update(next);
      const updateMs = performance.now() - t1;
      expect(engine.getDisplayValue(next.sheets[0].id, "C1")).toBe("5000050100");
      expect({
        cellWrites: setCellContents.mock.calls,
        rebuilds: destroy.mock.calls.length
      }).toEqual({
        cellWrites: [[{ sheet: 0, col: 0, row: 0 }, 101]],
        rebuilds: 0
      });

      // eslint-disable-next-line no-console
      console.log(`build(200k cells)=${Math.round(buildMs)}ms update(1 cell)=${Math.round(updateMs)}ms`);
      // Wall-clock assertions can flake on contended/instrumented runners; set
      // SKIP_PERF_ASSERT=1 to keep only the correctness checks.
      if (!process.env.SKIP_PERF_ASSERT) {
        expect(buildMs).toBeLessThan(30_000);
        expect(updateMs).toBeLessThan(2_000);
      }
    } finally {
      setCellContents.mockRestore();
      destroy.mockRestore();
      engine.destroy();
    }
  });

  it("evaluates ISO and US dates as real dates with working arithmetic", () => {
    let workbook = createBlankWorkbook();
    const sheetId = workbook.activeSheetId;
    workbook = setCellContent(workbook, sheetId, "A1", "2026-01-15");
    workbook = setCellContent(workbook, sheetId, "A2", "01/20/2026");
    workbook = setCellContent(workbook, sheetId, "A3", "=A2-A1");
    const engine = createFormulaEngine(workbook);
    expect(engine.getDisplayValue(sheetId, "A3")).toBe("5");
    engine.destroy();
  });

  it("displays booleans as TRUE/FALSE like Excel", () => {
    let workbook = createBlankWorkbook();
    const sheetId = workbook.activeSheetId;
    workbook = setCellContent(workbook, sheetId, "A1", "=1>0");
    const engine = createFormulaEngine(workbook);
    expect(engine.getDisplayValue(sheetId, "A1")).toBe("TRUE");
    engine.destroy();
  });
});
