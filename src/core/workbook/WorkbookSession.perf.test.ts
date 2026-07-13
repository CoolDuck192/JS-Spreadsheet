import { describe, expect, it } from "vitest";
import type { WorkbookModel } from "../../types";
import type { FormulaEngine } from "../../lib/formulaEngine";
import { getCellContent } from "../../lib/workbook";
import { createBlankWorkbook } from "../../lib/workbook";
import { createWorkbookSession } from "./WorkbookSession";

const LARGE_ROW_COUNT = 100_000;

describe("WorkbookSession at 100,000 rows", () => {
  it("keeps projection, publication, command bursts, and weighted history bounded", () => {
    const workbook = createLargeWorkbook();
    const sheetId = workbook.activeSheetId;
    const buildStarted = performance.now();
    const session = createWorkbookSession({ workbook });
    const buildMs = performance.now() - buildStarted;
    const publications: Array<{ revision: string; total: unknown }> = [];
    session.subscribe(() => {
      publications.push({
        revision: session.getSnapshot().revision,
        total: session.getCellEvaluation(sheetId, "B1")
      });
    });

    try {
      const editStarted = performance.now();
      expect(session.dispatch({ type: "cell.set", sheetId, address: "A1", input: "101" }))
        .toEqual({ status: "committed", revision: "1", changed: true });
      const editMs = performance.now() - editStarted;
      expect(session.getCellEvaluation(sheetId, "B1")).toBe(5_000_050_100);
      expect(publications).toEqual([{ revision: "1", total: 5_000_050_100 }]);

      const burstStarted = performance.now();
      const burst = session.dispatch({
        type: "transaction",
        commands: Array.from({ length: 100 }, (_, index) => ({
          type: "cell.set" as const,
          sheetId,
          address: `A${index + 2}`,
          input: String(index + 2 + LARGE_ROW_COUNT)
        }))
      });
      const burstMs = performance.now() - burstStarted;
      expect(burst).toEqual({ status: "committed", revision: "2", changed: true });
      expect(session.getCellEvaluation(sheetId, "B1")).toBe(5_010_050_100);
      expect(publications).toEqual([
        { revision: "1", total: 5_000_050_100 },
        { revision: "2", total: 5_010_050_100 }
      ]);

      // eslint-disable-next-line no-console
      console.log(
        `workbook-session build=${Math.round(buildMs)}ms edit+recalc=${Math.round(editMs)}ms `
        + `atomic-100-command-burst=${Math.round(burstMs)}ms`
      );
      if (!process.env.SKIP_PERF_ASSERT) {
        expect(buildMs).toBeLessThan(30_000);
        expect(editMs).toBeLessThan(2_000);
        expect(burstMs).toBeLessThan(10_000);
      }
    } finally {
      session.destroy();
    }

    const bulkSession = createWorkbookSession({
      workbook,
      formulaEngineFactory: createRawProjection
    });
    try {
      const bulkClearStarted = performance.now();
      const bulkClear = bulkSession.dispatch({
        type: "range.clear",
        sheetId,
        range: {
          start: { row: 0, column: 0 },
          end: { row: LARGE_ROW_COUNT - 1, column: 0 }
        },
        mode: "contents"
      });
      const bulkClearMs = performance.now() - bulkClearStarted;
      expect(bulkClear).toEqual({ status: "committed", revision: "1", changed: true });
      expect(getCellContent(bulkSession.getSnapshot().workbook, sheetId, "A1")).toBeNull();
      expect(getCellContent(
        bulkSession.getSnapshot().workbook,
        sheetId,
        `A${LARGE_ROW_COUNT}`
      )).toBeNull();
      // eslint-disable-next-line no-console
      console.log(`table-free-100k-content-clear=${Math.round(bulkClearMs)}ms`);
      if (!process.env.SKIP_PERF_ASSERT) {
        expect(bulkClearMs).toBeLessThan(10_000);
      }
    } finally {
      bulkSession.destroy();
    }

    const historyStarted = performance.now();
    const historySession = createWorkbookSession({
      workbook,
      formulaEngineFactory: createRawProjection,
      history: { maxEntries: 10, maxWeight: 200_010 }
    });
    try {
      for (let index = 0; index < 4; index += 1) {
        expect(historySession.dispatch({
          type: "cell.set",
          sheetId,
          address: `A${index + 1}`,
          input: String(index + 500)
        })).toMatchObject({ status: "committed", changed: true });
      }
      let undoCount = 0;
      while (historySession.getSnapshot().canUndo) {
        expect(historySession.dispatch({ type: "history.undo" }))
          .toMatchObject({ status: "committed", changed: true });
        undoCount += 1;
      }
      const historyMs = performance.now() - historyStarted;
      expect(undoCount).toBe(2);
      expect(historySession.getSnapshot().canUndo).toBe(false);
      expect(historySession.getSnapshot().canRedo).toBe(true);
      // eslint-disable-next-line no-console
      console.log(`weighted-history-trim=${Math.round(historyMs)}ms retained-undo=${undoCount}`);
      if (!process.env.SKIP_PERF_ASSERT) {
        expect(historyMs).toBeLessThan(5_000);
      }
    } finally {
      historySession.destroy();
    }
  }, 60_000);
});

function createLargeWorkbook(): WorkbookModel {
  const workbook = createBlankWorkbook();
  const cells: Record<string, string | number> = {};
  for (let row = 1; row <= LARGE_ROW_COUNT; row += 1) {
    cells[`A${row}`] = row;
  }
  cells.B1 = `=SUM(A1:A${LARGE_ROW_COUNT})`;
  return {
    ...workbook,
    sheets: [{
      ...workbook.sheets[0],
      rowCount: LARGE_ROW_COUNT,
      cells
    }]
  };
}

function createRawProjection(workbook: WorkbookModel): FormulaEngine {
  let projected = workbook;
  return {
    getDisplayValue(sheetId, address) {
      const value = getCellContent(projected, sheetId, address);
      return value === null ? "" : String(value);
    },
    getComputedValue(sheetId, address) {
      return getCellContent(projected, sheetId, address);
    },
    getRawContent(sheetId, address) {
      return getCellContent(projected, sheetId, address);
    },
    update(nextWorkbook) {
      projected = nextWorkbook;
    },
    rebuild(nextWorkbook) {
      projected = nextWorkbook;
    },
    destroy() {}
  };
}
