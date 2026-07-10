import { describe, expect, it } from "vitest";
import type { SheetModel, WorkbookModel } from "../../types";
import { addSheet, createBlankWorkbook, setCellContent } from "../../lib/workbook";
import {
  commitWorkbookHistory,
  createWorkbookHistory,
  estimateWorkbookWeight,
  getWorkbookHistoryStats,
  redoWorkbookHistory,
  undoWorkbookHistory
} from "./history";

const range = {
  start: { row: 0, column: 0 },
  end: { row: 1, column: 1 }
};

function replaceActiveSheet(
  workbook: WorkbookModel,
  changes: Partial<SheetModel>
): WorkbookModel {
  return {
    ...workbook,
    sheets: workbook.sheets.map((sheet) =>
      sheet.id === workbook.activeSheetId ? { ...sheet, ...changes } : sheet
    )
  };
}

function workbookWithCells(cellCount: number): WorkbookModel {
  const workbook = createBlankWorkbook();
  const cells = Object.fromEntries(
    Array.from({ length: cellCount }, (_, index) => [`A${index + 1}`, index])
  );
  return replaceActiveSheet(workbook, {
    rowCount: Math.max(workbook.sheets[0].rowCount, cellCount),
    cells
  });
}

describe("workbook history", () => {
  it("counts every current persisted cell and metadata collection deterministically", () => {
    const workbook = createBlankWorkbook();
    const sheet = workbook.sheets[0];
    const baselineWeight = estimateWorkbookWeight(workbook);
    const variants: Record<string, WorkbookModel> = {
      cells: replaceActiveSheet(workbook, { cells: { A1: "value" } }),
      formats: replaceActiveSheet(workbook, { formats: { A1: { bold: true } } }),
      columnWidths: replaceActiveSheet(workbook, { columnWidths: { "0": 120 } }),
      rowHeights: replaceActiveSheet(workbook, { rowHeights: { "0": 28 } }),
      hiddenColumns: replaceActiveSheet(workbook, { hiddenColumns: { "0": true } }),
      hiddenRows: replaceActiveSheet(workbook, { hiddenRows: { "0": true } }),
      comments: replaceActiveSheet(workbook, { comments: { A1: "note" } }),
      hyperlinks: replaceActiveSheet(workbook, { hyperlinks: { A1: "https://example.com" } }),
      validations: replaceActiveSheet(workbook, {
        validations: { A1: { type: "list", values: ["one", "two"] } }
      }),
      conditionalFormats: replaceActiveSheet(workbook, {
        conditionalFormats: [{
          id: "conditional-1",
          range,
          condition: { type: "greaterThan", value: "1" },
          format: { backgroundColor: "#ffffff" }
        }]
      }),
      autoFilterRange: replaceActiveSheet(workbook, { autoFilterRange: range }),
      filters: replaceActiveSheet(workbook, {
        filters: [{
          id: "filter-1",
          range,
          column: 0,
          operator: "equals",
          value: "one",
          values: ["one", "two"]
        }]
      }),
      charts: replaceActiveSheet(workbook, {
        charts: [{ id: "chart-1", title: "Chart", type: "bar", range, anchor: { row: 2, column: 2 } }]
      }),
      merges: replaceActiveSheet(workbook, {
        merges: [{ id: "merge-1", range }]
      }),
      lockedCells: replaceActiveSheet(workbook, {
        protection: { ...sheet.protection, lockedCells: { A1: true } }
      }),
      unlockedCells: replaceActiveSheet(workbook, {
        protection: { ...sheet.protection, unlockedCells: { A1: true } }
      }),
      namedRanges: {
        ...workbook,
        namedRanges: [{ name: "Selection", sheetId: sheet.id, range }]
      },
      sheets: addSheet(workbook, "Second")
    };

    for (const [plane, variant] of Object.entries(variants)) {
      expect(estimateWorkbookWeight(variant), plane).toBeGreaterThan(baselineWeight);
      expect(estimateWorkbookWeight(structuredClone(variant)), plane).toBe(
        estimateWorkbookWeight(variant)
      );
    }
  });

  it("bounds retained snapshots by both count and total logical weight", () => {
    const weightedWorkbook = workbookWithCells(25_000);
    const snapshotWeight = estimateWorkbookWeight(weightedWorkbook);
    let history = createWorkbookHistory(weightedWorkbook);

    for (let index = 0; index < 105; index += 1) {
      history = commitWorkbookHistory(history, { ...history.present });
    }

    const stats = getWorkbookHistoryStats(history);
    const expectedCount = Math.min(100, Math.floor(2_000_000 / snapshotWeight));

    expect(history.limits).toEqual({ maxEntries: 100, maxWeight: 2_000_000 });
    expect(history.past).toHaveLength(expectedCount);
    expect(stats).toMatchObject({
      pastCount: expectedCount,
      futureCount: 0,
      retainedCount: expectedCount,
      retainedWeight: expectedCount * snapshotWeight
    });
    expect(stats.retainedWeight).toBeLessThanOrEqual(2_000_000);
  });

  it("keeps exact undo and redo order after old snapshots are trimmed", () => {
    const initial = createBlankWorkbook();
    const sheetId = initial.activeSheetId;
    const snapshots = [initial];
    let history = createWorkbookHistory(initial, { maxEntries: 3 });

    for (let value = 1; value <= 5; value += 1) {
      const next = setCellContent(history.present, sheetId, "A1", value);
      snapshots.push(next);
      history = commitWorkbookHistory(history, next);
    }

    expect(history.past).toEqual([snapshots[2], snapshots[3], snapshots[4]]);

    for (const expected of [snapshots[4], snapshots[3], snapshots[2]]) {
      history = undoWorkbookHistory(history);
      expect(history.present).toBe(expected);
    }
    expect(undoWorkbookHistory(history)).toBe(history);

    for (const expected of [snapshots[3], snapshots[4], snapshots[5]]) {
      history = redoWorkbookHistory(history);
      expect(history.present).toBe(expected);
    }
    expect(redoWorkbookHistory(history)).toBe(history);
  });

  it("clears redo snapshots when a new edit follows undo", () => {
    const initial = createBlankWorkbook();
    const sheetId = initial.activeSheetId;
    const first = setCellContent(initial, sheetId, "A1", 1);
    const second = setCellContent(first, sheetId, "A1", 2);
    let history = createWorkbookHistory(initial);
    history = commitWorkbookHistory(history, first);
    history = commitWorkbookHistory(history, second);
    history = undoWorkbookHistory(history);

    expect(history.future).toEqual([second]);

    const replacement = setCellContent(history.present, sheetId, "A1", 99);
    history = commitWorkbookHistory(history, replacement);

    expect(history.present).toBe(replacement);
    expect(history.future).toEqual([]);
    expect(getWorkbookHistoryStats(history).futureCount).toBe(0);
  });

  it("does not retain an oversized snapshot or leave a non-adjacent undo target", () => {
    const initial = createBlankWorkbook();
    const oversized = workbookWithCells(50);
    const replacement = createBlankWorkbook();
    const maxWeight = estimateWorkbookWeight(oversized) - 1;
    let history = createWorkbookHistory(initial, { maxWeight });

    history = commitWorkbookHistory(history, oversized);
    expect(history.past).toEqual([initial]);

    history = commitWorkbookHistory(history, replacement);

    expect(history.past).toEqual([]);
    expect(getWorkbookHistoryStats(history).retainedWeight).toBe(0);
    expect(undoWorkbookHistory(history)).toBe(history);
  });

  it("retains immutable workbook references and shares untouched sheets", () => {
    const initial = addSheet(createBlankWorkbook(), "Untouched");
    const editedSheetId = initial.sheets[0].id;
    const untouchedSheet = initial.sheets[1];
    const edited = setCellContent(initial, editedSheetId, "A1", "changed");
    let history = createWorkbookHistory(initial);

    expect(edited.sheets[1]).toBe(untouchedSheet);

    history = commitWorkbookHistory(history, edited);
    expect(history.past[0]).toBe(initial);
    expect(history.past[0].sheets[1]).toBe(untouchedSheet);
    expect(history.present.sheets[1]).toBe(untouchedSheet);

    history = undoWorkbookHistory(history);
    expect(history.present).toBe(initial);
    expect(history.future[0]).toBe(edited);

    history = redoWorkbookHistory(history);
    expect(history.present).toBe(edited);
  });
});
