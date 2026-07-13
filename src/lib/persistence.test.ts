import { describe, expect, it } from "vitest";
import {
  addSheet,
  createBlankWorkbook,
  defineNamedRange,
  isSheetHidden,
  mergeCells,
  setCellBorders,
  setCellComment,
  setCellContent,
  setCellHyperlink,
  setColumnsHidden,
  setColumnWidth,
  setRangeReadOnly,
  setSheetFreezePanes,
  setRowsHidden,
  setSheetHidden,
  setSheetProtection,
  setRowHeight
} from "./workbook";
import {
  loadWorkbook,
  saveWorkbook,
  WorkbookLoadError,
  WORKBOOK_QUARANTINE_KEY_PREFIX,
  WORKBOOK_STORAGE_KEY
} from "./persistence";

class MemoryStorage implements Storage {
  private values = new Map<string, string>();

  get length() {
    return this.values.size;
  }

  clear(): void {
    this.values.clear();
  }

  getItem(key: string): string | null {
    return this.values.get(key) ?? null;
  }

  key(index: number): string | null {
    return Array.from(this.values.keys())[index] ?? null;
  }

  removeItem(key: string): void {
    this.values.delete(key);
  }

  setItem(key: string, value: string): void {
    this.values.set(key, value);
  }
}

function loadStoredWorkbook(storage: Storage) {
  const workbook = loadWorkbook(storage);
  expect(workbook).not.toBeNull();
  return workbook!;
}

describe("persistence", () => {
  it("saves and restores a valid workbook", () => {
    const storage = new MemoryStorage();
    let workbook = createBlankWorkbook();
    workbook = setCellContent(workbook, workbook.activeSheetId, "A1", "saved");
    workbook = setCellBorders(workbook, workbook.activeSheetId, {
      start: { row: 0, column: 0 },
      end: { row: 0, column: 0 }
    }, "bottom");
    workbook = setCellComment(workbook, workbook.activeSheetId, "A1", "Important note");
    workbook = setCellHyperlink(workbook, workbook.activeSheetId, "A1", "https://example.com/report");
    workbook = setSheetProtection(workbook, workbook.activeSheetId, true);
    workbook = setRangeReadOnly(workbook, workbook.activeSheetId, {
      start: { row: 0, column: 0 },
      end: { row: 0, column: 0 }
    }, false);
    workbook = defineNamedRange(workbook, workbook.activeSheetId, "SavedCell", {
      start: { row: 0, column: 0 },
      end: { row: 0, column: 0 }
    });
    workbook = mergeCells(workbook, workbook.activeSheetId, {
      start: { row: 0, column: 0 },
      end: { row: 0, column: 1 }
    });
    workbook = setColumnWidth(workbook, workbook.activeSheetId, 1, 136);
    workbook = setRowHeight(workbook, workbook.activeSheetId, 2, 44);
    workbook = setColumnsHidden(workbook, workbook.activeSheetId, 3, 3, true);
    workbook = setRowsHidden(workbook, workbook.activeSheetId, 4, 4, true);
    workbook = setSheetFreezePanes(workbook, workbook.activeSheetId, { freezeTopRow: true, freezeFirstColumn: true });
    workbook = addSheet(workbook, "Archive");
    const hiddenSheetId = workbook.activeSheetId;
    workbook = setSheetHidden(workbook, hiddenSheetId, true);

    saveWorkbook(storage, workbook);

    expect(loadWorkbook(storage)).toEqual(workbook);
  });

  it("migrates hidden sheet visibility while keeping the active sheet visible", () => {
    const storage = new MemoryStorage();
    storage.setItem(
      "javascript-spreadsheet-workbook",
      JSON.stringify({
        version: 1,
        activeSheetId: "sheet-2",
        namedRanges: [],
        sheets: [
          {
            id: "sheet-1",
            name: "Visible",
            rowCount: 100,
            columnCount: 26,
            cells: {}
          },
          {
            id: "sheet-2",
            name: "Hidden",
            rowCount: 100,
            columnCount: 26,
            isHidden: true,
            cells: {}
          }
        ]
      })
    );

    const workbook = loadStoredWorkbook(storage);

    expect(workbook.activeSheetId).toBe("sheet-1");
    expect(isSheetHidden(workbook, "sheet-2")).toBe(true);
  });

  it("migrates old workbooks without dimension metadata", () => {
    const storage = new MemoryStorage();
    storage.setItem(
      "javascript-spreadsheet-workbook",
      JSON.stringify({
        version: 1,
        activeSheetId: "sheet-1",
        sheets: [
          {
            id: "sheet-1",
            name: "Legacy",
            rowCount: 100,
            columnCount: 26,
            cells: {}
          }
        ]
      })
    );

    const workbook = loadStoredWorkbook(storage);

    expect(workbook.sheets[0].columnWidths).toEqual({});
    expect(workbook.sheets[0].rowHeights).toEqual({});
    expect(workbook.sheets[0].hiddenColumns).toEqual({});
    expect(workbook.sheets[0].hiddenRows).toEqual({});
    expect(workbook.sheets[0].freezeTopRow).toBe(false);
    expect(workbook.sheets[0].freezeFirstColumn).toBe(false);
    expect(workbook.sheets[0].autoFilterRange).toBeUndefined();
    expect(workbook.sheets[0].comments).toEqual({});
    expect(workbook.sheets[0].hyperlinks).toEqual({});
    expect(workbook.sheets[0].protection).toEqual({ isProtected: false, lockedCells: {}, unlockedCells: {} });
    expect(workbook.sheets[0].merges).toEqual([]);
    expect(workbook.namedRanges).toEqual([]);
  });

  it("migrates legacy filter criteria into an auto filter range", () => {
    const storage = new MemoryStorage();
    storage.setItem(
      "javascript-spreadsheet-workbook",
      JSON.stringify({
        version: 1,
        activeSheetId: "sheet-1",
        namedRanges: [],
        sheets: [
          {
            id: "sheet-1",
            name: "Legacy",
            rowCount: 100,
            columnCount: 26,
            cells: {},
            filters: [
              {
                id: "filter-1",
                range: {
                  start: { row: 0, column: 0 },
                  end: { row: 5, column: 2 }
                },
                column: 0,
                operator: "equals",
                value: "West",
                hasHeader: true
              }
            ]
          }
        ]
      })
    );

    const workbook = loadStoredWorkbook(storage);

    expect(workbook.sheets[0].autoFilterRange).toEqual({
      start: { row: 0, column: 0 },
      end: { row: 5, column: 2 }
    });
  });

  it("repairs stale active sheet ids when loading saved workbooks", () => {
    const storage = new MemoryStorage();
    storage.setItem(
      "javascript-spreadsheet-workbook",
      JSON.stringify({
        version: 1,
        activeSheetId: "missing-sheet",
        namedRanges: [],
        sheets: [
          {
            id: "sheet-1",
            name: "Recovered",
            rowCount: 100,
            columnCount: 26,
            cells: { A1: "still here" }
          }
        ]
      })
    );

    const workbook = loadStoredWorkbook(storage);

    expect(workbook.activeSheetId).toBe("sheet-1");
    expect(workbook.sheets[0].cells.A1).toBe("still here");
  });

  it("returns null when browser storage has no workbook", () => {
    const storage = new MemoryStorage();

    expect(loadWorkbook(storage)).toBeNull();
  });

  it("quarantines invalid JSON byte-for-byte and throws a typed error", () => {
    const storage = new MemoryStorage();
    const serialized = "{nope";
    storage.setItem(WORKBOOK_STORAGE_KEY, serialized);

    let failure: unknown;
    try {
      loadWorkbook(storage);
    } catch (error) {
      failure = error;
    }

    expect(failure).toBeInstanceOf(WorkbookLoadError);
    const loadError = failure as WorkbookLoadError;
    expect(loadError.reason).toBe("invalid-json");
    expect(loadError.quarantineKey).toMatch(new RegExp(`^${WORKBOOK_QUARANTINE_KEY_PREFIX}`));
    expect(storage.getItem(WORKBOOK_STORAGE_KEY)).toBe(serialized);
    expect(storage.getItem(loadError.quarantineKey!)).toBe(serialized);
  });

  it("quarantines an unmigratable payload without changing the primary value", () => {
    const storage = new MemoryStorage();
    const serialized = JSON.stringify({ version: 99, activeSheetId: "x", sheets: [] });
    storage.setItem(WORKBOOK_STORAGE_KEY, serialized);

    let failure: unknown;
    try {
      loadWorkbook(storage);
    } catch (error) {
      failure = error;
    }

    expect(failure).toBeInstanceOf(WorkbookLoadError);
    const loadError = failure as WorkbookLoadError;
    expect(loadError.reason).toBe("migration-failed");
    expect(storage.getItem(WORKBOOK_STORAGE_KEY)).toBe(serialized);
    expect(storage.getItem(loadError.quarantineKey!)).toBe(serialized);
  });
});
