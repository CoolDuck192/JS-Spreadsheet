import { loadWorkbook, saveWorkbook, WORKBOOK_STORAGE_KEY } from "../lib/persistence";
import type { WorkbookStorage } from "../App";

const BROWSER_AUTOSAVE_CELL_LIMIT = 100_000;

export class BrowserWorkbookStorageCapacityError extends Error {
  readonly cellLimit = BROWSER_AUTOSAVE_CELL_LIMIT;

  constructor() {
    super("Workbook is too large for browser autosave. Use Export XLSX to save your work.");
    this.name = "BrowserWorkbookStorageCapacityError";
  }
}

export function createBrowserWorkbookStorage(storage: Storage): WorkbookStorage {
  return {
    load() {
      return loadWorkbook(storage);
    },
    save(workbook) {
      if (hasMoreThanCellLimit(workbook, BROWSER_AUTOSAVE_CELL_LIMIT)) {
        throw new BrowserWorkbookStorageCapacityError();
      }
      if (!saveWorkbook(storage, workbook)) {
        throw new Error("Browser storage write failed");
      }
    },
    clear() {
      storage.removeItem(WORKBOOK_STORAGE_KEY);
    }
  };
}

export function getDefaultBrowserWorkbookStorage(): WorkbookStorage | false {
  try {
    return typeof window === "undefined"
      ? false
      : createBrowserWorkbookStorage(window.localStorage);
  } catch {
    return false;
  }
}

function hasMoreThanCellLimit(
  workbook: Parameters<WorkbookStorage["save"]>[0],
  limit: number
): boolean {
  let count = 0;
  for (const sheet of workbook.sheets) {
    for (const address in sheet.cells) {
      void address;
      count += 1;
      if (count > limit) {
        return true;
      }
    }
  }
  return false;
}
