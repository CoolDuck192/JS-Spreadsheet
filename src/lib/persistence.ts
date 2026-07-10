import type { WorkbookModel } from "../types";
import { migrateWorkbookModel } from "../core/workbook/migrateWorkbook";
import { createBlankWorkbook } from "./workbook";

export const WORKBOOK_STORAGE_KEY = "javascript-spreadsheet-workbook";

export function saveWorkbook(storage: Storage, workbook: WorkbookModel): boolean {
  try {
    storage.setItem(WORKBOOK_STORAGE_KEY, JSON.stringify(workbook));
    return true;
  } catch {
    // Quota exceeded or storage unavailable — the workbook stays usable in memory.
    return false;
  }
}

export function loadWorkbook(storage: Storage): WorkbookModel {
  const serialized = storage.getItem(WORKBOOK_STORAGE_KEY);
  if (!serialized) {
    return createBlankWorkbook();
  }

  try {
    const parsed: unknown = JSON.parse(serialized);
    return migrateWorkbookModel(parsed) ?? createBlankWorkbook();
  } catch {
    return createBlankWorkbook();
  }
}
