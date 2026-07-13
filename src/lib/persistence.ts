import type { WorkbookModel } from "../types";
import { migrateWorkbookModel } from "../core/workbook/migrateWorkbook";

export const WORKBOOK_STORAGE_KEY = "javascript-spreadsheet-workbook";
export const WORKBOOK_QUARANTINE_KEY_PREFIX = `${WORKBOOK_STORAGE_KEY}.corrupt.`;

export type WorkbookLoadFailureReason = "invalid-json" | "migration-failed";

export class WorkbookLoadError extends Error {
  readonly name = "WorkbookLoadError";

  constructor(
    readonly reason: WorkbookLoadFailureReason,
    readonly quarantineKey: string | null
  ) {
    super("Stored workbook could not be opened");
  }
}

export function saveWorkbook(storage: Storage, workbook: WorkbookModel): boolean {
  try {
    storage.setItem(WORKBOOK_STORAGE_KEY, JSON.stringify(workbook));
    return true;
  } catch {
    // Quota exceeded or storage unavailable — the workbook stays usable in memory.
    return false;
  }
}

export function loadWorkbook(storage: Storage): WorkbookModel | null {
  const serialized = storage.getItem(WORKBOOK_STORAGE_KEY);
  if (serialized === null) {
    return null;
  }

  let parsed: unknown;
  try {
    parsed = JSON.parse(serialized);
  } catch {
    throw createLoadError(storage, serialized, "invalid-json");
  }

  let workbook: WorkbookModel | null;
  try {
    workbook = migrateWorkbookModel(parsed);
  } catch {
    workbook = null;
  }
  if (!workbook) {
    throw createLoadError(storage, serialized, "migration-failed");
  }
  return workbook;
}

function createLoadError(
  storage: Storage,
  serialized: string,
  reason: WorkbookLoadFailureReason
): WorkbookLoadError {
  return new WorkbookLoadError(reason, quarantinePayload(storage, serialized));
}

function quarantinePayload(storage: Storage, serialized: string): string | null {
  try {
    for (let index = 0; index < storage.length; index += 1) {
      const key = storage.key(index);
      if (key?.startsWith(WORKBOOK_QUARANTINE_KEY_PREFIX) && storage.getItem(key) === serialized) {
        return key;
      }
    }

    const stem = `${WORKBOOK_QUARANTINE_KEY_PREFIX}${Date.now()}`;
    let key = stem;
    let suffix = 1;
    while (storage.getItem(key) !== null) {
      key = `${stem}.${suffix}`;
      suffix += 1;
    }
    storage.setItem(key, serialized);
    return key;
  } catch {
    return null;
  }
}
