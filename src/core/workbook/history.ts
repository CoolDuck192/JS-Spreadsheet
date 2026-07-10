import type {
  CellFormat,
  DataValidationRule,
  SheetModel,
  WorkbookHistory,
  WorkbookHistoryLimits,
  WorkbookModel
} from "../../types";

export type { WorkbookHistory, WorkbookHistoryLimits } from "../../types";

export const DEFAULT_MAX_ENTRIES = 100;
export const DEFAULT_MAX_WEIGHT = 2_000_000;

export type WorkbookHistoryStats = Readonly<{
  pastCount: number;
  pastWeight: number;
  futureCount: number;
  futureWeight: number;
  retainedCount: number;
  retainedWeight: number;
  presentWeight: number;
}>;

export function createWorkbookHistory(
  workbook: WorkbookModel,
  limits: Partial<WorkbookHistoryLimits> = {}
): WorkbookHistory {
  return {
    past: [],
    present: workbook,
    future: [],
    limits: normalizeLimits(limits)
  };
}

export function commitWorkbookHistory(
  history: WorkbookHistory,
  workbook: WorkbookModel
): WorkbookHistory {
  const limits = resolveHistoryLimits(history);
  if (history.present === workbook) {
    if (history.future.length === 0) {
      return history;
    }

    const retained = trimRetainedSnapshots(history.past, [], limits);
    return {
      ...history,
      ...retained,
      limits
    };
  }

  const { past } = trimRetainedSnapshots(
    [...history.past, history.present],
    [],
    limits
  );
  return {
    ...history,
    past,
    present: workbook,
    future: [],
    limits
  };
}

export function undoWorkbookHistory(history: WorkbookHistory): WorkbookHistory {
  const previous = history.past.at(-1);
  if (!previous) {
    return history;
  }

  const limits = resolveHistoryLimits(history);
  const retained = trimRetainedSnapshots(
    history.past.slice(0, -1),
    [history.present, ...history.future],
    limits
  );
  return {
    ...history,
    ...retained,
    present: previous,
    limits
  };
}

export function redoWorkbookHistory(history: WorkbookHistory): WorkbookHistory {
  const next = history.future[0];
  if (!next) {
    return history;
  }

  const limits = resolveHistoryLimits(history);
  const retained = trimRetainedSnapshots(
    [...history.past, history.present],
    history.future.slice(1),
    limits
  );
  return {
    ...history,
    ...retained,
    present: next,
    limits
  };
}

export function getWorkbookHistoryStats(history: WorkbookHistory): WorkbookHistoryStats {
  const estimateWeight = createOperationWeightEstimator();
  const pastWeight = totalWeight(history.past, estimateWeight);
  const futureWeight = totalWeight(history.future, estimateWeight);
  return {
    pastCount: history.past.length,
    pastWeight,
    futureCount: history.future.length,
    futureWeight,
    retainedCount: history.past.length + history.future.length,
    retainedWeight: pastWeight + futureWeight,
    presentWeight: estimateWeight(history.present)
  };
}

/**
 * Estimates retained memory in deterministic logical entries rather than
 * runtime-specific bytes. A populated cell or persisted metadata entry counts
 * independently, including nested persisted lists such as validation/filter
 * values and format fields.
 */
export function estimateWorkbookWeight(workbook: WorkbookModel): number {
  return estimateWorkbookWeightWith(workbook, estimateSheetWeight);
}

function estimateWorkbookWeightWith(
  workbook: WorkbookModel,
  estimateSheet: (sheet: SheetModel) => number
): number {
  let weight = 1 + (workbook.namedRanges ?? []).length;
  for (const sheet of workbook.sheets) {
    weight += estimateSheet(sheet);
  }
  return weight;
}

function estimateSheetWeight(sheet: SheetModel): number {
  let weight = 1;
  weight += recordEntryCount(sheet.cells);
  weight += sumRecordEntries(sheet.formats, estimateCellFormatWeight);
  weight += recordEntryCount(sheet.columnWidths);
  weight += recordEntryCount(sheet.rowHeights);
  weight += recordEntryCount(sheet.hiddenColumns);
  weight += recordEntryCount(sheet.hiddenRows);
  weight += recordEntryCount(sheet.comments);
  weight += recordEntryCount(sheet.hyperlinks);
  weight += sumRecordEntries(sheet.validations, estimateValidationWeight);
  weight += (sheet.conditionalFormats ?? []).reduce(
    (total, rule) => total + 1 + estimateConditionWeight(rule.condition) + estimateFormatFields(rule.format),
    0
  );
  weight += sheet.autoFilterRange ? 1 : 0;
  weight += (sheet.filters ?? []).reduce(
    (total, filter) => total + 1 + (filter.values?.length ?? 0),
    0
  );
  weight += (sheet.charts ?? []).length;
  weight += (sheet.merges ?? []).length;
  weight += recordEntryCount(sheet.protection?.lockedCells);
  weight += recordEntryCount(sheet.protection?.unlockedCells);
  weight += sheet.isHidden ? 1 : 0;
  weight += sheet.tabColor ? 1 : 0;
  weight += sheet.freezeTopRow ? 1 : 0;
  weight += sheet.freezeFirstColumn ? 1 : 0;
  weight += sheet.protection?.isProtected ? 1 : 0;
  return weight;
}

function estimateCellFormatWeight(format: CellFormat): number {
  return 1 + estimateFormatFields(format);
}

function estimateFormatFields(format: CellFormat): number {
  let weight = 0;
  for (const [key, value] of Object.entries(format)) {
    if (value === undefined) {
      continue;
    }
    if (key === "borders") {
      weight += Object.values(format.borders ?? {}).filter(Boolean).length;
    } else {
      weight += 1;
    }
  }
  return weight;
}

function estimateValidationWeight(rule: DataValidationRule): number {
  let weight = 1;
  if (rule.type === "list") {
    weight += rule.values.length;
  } else {
    weight += rule.min === undefined ? 0 : 1;
    weight += rule.max === undefined ? 0 : 1;
  }
  weight += rule.allowBlank === undefined ? 0 : 1;
  return weight;
}

function estimateConditionWeight(condition: SheetModel["conditionalFormats"][number]["condition"]): number {
  switch (condition.type) {
    case "between":
      return 2;
    case "greaterThan":
    case "lessThan":
    case "equalTo":
    case "textContains":
    case "top":
    case "bottom":
    case "dataBar":
      return 1;
    case "colorScale":
      return 2;
    case "blank":
    case "notBlank":
    case "duplicate":
    case "unique":
      return 0;
  }
}

function recordEntryCount(record: Readonly<Record<string, unknown>> | undefined): number {
  return record ? Object.keys(record).length : 0;
}

function sumRecordEntries<T>(
  record: Readonly<Record<string, T>> | undefined,
  estimateEntry: (entry: T) => number
): number {
  return record ? Object.values(record).reduce((total, entry) => total + estimateEntry(entry), 0) : 0;
}

function createOperationWeightEstimator(): (workbook: WorkbookModel) => number {
  const sheetWeights = new WeakMap<SheetModel, number>();
  return (workbook) => estimateWorkbookWeightWith(workbook, (sheet) => {
    const measured = sheetWeights.get(sheet);
    if (measured !== undefined) {
      return measured;
    }

    const weight = estimateSheetWeight(sheet);
    sheetWeights.set(sheet, weight);
    return weight;
  });
}

function totalWeight(
  workbooks: readonly WorkbookModel[],
  estimateWeight: (workbook: WorkbookModel) => number
): number {
  return workbooks.reduce((total, workbook) => total + estimateWeight(workbook), 0);
}

function trimRetainedSnapshots(
  initialPast: WorkbookModel[],
  initialFuture: WorkbookModel[],
  limits: WorkbookHistoryLimits
): Pick<WorkbookHistory, "past" | "future"> {
  const estimateWeight = createOperationWeightEstimator();
  const past = [...initialPast];
  const future = [...initialFuture];
  const pastWeights = past.map(estimateWeight);
  const futureWeights = future.map(estimateWeight);
  let count = past.length + future.length;
  let weight = pastWeights.reduce((total, entryWeight) => total + entryWeight, 0)
    + futureWeights.reduce((total, entryWeight) => total + entryWeight, 0);

  while (count > limits.maxEntries || weight > limits.maxWeight) {
    if (past.length >= future.length && past.length > 0) {
      past.shift();
      weight -= pastWeights.shift() ?? 0;
    } else {
      future.pop();
      weight -= futureWeights.pop() ?? 0;
    }
    count -= 1;
  }

  return { past, future };
}

function resolveHistoryLimits(history: WorkbookHistory): WorkbookHistoryLimits {
  return normalizeLimits(history.limits ?? {});
}

function normalizeLimits(limits: Partial<WorkbookHistoryLimits>): WorkbookHistoryLimits {
  return Object.freeze({
    maxEntries: normalizeLimit(limits.maxEntries, DEFAULT_MAX_ENTRIES, "maxEntries"),
    maxWeight: normalizeLimit(limits.maxWeight, DEFAULT_MAX_WEIGHT, "maxWeight")
  });
}

function normalizeLimit(value: number | undefined, fallback: number, name: string): number {
  if (value === undefined) {
    return fallback;
  }
  if (!Number.isFinite(value) || !Number.isInteger(value) || value < 0) {
    throw new RangeError(`${name} must be a finite non-negative integer`);
  }
  return value;
}
