import type { FilterOperator, SheetFilter } from "../types";
import { parseCellInput } from "../core/values/parseCellInput";
import { normalizeRange } from "./addressing";
import type { ComputedCellValue } from "./formulaEngine";

type FilterMatcher = {
  operator: FilterOperator;
  value: string;
  values?: readonly string[];
};

export function matchesFilterValue(value: ComputedCellValue, filter: FilterMatcher): boolean {
  const normalizedValue = comparableText(value);
  const normalizedFilterValue = filter.value.trim();

  switch (filter.operator) {
    case "contains":
      return normalizedValue.toLocaleLowerCase().includes(normalizedFilterValue.toLocaleLowerCase());
    case "equals": {
      if (isBlankValue(value)) {
        return normalizedFilterValue === "";
      }
      if (isComputedError(value) || typeof value === "boolean") {
        return normalizedValue.toLocaleLowerCase() === normalizedFilterValue.toLocaleLowerCase();
      }
      const numericValue = numericComparableValue(value);
      const numericFilterValue = parseNumericValue(normalizedFilterValue);
      if (numericValue !== null && numericFilterValue !== null) {
        return numericValue === numericFilterValue;
      }
      return normalizedValue.toLocaleLowerCase() === normalizedFilterValue.toLocaleLowerCase();
    }
    case "greaterThan": {
      const numericValue = numericComparableValue(value);
      const numericFilterValue = parseNumericValue(normalizedFilterValue);
      return numericValue !== null && numericFilterValue !== null && numericValue > numericFilterValue;
    }
    case "lessThan": {
      const numericValue = numericComparableValue(value);
      const numericFilterValue = parseNumericValue(normalizedFilterValue);
      return numericValue !== null && numericFilterValue !== null && numericValue < numericFilterValue;
    }
  }
}

export function getVisibleRows(
  rowCount: number,
  filters: readonly SheetFilter[],
  getCellValue: (row: number, column: number) => ComputedCellValue
): number[] {
  const rows: number[] = [];

  for (let row = 0; row < rowCount; row += 1) {
    if (filters.every((filter) => isRowVisibleForFilter(row, filter, getCellValue))) {
      rows.push(row);
    }
  }

  return rows;
}

export function isRowVisibleForFilter(
  row: number,
  filter: SheetFilter,
  getCellValue: (row: number, column: number) => ComputedCellValue
): boolean {
  const range = normalizeRange(filter.range);
  if (row < range.start.row || row > range.end.row) {
    return true;
  }

  if (filter.hasHeader !== false && row === range.start.row) {
    return true;
  }

  const value = getCellValue(row, filter.column);
  if (filter.values && filter.values.length > 0) {
    return filter.values.some((filterValue) => matchesFilterValue(value, { ...filter, value: filterValue }));
  }

  return matchesFilterValue(value, filter);
}

function parseNumericValue(value: string): number | null {
  if (value.trim() === "") {
    return null;
  }
  const typed = parseCellInput(value);
  if ((typed.kind === "number" || typed.kind === "date" || typed.kind === "dateTime") && typeof typed.stored === "number") {
    return typed.stored;
  }
  const parsed = Number(value.replace(/[$,%]/g, "").replace(/,/g, ""));
  return Number.isFinite(parsed) ? parsed : null;
}

function comparableText(value: ComputedCellValue): string {
  if (value === null) {
    return "";
  }
  if (isComputedError(value)) {
    return value.code;
  }
  if (typeof value === "boolean") {
    return value ? "TRUE" : "FALSE";
  }
  return String(value).trim();
}

function numericComparableValue(value: ComputedCellValue): number | null {
  if (typeof value === "number") {
    return Number.isFinite(value) ? value : null;
  }
  if (typeof value !== "string") {
    return null;
  }
  return parseNumericValue(value);
}

function isBlankValue(value: ComputedCellValue): boolean {
  return value === null || (typeof value === "string" && value.trim() === "");
}

function isComputedError(value: ComputedCellValue): value is { kind: "error"; code: string } {
  return typeof value === "object" && value !== null && "kind" in value && value.kind === "error";
}
