import type { FilterOperator, SheetFilter } from "../types";
import { parseCellInput } from "../core/values/parseCellInput";
import { normalizeRange } from "./addressing";
import type { ComputedCellValue } from "./formulaEngine";

export const BLANK_FILTER_VALUE = "\u0000js-spreadsheet:blank";
export const EMPTY_RESULT_FILTER_VALUE = "\u0000js-spreadsheet:empty-result";

export function foldDeterministicText(value: string): string {
  return value.toLowerCase();
}

export function compareDeterministicText(left: string, right: string): number {
  const leftTokens = tokenizeNaturalText(foldDeterministicText(left));
  const rightTokens = tokenizeNaturalText(foldDeterministicText(right));
  const length = Math.min(leftTokens.length, rightTokens.length);

  for (let index = 0; index < length; index += 1) {
    const comparison = compareNaturalToken(leftTokens[index], rightTokens[index]);
    if (comparison !== 0) {
      return comparison;
    }
  }

  return leftTokens.length - rightTokens.length;
}

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
      return foldDeterministicText(normalizedValue).includes(foldDeterministicText(normalizedFilterValue));
    case "equals": {
      if (normalizedFilterValue === BLANK_FILTER_VALUE) {
        return value === null;
      }
      if (normalizedFilterValue === EMPTY_RESULT_FILTER_VALUE) {
        return value === "";
      }
      if (isBlankValue(value)) {
        return normalizedFilterValue === "";
      }
      if (isComputedError(value) || typeof value === "boolean") {
        return foldDeterministicText(normalizedValue) === foldDeterministicText(normalizedFilterValue);
      }
      const numericValue = numericComparableValue(value);
      const numericFilterValue = parseNumericValue(normalizedFilterValue);
      if (numericValue !== null && numericFilterValue !== null) {
        return numericValue === numericFilterValue;
      }
      return foldDeterministicText(normalizedValue) === foldDeterministicText(normalizedFilterValue);
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

function tokenizeNaturalText(value: string): string[] {
  return value.match(/\d+|\D+/g) ?? [value];
}

function compareNaturalToken(left: string, right: string): number {
  const leftNumeric = /^\d+$/.test(left);
  const rightNumeric = /^\d+$/.test(right);
  if (leftNumeric && rightNumeric) {
    const normalizedLeft = left.replace(/^0+(?=\d)/, "");
    const normalizedRight = right.replace(/^0+(?=\d)/, "");
    if (normalizedLeft.length !== normalizedRight.length) {
      return normalizedLeft.length - normalizedRight.length;
    }
    if (normalizedLeft !== normalizedRight) {
      return normalizedLeft < normalizedRight ? -1 : 1;
    }
    return 0;
  }

  if (left === right) {
    return 0;
  }
  return left < right ? -1 : 1;
}
