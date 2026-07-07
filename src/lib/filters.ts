import type { FilterOperator, SheetFilter } from "../types";
import { normalizeRange } from "./addressing";

type FilterMatcher = {
  operator: FilterOperator;
  value: string;
  values?: readonly string[];
};

export function matchesFilterValue(value: string, filter: FilterMatcher): boolean {
  const normalizedValue = value.trim();
  const normalizedFilterValue = filter.value.trim();

  switch (filter.operator) {
    case "contains":
      return normalizedValue.toLocaleLowerCase().includes(normalizedFilterValue.toLocaleLowerCase());
    case "equals": {
      const numericValue = parseNumericValue(normalizedValue);
      const numericFilterValue = parseNumericValue(normalizedFilterValue);
      if (numericValue !== null && numericFilterValue !== null) {
        return numericValue === numericFilterValue;
      }
      return normalizedValue.toLocaleLowerCase() === normalizedFilterValue.toLocaleLowerCase();
    }
    case "greaterThan": {
      const numericValue = parseNumericValue(normalizedValue);
      const numericFilterValue = parseNumericValue(normalizedFilterValue);
      return numericValue !== null && numericFilterValue !== null && numericValue > numericFilterValue;
    }
    case "lessThan": {
      const numericValue = parseNumericValue(normalizedValue);
      const numericFilterValue = parseNumericValue(normalizedFilterValue);
      return numericValue !== null && numericFilterValue !== null && numericValue < numericFilterValue;
    }
  }
}

export function getVisibleRows(
  rowCount: number,
  filters: readonly SheetFilter[],
  getCellDisplayValue: (row: number, column: number) => string
): number[] {
  const rows: number[] = [];

  for (let row = 0; row < rowCount; row += 1) {
    if (filters.every((filter) => isRowVisibleForFilter(row, filter, getCellDisplayValue))) {
      rows.push(row);
    }
  }

  return rows;
}

export function isRowVisibleForFilter(
  row: number,
  filter: SheetFilter,
  getCellDisplayValue: (row: number, column: number) => string
): boolean {
  const range = normalizeRange(filter.range);
  if (row < range.start.row || row > range.end.row) {
    return true;
  }

  if (filter.hasHeader !== false && row === range.start.row) {
    return true;
  }

  const displayValue = getCellDisplayValue(row, filter.column);
  if (filter.values && filter.values.length > 0) {
    return filter.values.some((value) => matchesFilterValue(displayValue, { ...filter, value }));
  }

  return matchesFilterValue(displayValue, filter);
}

function parseNumericValue(value: string): number | null {
  const parsed = Number(value.replace(/[$,%]/g, "").replace(/,/g, ""));
  return Number.isFinite(parsed) ? parsed : null;
}
