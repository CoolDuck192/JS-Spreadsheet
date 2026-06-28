import type { CellCoord, CellRange, SheetModel } from "../types";
import { formatCellAddress, normalizeRange } from "./addressing";

export type AutoSumPlan = {
  target: CellCoord;
  source: CellRange;
  formula: string;
};

export type AutoFunctionName = "SUM" | "AVERAGE" | "COUNT" | "MAX" | "MIN";

export function createAutoSumPlan(
  sheet: SheetModel,
  selection: CellRange,
  getDisplayValue: (address: string) => string,
  functionName: AutoFunctionName = "SUM"
): AutoSumPlan | null {
  const source = normalizeRange(selection);
  const isExplicitRange = source.start.row !== source.end.row || source.start.column !== source.end.column;

  if (isExplicitRange) {
    const target = targetForSelectedRange(sheet, source);
    return target ? planFor(target, source, functionName) : null;
  }

  const target = source.start;
  const inferredSource = inferAdjacentNumericRange(sheet, target, getDisplayValue);
  return inferredSource ? planFor(target, inferredSource, functionName) : null;
}

function targetForSelectedRange(sheet: SheetModel, source: CellRange): CellCoord | null {
  const rowCount = source.end.row - source.start.row + 1;
  const columnCount = source.end.column - source.start.column + 1;

  if (rowCount > 1 && columnCount === 1 && source.end.row + 1 < sheet.rowCount) {
    return { row: source.end.row + 1, column: source.start.column };
  }

  if (columnCount > 1 && rowCount === 1 && source.end.column + 1 < sheet.columnCount) {
    return { row: source.start.row, column: source.end.column + 1 };
  }

  if (source.end.row + 1 < sheet.rowCount) {
    return { row: source.end.row + 1, column: source.start.column };
  }

  if (source.end.column + 1 < sheet.columnCount) {
    return { row: source.start.row, column: source.end.column + 1 };
  }

  return null;
}

function inferAdjacentNumericRange(
  sheet: SheetModel,
  target: CellCoord,
  getDisplayValue: (address: string) => string
): CellRange | null {
  return (
    contiguousNumericRun(sheet, target, { row: -1, column: 0 }, getDisplayValue) ??
    contiguousNumericRun(sheet, target, { row: 0, column: -1 }, getDisplayValue)
  );
}

function contiguousNumericRun(
  sheet: SheetModel,
  target: CellCoord,
  delta: { row: number; column: number },
  getDisplayValue: (address: string) => string
): CellRange | null {
  const cells: CellCoord[] = [];
  let current = { row: target.row + delta.row, column: target.column + delta.column };

  while (isInsideSheet(sheet, current)) {
    const address = formatCellAddress(current);
    if (!isNumericDisplayValue(getDisplayValue(address))) {
      break;
    }

    cells.push(current);
    current = { row: current.row + delta.row, column: current.column + delta.column };
  }

  if (cells.length === 0) {
    return null;
  }

  return normalizeRange({ start: cells.at(-1)!, end: cells[0] });
}

function planFor(target: CellCoord, source: CellRange, functionName: AutoFunctionName): AutoSumPlan {
  return {
    target,
    source,
    formula: `=${functionName}(${formatRangeAddress(source)})`
  };
}

export function formatRangeAddress(range: CellRange): string {
  const normalized = normalizeRange(range);
  const start = formatCellAddress(normalized.start);
  const end = formatCellAddress(normalized.end);
  return start === end ? start : `${start}:${end}`;
}

function isInsideSheet(sheet: SheetModel, coord: CellCoord): boolean {
  return coord.row >= 0 && coord.row < sheet.rowCount && coord.column >= 0 && coord.column < sheet.columnCount;
}

function isNumericDisplayValue(value: string): boolean {
  const trimmed = value.trim();
  if (trimmed === "") {
    return false;
  }

  const normalized = trimmed.replace(/[$,]/g, "").replace(/%$/, "");
  return Number.isFinite(Number(normalized));
}
