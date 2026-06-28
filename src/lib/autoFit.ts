import type { CellRange, SheetModel } from "../types";
import { columnIndexToName, formatCellAddress, normalizeRange } from "./addressing";
import { DEFAULT_ROW_HEIGHT, clampColumnWidth, clampRowHeight } from "./sheetDimensions";

const COLUMN_CHARACTER_WIDTH = 7;
const COLUMN_PADDING = 27;
const ROW_LINE_HEIGHT = 18;
const ROW_PADDING = 10;

export type AutoFitColumnSize = {
  column: number;
  width: number;
};

export type AutoFitRowSize = {
  row: number;
  height: number;
};

export function createAutoFitColumnPlan(
  sheet: SheetModel,
  range: CellRange,
  getDisplayValue: (address: string) => string
): AutoFitColumnSize[] {
  const normalized = normalizeRange(range);
  const startColumn = clampIndex(normalized.start.column, sheet.columnCount);
  const endColumn = clampIndex(normalized.end.column, sheet.columnCount);
  const startRow = clampIndex(normalized.start.row, sheet.rowCount);
  const endRow = clampIndex(normalized.end.row, sheet.rowCount);

  if (startColumn > endColumn || startRow > endRow) {
    return [];
  }

  return Array.from({ length: endColumn - startColumn + 1 }, (_, index) => startColumn + index)
    .filter((column) => !(sheet.hiddenColumns ?? {})[String(column)])
    .map((column) => {
      const candidates = [columnIndexToName(column)];
      for (let row = startRow; row <= endRow; row += 1) {
        if ((sheet.hiddenRows ?? {})[String(row)]) {
          continue;
        }
        candidates.push(getDisplayValue(formatCellAddress({ row, column })));
      }

      return { column, width: Math.max(...candidates.map(textToAutoFitColumnWidth)) };
    });
}

export function createAutoFitRowPlan(
  sheet: SheetModel,
  range: CellRange,
  getDisplayValue: (address: string) => string
): AutoFitRowSize[] {
  const normalized = normalizeRange(range);
  const startColumn = clampIndex(normalized.start.column, sheet.columnCount);
  const endColumn = clampIndex(normalized.end.column, sheet.columnCount);
  const startRow = clampIndex(normalized.start.row, sheet.rowCount);
  const endRow = clampIndex(normalized.end.row, sheet.rowCount);

  if (startColumn > endColumn || startRow > endRow) {
    return [];
  }

  return Array.from({ length: endRow - startRow + 1 }, (_, index) => startRow + index)
    .filter((row) => !(sheet.hiddenRows ?? {})[String(row)])
    .map((row) => {
      const candidates = [String(row + 1)];
      for (let column = startColumn; column <= endColumn; column += 1) {
        if ((sheet.hiddenColumns ?? {})[String(column)]) {
          continue;
        }
        candidates.push(getDisplayValue(formatCellAddress({ row, column })));
      }

      return { row, height: Math.max(...candidates.map(textToAutoFitRowHeight)) };
    });
}

export function textToAutoFitColumnWidth(text: string): number {
  const longestLine = Math.max(1, ...splitDisplayLines(text).map((line) => line.length));
  return clampColumnWidth(longestLine * COLUMN_CHARACTER_WIDTH + COLUMN_PADDING);
}

export function textToAutoFitRowHeight(text: string): number {
  const lineCount = Math.max(1, splitDisplayLines(text).length);
  return clampRowHeight(lineCount <= 1 ? DEFAULT_ROW_HEIGHT : lineCount * ROW_LINE_HEIGHT + ROW_PADDING);
}

function splitDisplayLines(text: string): string[] {
  return String(text ?? "").replace(/\r\n/g, "\n").replace(/\r/g, "\n").split("\n");
}

function clampIndex(index: number, length: number): number {
  return Math.min(Math.max(0, Math.floor(index)), Math.max(length - 1, 0));
}
