import type { CellRange, SheetModel } from "../types";
import { columnIndexToName, formatCellAddress, normalizeRange } from "./addressing";
import { DEFAULT_ROW_HEIGHT, clampColumnWidth, clampRowHeight } from "./sheetDimensions";

const COLUMN_CHARACTER_WIDTH = 7;
const COLUMN_PADDING = 27;
const ROW_LINE_HEIGHT = 18;
const ROW_PADDING = 10;
// The measurement constants above assume the grid's default 12px text; cells
// with an explicit fontSize (in points) scale relative to that.
const DEFAULT_CELL_FONT_PX = 12;
const PX_PER_POINT = 4 / 3;

function fontScaleFor(fontSize: number | undefined): number {
  if (!fontSize || fontSize <= 0) {
    return 1;
  }
  return (fontSize * PX_PER_POINT) / DEFAULT_CELL_FONT_PX;
}

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
      let width = textToAutoFitColumnWidth(columnIndexToName(column));
      for (let row = startRow; row <= endRow; row += 1) {
        if ((sheet.hiddenRows ?? {})[String(row)]) {
          continue;
        }
        const address = formatCellAddress({ row, column });
        const scale = fontScaleFor((sheet.formats ?? {})[address]?.fontSize);
        width = Math.max(width, textToAutoFitColumnWidth(getDisplayValue(address), scale));
      }

      return { column, width };
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
      let height = textToAutoFitRowHeight(String(row + 1));
      for (let column = startColumn; column <= endColumn; column += 1) {
        if ((sheet.hiddenColumns ?? {})[String(column)]) {
          continue;
        }
        const address = formatCellAddress({ row, column });
        const scale = fontScaleFor((sheet.formats ?? {})[address]?.fontSize);
        height = Math.max(height, textToAutoFitRowHeight(getDisplayValue(address), scale));
      }

      return { row, height };
    });
}

export function textToAutoFitColumnWidth(text: string, fontScale = 1): number {
  const longestLine = Math.max(1, ...splitDisplayLines(text).map((line) => line.length));
  return clampColumnWidth(Math.ceil(longestLine * COLUMN_CHARACTER_WIDTH * fontScale) + COLUMN_PADDING);
}

export function textToAutoFitRowHeight(text: string, fontScale = 1): number {
  const lineCount = Math.max(1, splitDisplayLines(text).length);
  if (lineCount <= 1) {
    return clampRowHeight(fontScale <= 1 ? DEFAULT_ROW_HEIGHT : Math.ceil(ROW_LINE_HEIGHT * fontScale) + ROW_PADDING);
  }
  return clampRowHeight(Math.ceil(lineCount * ROW_LINE_HEIGHT * fontScale) + ROW_PADDING);
}

function splitDisplayLines(text: string): string[] {
  return String(text ?? "").replace(/\r\n/g, "\n").replace(/\r/g, "\n").split("\n");
}

function clampIndex(index: number, length: number): number {
  return Math.min(Math.max(0, Math.floor(index)), Math.max(length - 1, 0));
}
