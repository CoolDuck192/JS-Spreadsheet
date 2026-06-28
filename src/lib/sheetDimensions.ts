export const DEFAULT_COLUMN_WIDTH = 96;
export const DEFAULT_ROW_HEIGHT = 28;
export const MIN_COLUMN_WIDTH = 56;
export const MAX_COLUMN_WIDTH = 360;
export const MIN_ROW_HEIGHT = 22;
export const MAX_ROW_HEIGHT = 96;

export function clampColumnWidth(width: number): number {
  return clampDimension(width, MIN_COLUMN_WIDTH, MAX_COLUMN_WIDTH);
}

export function clampRowHeight(height: number): number {
  return clampDimension(height, MIN_ROW_HEIGHT, MAX_ROW_HEIGHT);
}

function clampDimension(value: number, min: number, max: number): number {
  if (!Number.isFinite(value)) {
    return min;
  }
  return Math.min(max, Math.max(min, Math.round(value)));
}
