export type AxisKey = string | number;

export type AxisMeasurement<TKey extends AxisKey = AxisKey> = {
  index: number;
  key: TKey;
  size: number;
  start: number;
  end: number;
};

export type AxisRange = {
  first: number;
  last: number;
};

export function measureAxis<TKey extends AxisKey>(
  count: number,
  getKey: (index: number) => TKey,
  getSize: (index: number) => number,
  isHidden: (index: number) => boolean = () => false
): AxisMeasurement<TKey>[] {
  const measurements: AxisMeasurement<TKey>[] = [];
  let cursor = 0;
  for (let index = 0; index < count; index += 1) {
    if (isHidden(index)) {
      continue;
    }
    const requestedSize = getSize(index);
    const size = Number.isFinite(requestedSize) ? Math.max(0, requestedSize) : 0;
    measurements.push({
      index,
      key: getKey(index),
      size,
      start: cursor,
      end: cursor + size
    });
    cursor += size;
  }
  return measurements;
}

export function findVisibleRange(
  measurements: readonly AxisMeasurement[],
  start: number,
  end: number,
  overscan: number
): AxisRange {
  if (measurements.length === 0) {
    return { first: 0, last: 0 };
  }
  const safeOverscan = Math.max(0, Math.floor(overscan));
  const first = Math.max(0, lowerBoundByEnd(measurements, start) - safeOverscan);
  const last = Math.min(measurements.length, upperBoundByStart(measurements, end) + safeOverscan);
  return { first, last: Math.max(first, last) };
}

function lowerBoundByEnd(measurements: readonly AxisMeasurement[], target: number): number {
  let low = 0;
  let high = measurements.length;
  while (low < high) {
    const middle = (low + high) >> 1;
    if (measurements[middle].end < target) {
      low = middle + 1;
    } else {
      high = middle;
    }
  }
  return low;
}

function upperBoundByStart(measurements: readonly AxisMeasurement[], target: number): number {
  let low = 0;
  let high = measurements.length;
  while (low < high) {
    const middle = (low + high) >> 1;
    if (measurements[middle].start <= target) {
      low = middle + 1;
    } else {
      high = middle;
    }
  }
  return low;
}
