import { describe, expect, it } from "vitest";
import { findVisibleRange, measureAxis } from "./axis";

describe("axis viewport primitives", () => {
  it("measures variable-size items while retaining source indexes and keys", () => {
    const measurements = measureAxis(
      5,
      (index) => `column-${index}`,
      (index) => [40, 80, 0, 120, 60][index],
      (index) => index === 2
    );

    expect(measurements).toEqual([
      { index: 0, key: "column-0", size: 40, start: 0, end: 40 },
      { index: 1, key: "column-1", size: 80, start: 40, end: 120 },
      { index: 3, key: "column-3", size: 120, start: 120, end: 240 },
      { index: 4, key: "column-4", size: 60, start: 240, end: 300 }
    ]);
  });

  it("finds the visible half-open slice with binary-search boundary semantics", () => {
    const measurements = measureAxis(6, (index) => index, () => 100);

    expect(findVisibleRange(measurements, 150, 350, 0)).toEqual({ first: 1, last: 4 });
    expect(findVisibleRange(measurements, 100, 200, 0)).toEqual({ first: 0, last: 3 });
    expect(findVisibleRange(measurements, 10_000, 10_100, 0)).toEqual({ first: 6, last: 6 });
  });

  it("clamps overscan at both ends without mutating the measurements", () => {
    const measurements = measureAxis(8, (index) => index, (index) => 20 + index);
    const snapshot = measurements.map((measurement) => ({ ...measurement }));

    expect(findVisibleRange(measurements, 0, 10, 3)).toEqual({ first: 0, last: 4 });
    expect(findVisibleRange(measurements, 1_000, 1_010, 3)).toEqual({ first: 5, last: 8 });
    expect(measurements).toEqual(snapshot);
  });

  it("handles an empty axis", () => {
    expect(findVisibleRange([], 0, 500, 4)).toEqual({ first: 0, last: 0 });
  });
});
