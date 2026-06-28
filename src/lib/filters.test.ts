import { describe, expect, it } from "vitest";
import type { SheetFilter } from "../types";
import { getVisibleRows, matchesFilterValue } from "./filters";

describe("filters", () => {
  it("matches text and numeric filter values", () => {
    expect(matchesFilterValue("West", { operator: "contains", value: "we" })).toBe(true);
    expect(matchesFilterValue("East", { operator: "equals", value: "West" })).toBe(false);
    expect(matchesFilterValue("12", { operator: "greaterThan", value: "10" })).toBe(true);
    expect(matchesFilterValue("8", { operator: "lessThan", value: "10" })).toBe(true);
  });

  it("keeps headers visible and hides non-matching rows inside filtered ranges", () => {
    const filter: SheetFilter = {
      id: "filter-1",
      range: { start: { row: 0, column: 0 }, end: { row: 3, column: 1 } },
      column: 0,
      operator: "equals",
      value: "West",
      hasHeader: true
    };
    const rows = [
      ["Region", "Sales"],
      ["West", "10"],
      ["East", "8"],
      ["West", "12"]
    ];

    expect(getVisibleRows(4, [filter], (row, column) => rows[row]?.[column] ?? "")).toEqual([0, 1, 3]);
  });

  it("matches any selected AutoFilter value", () => {
    const filter: SheetFilter = {
      id: "filter-1",
      range: { start: { row: 0, column: 0 }, end: { row: 4, column: 1 } },
      column: 0,
      operator: "equals",
      value: "West",
      values: ["East", "West"],
      hasHeader: true
    };
    const rows = [
      ["Region", "Sales"],
      ["West", "10"],
      ["East", "8"],
      ["North", "11"],
      ["West", "12"]
    ];

    expect(getVisibleRows(5, [filter], (row, column) => rows[row]?.[column] ?? "")).toEqual([0, 1, 2, 4]);
  });
});
