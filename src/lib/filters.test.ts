import { describe, expect, it } from "vitest";
import type { SheetFilter } from "../types";
import { getVisibleRows, matchesFilterValue } from "./filters";

const BLANK_FILTER_VALUE = "\u0000js-spreadsheet:blank";
const EMPTY_RESULT_FILTER_VALUE = "\u0000js-spreadsheet:empty-result";

describe("filters", () => {
  it("matches text and numeric filter values", () => {
    expect(matchesFilterValue("West", { operator: "contains", value: "we" })).toBe(true);
    expect(matchesFilterValue("East", { operator: "equals", value: "West" })).toBe(false);
    expect(matchesFilterValue("12", { operator: "greaterThan", value: "10" })).toBe(true);
    expect(matchesFilterValue("8", { operator: "lessThan", value: "10" })).toBe(true);
  });

  it("does not coerce blank or formula-empty values to zero", () => {
    expect(matchesFilterValue("", { operator: "equals", value: "0" })).toBe(false);
    expect(matchesFilterValue(null, { operator: "equals", value: "0" })).toBe(false);
    expect(matchesFilterValue(0, { operator: "equals", value: "0" })).toBe(true);
    expect(matchesFilterValue("", { operator: "equals", value: "" })).toBe(true);
    expect(matchesFilterValue(null, { operator: "equals", value: "" })).toBe(true);
  });

  it("matches blank and formula-empty sentinels without collapsing either into zero", () => {
    expect(matchesFilterValue(null, { operator: "equals", value: BLANK_FILTER_VALUE })).toBe(true);
    expect(matchesFilterValue("", { operator: "equals", value: BLANK_FILTER_VALUE })).toBe(false);
    expect(matchesFilterValue("", { operator: "equals", value: EMPTY_RESULT_FILTER_VALUE })).toBe(true);
    expect(matchesFilterValue(null, { operator: "equals", value: EMPTY_RESULT_FILTER_VALUE })).toBe(false);
    expect(matchesFilterValue(0, { operator: "equals", value: BLANK_FILTER_VALUE })).toBe(false);
    expect(matchesFilterValue(0, { operator: "equals", value: EMPTY_RESULT_FILTER_VALUE })).toBe(false);
  });

  it("keeps typed AutoFilter sentinels stable through SheetFilter serialization", () => {
    const filter: SheetFilter = {
      id: "filter-typed-values",
      range: { start: { row: 0, column: 0 }, end: { row: 3, column: 0 } },
      column: 0,
      operator: "equals",
      value: BLANK_FILTER_VALUE,
      values: [BLANK_FILTER_VALUE, EMPTY_RESULT_FILTER_VALUE, "0"],
      hasHeader: true
    };

    expect(JSON.parse(JSON.stringify(filter))).toEqual(filter);
  });

  it("matches typed dates, booleans, and errors without zero coercion", () => {
    expect(matchesFilterValue(46037, { operator: "equals", value: "2026-01-15" })).toBe(true);
    expect(matchesFilterValue(46037, { operator: "greaterThan", value: "2026-01-14" })).toBe(true);
    expect(matchesFilterValue(true, { operator: "equals", value: "TRUE" })).toBe(true);
    expect(matchesFilterValue(false, { operator: "equals", value: "0" })).toBe(false);
    expect(matchesFilterValue({ kind: "error", code: "#DIV/0!" }, { operator: "equals", value: "#div/0!" })).toBe(true);
    expect(matchesFilterValue({ kind: "error", code: "#DIV/0!" }, { operator: "equals", value: "" })).toBe(false);
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
