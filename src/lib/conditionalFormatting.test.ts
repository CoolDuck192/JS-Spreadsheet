import { describe, expect, it } from "vitest";
import type { ConditionalFormatRule } from "../types";
import { getConditionalDataBarForValue, getConditionalFormatForValue } from "./conditionalFormatting";

describe("conditionalFormatting", () => {
  it("matches numeric and text rules", () => {
    const greaterThan: ConditionalFormatRule = {
      id: "rule-1",
      range: { start: { row: 0, column: 0 }, end: { row: 2, column: 0 } },
      condition: { type: "greaterThan", value: "10" },
      format: { backgroundColor: "#fff1d6", textColor: "#8a4b00", bold: true }
    };
    const textContains: ConditionalFormatRule = {
      id: "rule-2",
      range: { start: { row: 0, column: 0 }, end: { row: 2, column: 0 } },
      condition: { type: "textContains", value: "risk" },
      format: { backgroundColor: "#fde8e8", textColor: "#9b1c1c" }
    };

    expect(getConditionalFormatForValue("12", [greaterThan])).toEqual(greaterThan.format);
    expect(getConditionalFormatForValue("8", [greaterThan])).toBeNull();
    expect(getConditionalFormatForValue("High risk", [textContains])).toEqual(textContains.format);
  });

  it("matches blank and nonblank rules", () => {
    const blank: ConditionalFormatRule = {
      id: "rule-1",
      range: { start: { row: 0, column: 0 }, end: { row: 2, column: 0 } },
      condition: { type: "blank" },
      format: { backgroundColor: "#fde8e8", textColor: "#9b1c1c" }
    };
    const notBlank: ConditionalFormatRule = {
      id: "rule-2",
      range: { start: { row: 0, column: 0 }, end: { row: 2, column: 0 } },
      condition: { type: "notBlank" },
      format: { backgroundColor: "#eaf7f2", textColor: "#17634a" }
    };

    expect(getConditionalFormatForValue("", [blank])).toEqual(blank.format);
    expect(getConditionalFormatForValue("   ", [blank])).toEqual(blank.format);
    expect(getConditionalFormatForValue("Ready", [blank])).toBeNull();
    expect(getConditionalFormatForValue("Ready", [notBlank])).toEqual(notBlank.format);
    expect(getConditionalFormatForValue("", [notBlank])).toBeNull();
  });

  it("matches duplicate and unique nonblank values in the rule range", () => {
    const duplicate: ConditionalFormatRule = {
      id: "rule-1",
      range: { start: { row: 0, column: 0 }, end: { row: 3, column: 0 } },
      condition: { type: "duplicate" },
      format: { backgroundColor: "#fde8e8", textColor: "#9b1c1c" }
    };
    const unique: ConditionalFormatRule = {
      id: "rule-2",
      range: { start: { row: 0, column: 0 }, end: { row: 3, column: 0 } },
      condition: { type: "unique" },
      format: { backgroundColor: "#eaf7f2", textColor: "#17634a" }
    };
    const getRuleValues = () => ["West", "East", "west", ""];

    expect(getConditionalFormatForValue("West", [duplicate], { getRuleValues })).toEqual(duplicate.format);
    expect(getConditionalFormatForValue("East", [duplicate], { getRuleValues })).toBeNull();
    expect(getConditionalFormatForValue("", [duplicate], { getRuleValues })).toBeNull();
    expect(getConditionalFormatForValue("East", [unique], { getRuleValues })).toEqual(unique.format);
    expect(getConditionalFormatForValue("West", [unique], { getRuleValues })).toBeNull();
    expect(getConditionalFormatForValue("", [unique], { getRuleValues })).toBeNull();
  });

  it("matches top and bottom numeric values in the rule range", () => {
    const top: ConditionalFormatRule = {
      id: "rule-1",
      range: { start: { row: 0, column: 0 }, end: { row: 4, column: 0 } },
      condition: { type: "top", count: 2 } as ConditionalFormatRule["condition"],
      format: { backgroundColor: "#fff1d6", textColor: "#8a4b00" }
    };
    const bottom: ConditionalFormatRule = {
      id: "rule-2",
      range: { start: { row: 0, column: 0 }, end: { row: 4, column: 0 } },
      condition: { type: "bottom", count: 2 } as ConditionalFormatRule["condition"],
      format: { backgroundColor: "#eaf7f2", textColor: "#17634a" }
    };
    const getRuleValues = () => ["10", "30", "20", "30", "n/a"];

    expect(getConditionalFormatForValue("30", [top], { getRuleValues })).toEqual(top.format);
    expect(getConditionalFormatForValue("20", [top], { getRuleValues })).toBeNull();
    expect(getConditionalFormatForValue("30", [bottom], { getRuleValues })).toBeNull();
    expect(getConditionalFormatForValue("20", [bottom], { getRuleValues })).toEqual(bottom.format);
    expect(getConditionalFormatForValue("10", [bottom], { getRuleValues })).toEqual(bottom.format);
    expect(getConditionalFormatForValue("n/a", [bottom], { getRuleValues })).toBeNull();
  });

  it("calculates data bar percentages from numeric rule values", () => {
    const dataBar: ConditionalFormatRule = {
      id: "rule-1",
      range: { start: { row: 0, column: 0 }, end: { row: 3, column: 0 } },
      condition: { type: "dataBar", color: "#2f7d9f" } as ConditionalFormatRule["condition"],
      format: {}
    };
    const getRuleValues = () => ["10", "20", "30", "n/a"];

    expect(getConditionalDataBarForValue("20", [dataBar], { getRuleValues })).toEqual({
      color: "#2f7d9f",
      percent: 67
    });
    expect(getConditionalDataBarForValue("n/a", [dataBar], { getRuleValues })).toBeNull();
  });

  it("calculates color scale backgrounds from numeric rule values", () => {
    const colorScale: ConditionalFormatRule = {
      id: "rule-1",
      range: { start: { row: 0, column: 0 }, end: { row: 3, column: 0 } },
      condition: { type: "colorScale", minColor: "#ffffff", maxColor: "#000000" } as ConditionalFormatRule["condition"],
      format: {}
    };
    const getRuleValues = () => ["10", "20", "30", "n/a"];

    expect(getConditionalFormatForValue("10", [colorScale], { getRuleValues })).toEqual({ backgroundColor: "#ffffff" });
    expect(getConditionalFormatForValue("20", [colorScale], { getRuleValues })).toEqual({ backgroundColor: "#808080" });
    expect(getConditionalFormatForValue("30", [colorScale], { getRuleValues })).toEqual({ backgroundColor: "#000000" });
    expect(getConditionalFormatForValue("n/a", [colorScale], { getRuleValues })).toBeNull();
  });
});
