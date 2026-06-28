import { describe, expect, it } from "vitest";
import { formatDisplayValue } from "./displayFormat";

describe("displayFormat", () => {
  it("formats numbers, currency, percent, and dates without changing raw values", () => {
    expect(formatDisplayValue("1234.5", { numberFormat: "number" })).toBe("1,234.5");
    expect(formatDisplayValue("1234.5", { numberFormat: "currency" })).toBe("$1,234.50");
    expect(formatDisplayValue("0.125", { numberFormat: "percent" })).toBe("12.5%");
    expect(formatDisplayValue("2026-06-28", { numberFormat: "date" })).toBe("Jun 28, 2026");
  });

  it("leaves text, blanks, errors, and general formatting alone", () => {
    expect(formatDisplayValue("West", { numberFormat: "currency" })).toBe("West");
    expect(formatDisplayValue("", { numberFormat: "number" })).toBe("");
    expect(formatDisplayValue("#DIV/0!", { numberFormat: "percent" })).toBe("#DIV/0!");
    expect(formatDisplayValue("1234.5", { numberFormat: "general" })).toBe("1234.5");
  });
});
