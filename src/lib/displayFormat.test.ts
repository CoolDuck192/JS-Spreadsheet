import { describe, expect, it } from "vitest";
import { formatDisplayValue } from "./displayFormat";

describe("displayFormat", () => {
  it("formats numbers, currency, percent, and dates without changing raw values", () => {
    expect(formatDisplayValue("1234.5", { numberFormat: "number" })).toBe("1,234.5");
    expect(formatDisplayValue("1234.5", { numberFormat: "currency" })).toBe("$1,234.50");
    expect(formatDisplayValue("0.125", { numberFormat: "percent" })).toBe("12.5%");
    expect(formatDisplayValue("2026-06-28", { numberFormat: "date" })).toBe("Jun 28, 2026");
    expect(formatDisplayValue("46037", { numberFormat: "date" })).toBe("Jan 15, 2026");
    expect(formatDisplayValue("46037.5", { numberFormat: "dateTime" })).toBe("Jan 15, 2026, 12:00 PM");
  });

  it("renders Excel-compatible serials on both sides of the 1900 leap-day offset", () => {
    expect(formatDisplayValue("15", { numberFormat: "date" })).toBe("Jan 15, 1900");
    expect(formatDisplayValue("60", { numberFormat: "date" })).toBe("Feb 29, 1900");
    expect(formatDisplayValue("60.5", { numberFormat: "dateTime" })).toBe("Feb 29, 1900, 12:00 PM");
    expect(formatDisplayValue("61", { numberFormat: "date" })).toBe("Mar 1, 1900");
  });

  it("leaves text, blanks, errors, and general formatting alone", () => {
    expect(formatDisplayValue("West", { numberFormat: "currency" })).toBe("West");
    expect(formatDisplayValue("", { numberFormat: "number" })).toBe("");
    expect(formatDisplayValue("#DIV/0!", { numberFormat: "percent" })).toBe("#DIV/0!");
    expect(formatDisplayValue("1234.5", { numberFormat: "general" })).toBe("1234.5");
  });
});
