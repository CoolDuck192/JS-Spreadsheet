import { describe, expect, it } from "vitest";
import { validateCellCandidate, validateCellValue } from "./validation";

describe("validation", () => {
  it("accepts only values from a configured list", () => {
    const rule = { type: "list", values: ["Open", "Closed"] } as const;

    expect(validateCellValue("Open", rule).valid).toBe(true);
    expect(validateCellValue("Blocked", rule)).toEqual({
      valid: false,
      message: "Choose one of: Open, Closed"
    });
  });

  it("validates numeric ranges while allowing blanks by default", () => {
    const rule = { type: "number", min: 1, max: 10 } as const;

    expect(validateCellValue("", rule).valid).toBe(true);
    expect(validateCellValue("5", rule).valid).toBe(true);
    expect(validateCellValue("0", rule)).toEqual({
      valid: false,
      message: "Enter a number between 1 and 10"
    });
    expect(validateCellValue("West", rule).valid).toBe(false);
  });

  it("validates text length ranges while allowing blanks by default", () => {
    const rule = { type: "textLength", min: 2, max: 5 } as const;

    expect(validateCellValue("", rule).valid).toBe(true);
    expect(validateCellValue("West", rule).valid).toBe(true);
    expect(validateCellValue("W", rule)).toEqual({
      valid: false,
      message: "Enter text between 2 and 5 characters"
    });
    expect(validateCellValue("Western", rule)).toEqual({
      valid: false,
      message: "Enter text between 2 and 5 characters"
    });
  });

  it("validates a numeric formula by its evaluated candidate", () => {
    const rule = { type: "number", min: 1, max: 10 } as const;

    expect(
      validateCellCandidate({ raw: "=5+5", parsed: "=5+5", evaluated: 10, formula: "=5+5" }, rule)
    ).toEqual({ valid: true });
    expect(
      validateCellCandidate({ raw: "=5+6", parsed: "=5+6", evaluated: 11, formula: "=5+6" }, rule)
    ).toEqual({ valid: false, message: "Enter a number between 1 and 10" });
  });

  it("uses parsed text for list and text-length candidates", () => {
    expect(
      validateCellCandidate(
        { raw: "'Open", parsed: "Open", evaluated: "Open" },
        { type: "list", values: ["Open", "Closed"] }
      )
    ).toEqual({ valid: true });
    expect(
      validateCellCandidate(
        { raw: "'West", parsed: "West", evaluated: "West" },
        { type: "textLength", min: 4, max: 4 }
      )
    ).toEqual({ valid: true });
    expect(
      validateCellCandidate(
        { raw: '=IF(TRUE,"Open","Closed")', parsed: '=IF(TRUE,"Open","Closed")', evaluated: "Open", formula: '=IF(TRUE,"Open","Closed")' },
        { type: "list", values: ["Open", "Closed"] }
      )
    ).toEqual({ valid: false, message: "Choose one of: Open, Closed" });
  });

  it("rejects formula errors with an explicit result message", () => {
    expect(
      validateCellCandidate(
        { raw: "=1/0", parsed: "=1/0", evaluated: { kind: "error", code: "#DIV/0!" }, formula: "=1/0" },
        { type: "number", min: 1 }
      )
    ).toEqual({ valid: false, message: "Formula evaluates to #DIV/0!" });
  });
});
