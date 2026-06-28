import { describe, expect, it } from "vitest";
import { validateCellValue } from "./validation";

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
});
