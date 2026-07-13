import { describe, expect, it } from "vitest";
import {
  normalizeExcelTableNameKey,
  validateExcelTableName,
  validateExcelTableNameForInterop
} from "./tableNames";

describe("Excel table names", () => {
  it.each(["_Sales", "\\Sales", "Équipe", "XFE1", "A1048577", "a".repeat(255)])(
    "accepts %s",
    (name) => expect(validateExcelTableName(name)).toEqual({
      valid: true,
      normalizedKey: normalizeExcelTableNameKey(name)
    })
  );

  it.each([
    ["", "empty"],
    ["   ", "empty"],
    [" Sales", "invalidCharacter"],
    ["Sales ", "invalidCharacter"],
    ["1Sales", "invalidCharacter"],
    ["Sales-2026", "invalidCharacter"],
    ["R", "reserved"],
    ["c", "reserved"],
    ["A1", "cellReference"],
    ["xfd1048576", "cellReference"],
    ["R1C1", "cellReference"],
    ["a".repeat(256), "tooLong"]
  ] as const)("rejects %s as %s", (name, reason) => {
    expect(validateExcelTableName(name)).toEqual({ valid: false, reason });
  });

  it("normalizes NFKC-equivalent names with locale-independent casing", () => {
    expect(normalizeExcelTableNameKey("Ｓａｌｅｓ")).toBe(normalizeExcelTableNameKey("sales"));
    expect(normalizeExcelTableNameKey("Cafe\u0301")).toBe(normalizeExcelTableNameKey("CAFÉ"));
    expect(normalizeExcelTableNameKey("I")).not.toBe(normalizeExcelTableNameKey("İ"));
  });

  it("preserves third-party cell-reference names without relaxing interactive validation", () => {
    expect(validateExcelTableName("T1")).toEqual({ valid: false, reason: "cellReference" });
    expect(validateExcelTableNameForInterop("T1")).toEqual({
      valid: true,
      normalizedKey: "t1"
    });
    expect(validateExcelTableNameForInterop("bad name")).toEqual({
      valid: false,
      reason: "invalidCharacter"
    });
  });
});
