import { describe, expect, it } from "vitest";
import { parseCellInput } from "./parseCellInput";

describe("parseCellInput", () => {
  it("classifies blank, formula, typed, escaped text, and ambiguous input", () => {
    expect(parseCellInput("")).toMatchObject({ kind: "blank", stored: null });
    expect(parseCellInput("=A1*2")).toMatchObject({ kind: "formula", stored: "=A1*2", formula: "=A1*2" });
    expect(parseCellInput("1.25e3")).toMatchObject({ kind: "number", stored: 1250 });
    expect(parseCellInput("FALSE")).toMatchObject({ kind: "boolean", stored: false });
    expect(parseCellInput("2026-01-15")).toMatchObject({
      kind: "date",
      stored: 46037,
      inferredNumberFormat: "date"
    });
    expect(parseCellInput("2026-01-15T12:00:00Z")).toMatchObject({
      kind: "dateTime",
      stored: 46037.5,
      inferredNumberFormat: "dateTime"
    });
    expect(parseCellInput("'00123")).toMatchObject({ kind: "text", stored: "00123" });
    expect(parseCellInput("1,23")).toMatchObject({ kind: "text", stored: "1,23" });
  });

  it("parses only finite numbers with unambiguous decimal or US grouping", () => {
    expect(parseCellInput("-12.5")).toMatchObject({ kind: "number", stored: -12.5 });
    expect(parseCellInput(".25")).toMatchObject({ kind: "number", stored: 0.25 });
    expect(parseCellInput("1,234.5")).toMatchObject({ kind: "number", stored: 1234.5 });

    for (const raw of ["NaN", "Infinity", "0x10", "12,34", "1 234", "1.2.3"]) {
      expect(parseCellInput(raw)).toMatchObject({ kind: "text", stored: raw });
    }
  });

  it("recognizes booleans case-insensitively without coercing other words", () => {
    expect(parseCellInput("true")).toMatchObject({ kind: "boolean", stored: true });
    expect(parseCellInput("False")).toMatchObject({ kind: "boolean", stored: false });
    expect(parseCellInput("truthy")).toMatchObject({ kind: "text", stored: "truthy" });
  });
});
