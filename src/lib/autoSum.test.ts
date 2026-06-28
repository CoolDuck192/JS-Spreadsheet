import { describe, expect, it } from "vitest";
import type { CellRange, SheetModel } from "../types";
import { createAutoSumPlan } from "./autoSum";
import { createBlankWorkbook, getActiveSheet, setCellContent } from "./workbook";

const range = (start: string, end = start): CellRange => ({
  start: addressToCoord(start),
  end: addressToCoord(end)
});

describe("autoSum", () => {
  it("plans a SUM formula below a selected column range", () => {
    const sheet = getActiveSheet(createBlankWorkbook());
    const plan = createAutoSumPlan(sheet, range("A1", "A3"), () => "");

    expect(plan).toEqual({
      target: addressToCoord("A4"),
      source: range("A1", "A3"),
      formula: "=SUM(A1:A3)"
    });
  });

  it.each([
    ["SUM", "=SUM(A1:A3)"],
    ["AVERAGE", "=AVERAGE(A1:A3)"],
    ["COUNT", "=COUNT(A1:A3)"],
    ["MAX", "=MAX(A1:A3)"],
    ["MIN", "=MIN(A1:A3)"]
  ] as const)("plans a %s formula for a selected range when requested", (functionName, formula) => {
    const sheet = getActiveSheet(createBlankWorkbook());
    const plan = createAutoSumPlan(sheet, range("A1", "A3"), () => "", functionName);

    expect(plan).toEqual({
      target: addressToCoord("A4"),
      source: range("A1", "A3"),
      formula
    });
  });

  it("plans a SUM formula to the right of a selected row range", () => {
    const sheet = getActiveSheet(createBlankWorkbook());
    const plan = createAutoSumPlan(sheet, range("B2", "D2"), () => "");

    expect(plan).toEqual({
      target: addressToCoord("E2"),
      source: range("B2", "D2"),
      formula: "=SUM(B2:D2)"
    });
  });

  it("infers contiguous numbers above a single active cell", () => {
    const sheet = sheetWithValues({
      A1: "10",
      A2: "20",
      A3: "30"
    });
    const plan = createAutoSumPlan(sheet, range("A4"), (address) => String(sheet.cells[address] ?? ""));

    expect(plan).toEqual({
      target: addressToCoord("A4"),
      source: range("A1", "A3"),
      formula: "=SUM(A1:A3)"
    });
  });

  it("infers contiguous numbers to the left when there are no numbers above", () => {
    const sheet = sheetWithValues({
      A2: "5",
      B2: "7",
      C2: "9"
    });
    const plan = createAutoSumPlan(sheet, range("D2"), (address) => String(sheet.cells[address] ?? ""));

    expect(plan).toEqual({
      target: addressToCoord("D2"),
      source: range("A2", "C2"),
      formula: "=SUM(A2:C2)"
    });
  });

  it("returns null when a single active cell has no nearby numeric run", () => {
    const sheet = sheetWithValues({ A1: "Name" });

    expect(createAutoSumPlan(sheet, range("B2"), (address) => String(sheet.cells[address] ?? ""))).toBeNull();
  });
});

function sheetWithValues(values: Record<string, string>): SheetModel {
  let workbook = createBlankWorkbook();
  for (const [address, value] of Object.entries(values)) {
    workbook = setCellContent(workbook, workbook.activeSheetId, address, value);
  }
  return getActiveSheet(workbook);
}

function addressToCoord(address: string) {
  const match = address.match(/^([A-Z]+)(\d+)$/);
  if (!match) {
    throw new Error(`Invalid test address: ${address}`);
  }

  return {
    row: Number(match[2]) - 1,
    column: match[1].split("").reduce((total, character) => total * 26 + character.charCodeAt(0) - 64, 0) - 1
  };
}
