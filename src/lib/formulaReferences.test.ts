import { describe, expect, it } from "vitest";
import { extractFormulaReferences, translateFormulaReferences } from "./formulaReferences";

describe("formulaReferences", () => {
  it("translates relative references while preserving absolute locks", () => {
    expect(translateFormulaReferences("=$A1+A$1+$A$1+A1", { rowOffset: 1, columnOffset: 1 })).toBe(
      "=$A2+B$1+$A$1+B2"
    );
  });

  it("translates both ends of a range reference", () => {
    expect(translateFormulaReferences("=SUM(A1:B2)", { rowOffset: 2, columnOffset: 1 })).toBe("=SUM(B3:C4)");
  });

  it("leaves references inside quoted strings unchanged", () => {
    expect(translateFormulaReferences('=IF(A1>0,"A1 ok","B2 no")', { rowOffset: 1, columnOffset: 0 })).toBe(
      '=IF(A2>0,"A1 ok","B2 no")'
    );
  });

  it("extracts unique local references while ignoring quoted and cross-sheet references", () => {
    expect(extractFormulaReferences('=SUM($A$1:B2)+C3+"D4"+Sheet2!E5+A1')).toEqual([
      {
        label: "A1:B2",
        range: {
          start: { row: 0, column: 0 },
          end: { row: 1, column: 1 }
        }
      },
      {
        label: "C3",
        range: {
          start: { row: 2, column: 2 },
          end: { row: 2, column: 2 }
        }
      },
      {
        label: "A1",
        range: {
          start: { row: 0, column: 0 },
          end: { row: 0, column: 0 }
        }
      }
    ]);
  });
});
