import { describe, expect, it } from "vitest";
import {
  extractFormulaReferences,
  rewriteFormulaForStructure,
  translateFormulaReferences
} from "./formulaReferences";

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

  it("rewrites absolute, mixed, and relative references for structural inserts", () => {
    expect(
      rewriteFormulaForStructure("=$A1+A$2+$A$3+B4", {
        formulaSheetName: "Data",
        editedSheetName: "Data",
        axis: "row",
        mode: "insert",
        index: 1,
        count: 2
      })
    ).toBe("=$A1+A$4+$A$5+B6");

    expect(
      rewriteFormulaForStructure("=$A1+B$2+$C$3+D4", {
        formulaSheetName: "Data",
        editedSheetName: "Data",
        axis: "column",
        mode: "insert",
        index: 1,
        count: 2
      })
    ).toBe("=$A1+D$2+$E$3+F4");
  });

  it("preserves function names, scientific notation, identifiers, and double-quoted text", () => {
    expect(
      rewriteFormulaForStructure('=LOG10(A1)+1E5+Q1_TOTAL+"A1 ""B2"" C3"+A2', {
        formulaSheetName: "Data",
        editedSheetName: "Data",
        axis: "row",
        mode: "insert",
        index: 0,
        count: 1
      })
    ).toBe('=LOG10(A2)+1E5+Q1_TOTAL+"A1 ""B2"" C3"+A3');
  });

  it("resolves quoted sheet names and doubled apostrophes from formulas on other sheets", () => {
    expect(
      rewriteFormulaForStructure("='Director''s Plan'!$A2+'Ops Q1'!B2+Sheet2!C2+A2", {
        formulaSheetName: "Summary",
        editedSheetName: "Director's Plan",
        axis: "row",
        mode: "insert",
        index: 1,
        count: 1
      })
    ).toBe("='Director''s Plan'!$A3+'Ops Q1'!B2+Sheet2!C2+A2");
  });

  it("rewrites local and same-sheet-qualified references without touching other sheets", () => {
    expect(
      rewriteFormulaForStructure("=A2+'Director''s Plan'!$B$2+Summary!C2", {
        formulaSheetName: "Director's Plan",
        editedSheetName: "Director's Plan",
        axis: "row",
        mode: "insert",
        index: 1,
        count: 1
      })
    ).toBe("=A3+'Director''s Plan'!$B$3+Summary!C2");
  });

  it("shrinks row ranges and invalidates only fully deleted references", () => {
    expect(
      rewriteFormulaForStructure("=SUM($A$2:$A$6)+SUM(B3:B6)+SUM(C3:C4)+D4", {
        formulaSheetName: "Data",
        editedSheetName: "Data",
        axis: "row",
        mode: "delete",
        index: 2,
        count: 2
      })
    ).toBe("=SUM($A$2:$A$4)+SUM(B3:B4)+SUM(#REF!)+#REF!");
  });

  it("shrinks column ranges and preserves qualified ranges on other sheets", () => {
    expect(
      rewriteFormulaForStructure("=SUM(B1:F1)+SUM(C2:F2)+SUM(C3:D3)+D4", {
        formulaSheetName: "Data",
        editedSheetName: "Data",
        axis: "column",
        mode: "delete",
        index: 2,
        count: 2
      })
    ).toBe("=SUM(B1:D1)+SUM(C2:D2)+SUM(#REF!)+#REF!");

    expect(
      rewriteFormulaForStructure("=SUM(Data!A1:C3)+SUM(A1:C3)", {
        formulaSheetName: "Summary",
        editedSheetName: "Data",
        axis: "column",
        mode: "insert",
        index: 1,
        count: 1
      })
    ).toBe("=SUM(Data!A1:D3)+SUM(A1:C3)");
  });
});
