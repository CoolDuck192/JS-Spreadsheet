import { describe, expect, it } from "vitest";
import {
  extractFormulaReferences,
  rewriteFormulaForRectangularRowEdit,
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

  it("preserves Unicode identifier boundaries while rewriting a standalone reference", () => {
    expect(
      rewriteFormulaForStructure("=ΔA1+A1Δ+e\u0301A1+A1\u0301+A1١+A1", {
        formulaSheetName: "Data",
        editedSheetName: "Data",
        axis: "row",
        mode: "insert",
        index: 0,
        count: 1
      })
    ).toBe("=ΔA1+A1Δ+e\u0301A1+A1\u0301+A1١+A2");
  });

  it("rewrites an unquoted Unicode sheet qualifier that targets the edited sheet", () => {
    expect(
      rewriteFormulaForStructure("=Δ!A1+A1", {
        formulaSheetName: "Summary",
        editedSheetName: "Δ",
        axis: "row",
        mode: "insert",
        index: 0,
        count: 1
      })
    ).toBe("=Δ!A2+A1");
  });

  it("preserves an unquoted Unicode qualifier that targets a different sheet", () => {
    expect(
      rewriteFormulaForStructure("=Ω!A1+A1", {
        formulaSheetName: "Summary",
        editedSheetName: "Δ",
        axis: "row",
        mode: "insert",
        index: 0,
        count: 1
      })
    ).toBe("=Ω!A1+A1");
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

  it("rewrites only the table-column slice of rectangular row inserts", () => {
    const context = {
      formulaSheetId: "Data",
      editedSheetId: "Data",
      tableColumnStart: 1,
      tableColumnEnd: 2,
      row: 2,
      count: 1,
      operation: "insert" as const,
      sheetBounds: { rowCount: 100, columnCount: 26 }
    };
    expect(rewriteFormulaForRectangularRowEdit("=SUM(B2:C4)+$B$4+D4", context)).toEqual({
      ok: true,
      formula: "=SUM(B2:C5)+$B$5+D4"
    });
    expect(rewriteFormulaForRectangularRowEdit("=SUM(A2:C4)", context)).toEqual({
      ok: true,
      formula: "=SUM((A2:A4,B2:C5))"
    });
  });

  it("shrinks or invalidates affected table-column slices on deletion", () => {
    const context = {
      formulaSheetId: "Data",
      editedSheetId: "Data",
      tableColumnStart: 1,
      tableColumnEnd: 2,
      row: 1,
      count: 2,
      operation: "delete" as const,
      sheetBounds: { rowCount: 100, columnCount: 26 }
    };
    expect(rewriteFormulaForRectangularRowEdit("=SUM(B2:C4)+B2+B4", context)).toEqual({
      ok: true,
      formula: "=SUM(B2:C2)+#REF!+B2"
    });
  });

  it("preserves quoted text, function names, scientific notation, and other sheets", () => {
    expect(rewriteFormulaForRectangularRowEdit(
      '=LOG10(B2)+1E10+"B2"+\'Rates 2026\'!B2',
      {
        formulaSheetId: "Data",
        editedSheetId: "Data",
        tableColumnStart: 1,
        tableColumnEnd: 2,
        row: 1,
        count: 1,
        operation: "insert",
        sheetBounds: { rowCount: 100, columnCount: 26 }
      }
    )).toEqual({
      ok: true,
      formula: '=LOG10(B3)+1E10+"B2"+\'Rates 2026\'!B2'
    });
  });

  it("does not reject external- or 3-D-shaped text inside formula string literals", () => {
    expect(rewriteFormulaForRectangularRowEdit(
      '=IF(B2,"\'[Book.xlsx]Data\'!B2 and \'Jan\'\'A\':\'Mar\'\'B\'!B2 and ""[escaped]""","")',
      {
        formulaSheetId: "Data",
        editedSheetId: "Data",
        tableColumnStart: 1,
        tableColumnEnd: 2,
        row: 1,
        count: 1,
        operation: "insert",
        sheetBounds: { rowCount: 100, columnCount: 26 }
      }
    )).toEqual({
      ok: true,
      formula: '=IF(B3,"\'[Book.xlsx]Data\'!B2 and \'Jan\'\'A\':\'Mar\'\'B\'!B2 and ""[escaped]""","")'
    });
  });

  it("preserves structured table references while rewriting nearby cell references", () => {
    const context = {
      formulaSheetId: "Data",
      editedSheetId: "Data",
      tableColumnStart: 1,
      tableColumnEnd: 2,
      row: 1,
      count: 1,
      operation: "insert" as const,
      sheetBounds: { rowCount: 100, columnCount: 26 }
    };

    expect(rewriteFormulaForRectangularRowEdit("=SUM(Sales[Quantity])+B2", context)).toEqual({
      ok: true,
      formula: "=SUM(Sales[Quantity])+B3"
    });
    expect(rewriteFormulaForRectangularRowEdit("=SUM([@Quantity])+C2", context)).toEqual({
      ok: true,
      formula: "=SUM([@Quantity])+C3"
    });
  });

  it.each([
    "Sales[B2]",
    "[@B2]",
    "Sales[[#Data],[B2]]"
  ])("does not rewrite cell-shaped labels inside structured reference %s", (structuredReference) => {
    const context = {
      formulaSheetId: "Data",
      editedSheetId: "Data",
      tableColumnStart: 1,
      tableColumnEnd: 2,
      row: 1,
      count: 1,
      operation: "insert" as const,
      sheetBounds: { rowCount: 100, columnCount: 26 }
    };

    expect(rewriteFormulaForRectangularRowEdit(`=SUM(${structuredReference})+C2`, context)).toEqual({
      ok: true,
      formula: `=SUM(${structuredReference})+C3`
    });
  });

  it("rejects external and 3-D references with a typed issue", () => {
    const context = {
      formulaSheetId: "Data",
      editedSheetId: "Data",
      tableColumnStart: 1,
      tableColumnEnd: 2,
      row: 1,
      count: 1,
      operation: "insert" as const,
      sheetBounds: { rowCount: 100, columnCount: 26 }
    };
    expect(rewriteFormulaForRectangularRowEdit("='[Book.xlsx]Data'!B2", context))
      .toMatchObject({ ok: false, issue: { code: "TABLE_FORMULA_REFERENCE_UNSUPPORTED" } });
    expect(rewriteFormulaForRectangularRowEdit("=[Book.xlsx]Data!B2", context))
      .toMatchObject({ ok: false, issue: { code: "TABLE_FORMULA_REFERENCE_UNSUPPORTED" } });
    expect(rewriteFormulaForRectangularRowEdit("=SUM(Jan:Mar!B2)", context))
      .toMatchObject({ ok: false, issue: { code: "TABLE_FORMULA_REFERENCE_UNSUPPORTED" } });
    expect(rewriteFormulaForRectangularRowEdit("=SUM('Jan''A':'Mar''B'!B2)", context))
      .toMatchObject({ ok: false, issue: { code: "TABLE_FORMULA_REFERENCE_UNSUPPORTED" } });
  });
});
