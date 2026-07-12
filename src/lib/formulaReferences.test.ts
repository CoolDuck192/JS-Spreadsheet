import { describe, expect, it } from "vitest";
import {
  extractFormulaReferences,
  rewriteFormulaForRectangularRowMove,
  rewriteFormulaForRectangularRowEdit,
  rewriteFormulaForStructure,
  translateFormulaReferences,
  translateFormulaRowsWithinColumns
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

  it("preserves single-quoted sheet names that resemble cell references", () => {
    expect(translateFormulaReferences(
      "='A1'!B2+'Q1 data'!C3+A1",
      { rowOffset: 1, columnOffset: 0 }
    )).toBe("='A1'!B3+'Q1 data'!C4+A2");
  });

  it("preserves external workbook references while translating same-sheet rows", () => {
    expect(translateFormulaRowsWithinColumns(
      "=[Book.xlsx]Sheet1!A3+A3",
      { rowOffset: 1, formulaSheetName: "Sheet1", columnStart: 0, columnEnd: 1 }
    )).toBe("=[Book.xlsx]Sheet1!A3+A4");
  });

  it("emits REF for negative range endpoints during bounded row translation", () => {
    expect(translateFormulaRowsWithinColumns(
      "=SUM(A1:B3)+Sheet1!A1",
      { rowOffset: -2, formulaSheetName: "Sheet1", columnStart: 0, columnEnd: 1 }
    )).toBe("=SUM(#REF!)+#REF!");
  });

  it("preserves mixed in-table and out-of-table ranges during bounded row translation", () => {
    expect(translateFormulaRowsWithinColumns(
      "=SUM($D3:A$3)",
      { rowOffset: 1, formulaSheetName: "Sheet1", columnStart: 0, columnEnd: 1 }
    )).toBe("=SUM($D3:A$3)");
  });

  it("maps totals-row moves down while preserving locks and unrelated references", () => {
    expect(rewriteFormulaForRectangularRowMove(
      "=$A$6+A$7+$B8+C6+Other!A6",
      {
        formulaSheetId: "Data",
        editedSheetId: "Data",
        tableColumnStart: 0,
        tableColumnEnd: 1,
        sourceRow: 5,
        targetRow: 7
      }
    )).toEqual({
      ok: true,
      formula: "=$A$8+A$6+$B7+C6+Other!A6"
    });
  });

  it("maps totals-row moves up on same- and cross-sheet dependents", () => {
    const context = {
      formulaSheetId: "Summary",
      editedSheetId: "Data",
      tableColumnStart: 0,
      tableColumnEnd: 1,
      sourceRow: 5,
      targetRow: 3
    };

    expect(rewriteFormulaForRectangularRowMove(
      "=Data!$A$6+Data!A$4+Data!$B5+A6+Other!A6",
      context
    )).toEqual({
      ok: true,
      formula: "=Data!$A$4+Data!A$5+Data!$B6+A6+Other!A6"
    });
  });

  it.each([
    {
      label: "up during growth",
      formula: "=A9+$A9+A$9+$A$9+A6+$A$6+SUM(A9:B10)+SUM(A$9:B$10)",
      context: {
        formulaSheetId: "Data",
        editedSheetId: "Data",
        tableColumnStart: 0,
        tableColumnEnd: 1,
        sourceRow: 5,
        targetRow: 7,
        formulaCell: { row: 6, column: 0 }
      },
      expected: "=A8+$A8+A$9+$A$9+A8+$A$8+SUM(A8:B9)+SUM(A$9:B$10)"
    },
    {
      label: "down during shrink",
      formula: "=A2+$A2+A$2+$A$2+A6+$A$6+SUM(A1:B2)+SUM(A$1:B$2)",
      context: {
        formulaSheetId: "Data",
        editedSheetId: "Data",
        tableColumnStart: 0,
        tableColumnEnd: 1,
        sourceRow: 5,
        targetRow: 3,
        formulaCell: { row: 3, column: 0 }
      },
      expected: "=A3+$A3+A$2+$A$2+A4+$A$4+SUM(A2:B3)+SUM(A$1:B$2)"
    }
  ])("uses the formula cell location when it moves $label", ({ formula, context, expected }) => {
    expect(rewriteFormulaForRectangularRowMove(formula, context)).toEqual({
      ok: true,
      formula: expected
    });
  });

  it.each([
    ["quoted external workbook", "='[Book.xlsx]Data'!B2"],
    ["unquoted external workbook", "=[Book.xlsx]Data!B2"],
    ["unquoted 3-D", "=SUM(Jan:Mar!B2)"],
    ["quoted 3-D", "=SUM('Jan''A':'Mar''B'!B2)"]
  ])("keeps a physically moved %s formula byte-identical when no local reference changes", (_label, formula) => {
    const context = {
      formulaSheetId: "Data",
      editedSheetId: "Data",
      tableColumnStart: 0,
      tableColumnEnd: 1,
      sourceRow: 5,
      targetRow: 7,
      formulaCell: { row: 6, column: 0 }
    };

    expect(rewriteFormulaForRectangularRowMove(formula, context)).toEqual({
      ok: true,
      formula
    });
    expect(rewriteFormulaForRectangularRowMove(`${formula}+A7`, context))
      .toMatchObject({ ok: false, issue: { code: "TABLE_FORMULA_REFERENCE_UNSUPPORTED" } });
  });

  it("merges final range images after permutation and moved-formula translation", () => {
    expect(rewriteFormulaForRectangularRowMove("=SUM(A7:B9)", {
      formulaSheetId: "Data",
      editedSheetId: "Data",
      tableColumnStart: 0,
      tableColumnEnd: 1,
      sourceRow: 5,
      targetRow: 7,
      formulaCell: { row: 6, column: 0 }
    })).toEqual({ ok: true, formula: "=SUM(A6:B8)" });

    expect(rewriteFormulaForRectangularRowMove("=SUM(A2:B4)", {
      formulaSheetId: "Data",
      editedSheetId: "Data",
      tableColumnStart: 0,
      tableColumnEnd: 1,
      sourceRow: 5,
      targetRow: 3,
      formulaCell: { row: 3, column: 0 }
    })).toEqual({ ok: true, formula: "=SUM(A3:B5)" });
  });

  it("rejects a final range image separated by a locked outside endpoint", () => {
    expect(rewriteFormulaForRectangularRowMove("=SUM(A7:B$9)", {
      formulaSheetId: "Data",
      editedSheetId: "Data",
      tableColumnStart: 0,
      tableColumnEnd: 1,
      sourceRow: 5,
      targetRow: 7,
      formulaCell: { row: 6, column: 0 }
    })).toMatchObject({
      ok: false,
      issue: { code: "TABLE_FORMULA_REFERENCE_UNSUPPORTED" }
    });
  });

  it("keeps contiguous row-range images representable in both move directions", () => {
    expect(rewriteFormulaForRectangularRowMove("=SUM(A7:B8)+SUM(A6:B8)", {
      formulaSheetId: "Data",
      editedSheetId: "Data",
      tableColumnStart: 0,
      tableColumnEnd: 1,
      sourceRow: 5,
      targetRow: 7
    })).toEqual({
      ok: true,
      formula: "=SUM(A6:B7)+SUM(A6:B8)"
    });

    expect(rewriteFormulaForRectangularRowMove("=SUM($A$4:B5)+SUM(A4:B6)", {
      formulaSheetId: "Data",
      editedSheetId: "Data",
      tableColumnStart: 0,
      tableColumnEnd: 1,
      sourceRow: 5,
      targetRow: 3
    })).toEqual({
      ok: true,
      formula: "=SUM($A$5:B6)+SUM(A4:B6)"
    });
  });

  it("rejects discontiguous and partial-column row-move images", () => {
    const context = {
      formulaSheetId: "Data",
      editedSheetId: "Data",
      tableColumnStart: 0,
      tableColumnEnd: 1,
      sourceRow: 5,
      targetRow: 7
    };

    expect(rewriteFormulaForRectangularRowMove("=SUM(A6:B7)", context))
      .toMatchObject({ ok: false, issue: { code: "TABLE_FORMULA_REFERENCE_UNSUPPORTED" } });
    expect(rewriteFormulaForRectangularRowMove("=SUM(A6:C7)", context))
      .toMatchObject({ ok: false, issue: { code: "TABLE_FORMULA_REFERENCE_UNSUPPORTED" } });
  });

  it("translates reversed ranges wholly inside the bounded columns while honoring row locks", () => {
    expect(translateFormulaRowsWithinColumns(
      "=SUM($B3:A$3)",
      { rowOffset: 1, formulaSheetName: "Sheet1", columnStart: 0, columnEnd: 1 }
    )).toBe("=SUM($B4:A$3)");
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

  it("ignores A1-shaped text inside single-quoted sheet qualifiers", () => {
    expect(extractFormulaReferences("='A1'!B2+C2")).toEqual([
      {
        label: "C2",
        range: {
          start: { row: 1, column: 2 },
          end: { row: 1, column: 2 }
        }
      }
    ]);
    expect(extractFormulaReferences("='Sheet''s'!B2+C2")).toEqual([
      {
        label: "C2",
        range: {
          start: { row: 1, column: 2 },
          end: { row: 1, column: 2 }
        }
      }
    ]);
  });

  it.each([
    "=OtherTable[Net'#Amount]+C2",
    "=OtherTable[Net''Amount]+C2"
  ])("extracts references after structured-reference apostrophe escapes in %s", (formula) => {
    expect(extractFormulaReferences(formula)).toEqual([
      {
        label: "C2",
        range: {
          start: { row: 1, column: 2 },
          end: { row: 1, column: 2 }
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
      rewriteFormulaForStructure("=SUM($A$2:$A$6)+SUM(B3:B6)+SUM(C3:C4)+D4+Data!E4", {
        formulaSheetName: "Data",
        editedSheetName: "Data",
        axis: "row",
        mode: "delete",
        index: 2,
        count: 2
      })
    ).toBe("=SUM($A$2:$A$4)+SUM(B3:B4)+SUM(#REF!)+#REF!+#REF!");
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

  it("rewrites table-contained row inserts and rejects mixed-column ranges that require a union", () => {
    const context = {
      formulaSheetId: "Data",
      editedSheetId: "Data",
      tableColumnStart: 1,
      tableColumnEnd: 2,
      tableRowEnd: 3,
      row: 2,
      count: 1,
      operation: "insert" as const,
      sheetBounds: { rowCount: 100, columnCount: 26 }
    };
    expect(rewriteFormulaForRectangularRowEdit("=SUM(B2:C4)+$B$4+D4", context)).toEqual({
      ok: true,
      formula: "=SUM(B2:C5)+$B$5+D4"
    });
    expect(rewriteFormulaForRectangularRowEdit("=SUM(A2:C4)", context))
      .toMatchObject({ ok: false, issue: { code: "TABLE_FORMULA_REFERENCE_UNSUPPORTED" } });

    expect(rewriteFormulaForRectangularRowEdit("=SUM(A2:C2)", context)).toEqual({
      ok: true,
      formula: "=SUM(A2:C2)"
    });
  });

  it("shrinks or invalidates affected table-column slices on deletion", () => {
    const context = {
      formulaSheetId: "Data",
      editedSheetId: "Data",
      tableColumnStart: 1,
      tableColumnEnd: 2,
      tableRowEnd: 3,
      row: 1,
      count: 2,
      operation: "delete" as const,
      sheetBounds: { rowCount: 100, columnCount: 26 }
    };
    expect(rewriteFormulaForRectangularRowEdit("=SUM(B2:C4)+B2+B4", context)).toEqual({
      ok: true,
      formula: "=SUM(B2:C2)+#REF!+B2"
    });
    expect(rewriteFormulaForRectangularRowEdit("=SUM(Data!B2:B3)", context)).toEqual({
      ok: true,
      formula: "=SUM(#REF!)"
    });
    expect(rewriteFormulaForRectangularRowEdit("=SUM(A2:C4)", context))
      .toMatchObject({ ok: false, issue: { code: "TABLE_FORMULA_REFERENCE_UNSUPPORTED" } });
  });

  it.each(["insert", "delete"] as const)(
    "keeps references below the old table rectangle fixed on %s",
    (operation) => {
      const context = {
        formulaSheetId: "Data",
        editedSheetId: "Data",
        tableColumnStart: 1,
        tableColumnEnd: 3,
        tableRowEnd: 10,
        row: 4,
        count: 1,
        operation,
        sheetBounds: { rowCount: 100, columnCount: 26 }
      };

      expect(rewriteFormulaForRectangularRowEdit("=B20*2+SUM(C15:C30)", context)).toEqual({
        ok: true,
        formula: "=B20*2+SUM(C15:C30)"
      });
    }
  );

  it("rewrites only the endpoints inside the old table rectangle on insertion", () => {
    const context = {
      formulaSheetId: "Data",
      editedSheetId: "Data",
      tableColumnStart: 1,
      tableColumnEnd: 3,
      tableRowEnd: 10,
      row: 4,
      count: 2,
      operation: "insert" as const,
      sheetBounds: { rowCount: 100, columnCount: 26 }
    };

    expect(rewriteFormulaForRectangularRowEdit(
      "=SUM(B5:B20)+SUM(C3:C11)+D11+SUM(B3:B20)",
      context
    )).toEqual({
      ok: true,
      formula: "=SUM(B7:B20)+SUM(C3:C13)+D13+SUM(B3:B20)"
    });
  });

  it("rewrites only the endpoints inside the old table rectangle on deletion", () => {
    const context = {
      formulaSheetId: "Data",
      editedSheetId: "Data",
      tableColumnStart: 1,
      tableColumnEnd: 3,
      tableRowEnd: 10,
      row: 4,
      count: 2,
      operation: "delete" as const,
      sheetBounds: { rowCount: 100, columnCount: 26 }
    };

    expect(rewriteFormulaForRectangularRowEdit(
      "=SUM(B7:B20)+SUM(C3:C11)+D11+SUM(B3:B20)",
      context
    )).toEqual({
      ok: true,
      formula: "=SUM(B5:B20)+SUM(C3:C9)+D9+SUM(B3:B20)"
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
        tableRowEnd: 3,
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
        tableRowEnd: 3,
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
      tableRowEnd: 3,
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
      tableRowEnd: 3,
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

  it.each([
    ["quoted external workbook", "='[Book.xlsx]Data'!B2"],
    ["unquoted external workbook", "=[Book.xlsx]Data!B2"],
    ["unquoted 3-D", "=SUM(Jan:Mar!B2)"],
    ["quoted 3-D", "=SUM('Jan''A':'Mar''B'!B2)"]
  ])("keeps a %s formula byte-identical unless an affected local reference is also present", (_label, formula) => {
    const context = {
      formulaSheetId: "Data",
      editedSheetId: "Data",
      tableColumnStart: 1,
      tableColumnEnd: 2,
      tableRowEnd: 3,
      row: 1,
      count: 1,
      operation: "insert" as const,
      sheetBounds: { rowCount: 100, columnCount: 26 }
    };

    expect(rewriteFormulaForRectangularRowEdit(formula, context)).toEqual({
      ok: true,
      formula
    });
    expect(rewriteFormulaForRectangularRowEdit(`${formula}+B2`, context))
      .toMatchObject({ ok: false, issue: { code: "TABLE_FORMULA_REFERENCE_UNSUPPORTED" } });
  });
});
