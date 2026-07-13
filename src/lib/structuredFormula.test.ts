import { describe, expect, it } from "vitest";
import type { StructuredTable } from "../types";
import { a1FormulaToStructured, structuredFormulaToA1 } from "./structuredFormula";

const salesTable: StructuredTable = {
  id: "table-id",
  name: "SalesTable",
  sheetId: "sales-sheet",
  range: {
    start: { row: 0, column: 0 },
    end: { row: 5, column: 6 }
  },
  headerRow: true,
  totalsRow: true,
  columns: [
    { id: "order-id", name: "Order ID", sheetColumn: 0 },
    { id: "order-date", name: "Order Date", sheetColumn: 1 },
    { id: "region", name: "Region", sheetColumn: 2 },
    { id: "units", name: "Units", sheetColumn: 3 },
    { id: "unit-price", name: "Unit Price", sheetColumn: 4 },
    { id: "amount", name: "Amount", sheetColumn: 5 },
    { id: "hostile", name: "Net]#' Amount", sheetColumn: 6 }
  ],
  rowIds: ["row-1", "row-2", "row-3", "row-4"]
};

function expectFormula(result: ReturnType<typeof structuredFormulaToA1> | ReturnType<typeof a1FormulaToStructured>) {
  expect(result).toMatchObject({ ok: true });
  if (!result.ok) throw new Error(result.issue.message);
  return result.formula;
}

describe("structuredFormula", () => {
  it("imports current-row shorthand and preserves quoted text, outside references, and digit-bearing functions", () => {
    const result = structuredFormulaToA1(
      '=[@Units]*[@[Unit Price]]+A9+LOG10(A2)+"[@Units]"',
      salesTable,
      1
    );

    expect(expectFormula(result)).toBe('=D2*E2+A9+LOG10(A2)+"[@Units]"');
  });

  it("imports table-qualified current-row and explicit #This Row references", () => {
    expect(expectFormula(structuredFormulaToA1("=SalesTable[@Region]", salesTable, 2))).toBe("=C3");
    expect(
      expectFormula(
        structuredFormulaToA1("=SalesTable[[#This Row],[Amount]]", salesTable, 2)
      )
    ).toBe("=F3");
  });

  it("imports headers, totals, data, all, and adjacent column ranges", () => {
    expect(
      expectFormula(structuredFormulaToA1("=SalesTable[[#Headers],[Region]]", salesTable, 1))
    ).toBe("=C$1");
    expect(
      expectFormula(structuredFormulaToA1("=SalesTable[[#Totals],[Amount]]", salesTable, 1))
    ).toBe("=F$6");
    expect(
      expectFormula(structuredFormulaToA1("=SUM(SalesTable[[#Data],[Amount]])", salesTable, 1))
    ).toBe("=SUM(F$2:F$5)");
    expect(
      expectFormula(structuredFormulaToA1("=SUM(SalesTable[[#All],[Amount]])", salesTable, 1))
    ).toBe("=SUM(F$1:F$6)");
    expect(
      expectFormula(
        structuredFormulaToA1("=SUM(SalesTable[[#Data],[Units]:[Amount]])", salesTable, 1)
      )
    ).toBe("=SUM(D$2:F$5)");
  });

  it("imports single item specifiers without treating them as column names", () => {
    expect(expectFormula(structuredFormulaToA1("=SalesTable[#Headers]", salesTable, 1))).toBe("=A$1:G$1");
    expect(expectFormula(structuredFormulaToA1("=SalesTable[#Data]", salesTable, 1))).toBe("=A$2:G$5");
    expect(expectFormula(structuredFormulaToA1("=SalesTable[#Totals]", salesTable, 1))).toBe("=A$6:G$6");
    expect(expectFormula(structuredFormulaToA1("=SalesTable[#All]", salesTable, 1))).toBe("=A$1:G$6");

    const escapedSelectorColumn: StructuredTable = {
      ...salesTable,
      range: { ...salesTable.range, end: { ...salesTable.range.end, column: 7 } },
      columns: [
        ...salesTable.columns,
        { id: "selector-column", name: "#Data", sheetColumn: 7 }
      ]
    };
    expect(expectFormula(
      structuredFormulaToA1("=SalesTable['#Data]", escapedSelectorColumn, 1)
    )).toBe("=H$2:H$5");
  });

  it("decodes and re-encodes escaped right brackets, pound signs, and apostrophes in headers", () => {
    const structured = "=SalesTable[@[Net']'#'' Amount]]";
    const imported = structuredFormulaToA1(structured, salesTable, 1);

    expect(expectFormula(imported)).toBe("=G2");
    expect(expectFormula(a1FormulaToStructured("=G2", salesTable, 1))).toBe(
      "=SalesTable[@[Net']'#'' Amount]]"
    );
  });

  it("exports exact table shapes while preserving A1 references outside the table", () => {
    expect(
      expectFormula(a1FormulaToStructured('=D2*E2+A9+LOG10(A9)+"D2"', salesTable, 1))
    ).toBe('=SalesTable[@Units]*SalesTable[@[Unit Price]]+A9+LOG10(A9)+"D2"');
    expect(expectFormula(a1FormulaToStructured("=C1+F6", salesTable, 1))).toBe(
      "=SalesTable[[#Headers],[Region]]+SalesTable[[#Totals],[Amount]]"
    );
    expect(expectFormula(a1FormulaToStructured("=SUM(F2:F5)", salesTable, 1))).toBe(
      "=SUM(SalesTable[[#Data],[Amount]])"
    );
    expect(expectFormula(a1FormulaToStructured("=SUM(D2:F5)", salesTable, 1))).toBe(
      "=SUM(SalesTable[[#Data],[Units]:[Amount]])"
    );
    expect(expectFormula(a1FormulaToStructured("=SUM(F1:F6)", salesTable, 1))).toBe(
      "=SUM(SalesTable[[#All],[Amount]])"
    );
  });

  it("distinguishes a locked one-row data selector from an unlocked current row", () => {
    const oneRowTable: StructuredTable = {
      id: "one-row-table",
      name: "Calculated",
      sheetId: "sheet-1",
      range: {
        start: { row: 0, column: 2 },
        end: { row: 1, column: 2 }
      },
      headerRow: true,
      totalsRow: false,
      columns: [{ id: "quantity", name: "Quantity", sheetColumn: 2 }],
      rowIds: ["row-1"]
    };

    expect(expectFormula(a1FormulaToStructured("=SUM(C2)", oneRowTable, 1))).toBe(
      "=SUM(Calculated[@Quantity])"
    );
    expect(expectFormula(a1FormulaToStructured("=SUM(C$2)", oneRowTable, 1))).toBe(
      "=SUM(Calculated[[#Data],[Quantity]])"
    );
    expect(expectFormula(a1FormulaToStructured("=SUM(C2:C$2)", oneRowTable, 1))).toBe(
      "=SUM(C2:C$2)"
    );
  });

  it.each([
    ["an A1-shaped sheet name", "='A1'!B2+C2", "='A1'!B2+SalesTable[@Region]"],
    ["an escaped apostrophe in a sheet name", "='Sheet''s'!B2+C2", "='Sheet''s'!B2+SalesTable[@Region]"]
  ])("preserves %s while exporting local A1 references", (_label, formula, expected) => {
    expect(expectFormula(a1FormulaToStructured(formula, salesTable, 1))).toBe(expected);
  });

  it("skips a quoted A1 qualifier when the same local range key is recognized", () => {
    expect(expectFormula(a1FormulaToStructured("='A1'!B2+A1", salesTable, 1))).toBe(
      "='A1'!B2+SalesTable[[#Headers],[Order ID]]"
    );
  });

  it.each([
    ["=OtherTable[Net'#Amount]+C2", "=OtherTable[Net'#Amount]+SalesTable[@Region]"],
    ["=OtherTable[Net''Amount]+C2", "=OtherTable[Net''Amount]+SalesTable[@Region]"]
  ])("preserves structured-reference apostrophe escapes in %s", (formula, expected) => {
    expect(expectFormula(a1FormulaToStructured(formula, salesTable, 1))).toBe(expected);
  });

  it("round-trips an equivalent structured formula from the first body row", () => {
    const source = "=SalesTable[@Units]*SalesTable[@[Unit Price]]";
    const a1 = structuredFormulaToA1(source, salesTable, 1);
    expect(expectFormula(a1)).toBe("=D2*E2");
    expect(expectFormula(a1FormulaToStructured(expectFormula(a1), salesTable, 1))).toBe(source);
  });

  it("returns typed issues for malformed or ambiguous structured syntax", () => {
    expect(
      structuredFormulaToA1("=SalesTable[[#Headers],[#Totals],[Amount]]", salesTable, 1)
    ).toMatchObject({ ok: false, issue: { code: "TABLE_FORMULA_REFERENCE_UNSUPPORTED" } });
    expect(structuredFormulaToA1("=OtherTable[@Amount]", salesTable, 1)).toMatchObject({
      ok: false,
      issue: { code: "TABLE_FORMULA_REFERENCE_UNSUPPORTED" }
    });
    expect(structuredFormulaToA1("=SalesTable[[#Totals],[Amount]]", { ...salesTable, totalsRow: false }, 1))
      .toMatchObject({ ok: false, issue: { code: "TABLE_FORMULA_REFERENCE_UNSUPPORTED" } });
  });

  it("does not rewrite formulas without a leading equals sign", () => {
    expect(expectFormula(structuredFormulaToA1("[@Units]", salesTable, 1))).toBe("[@Units]");
    expect(expectFormula(a1FormulaToStructured("D2", salesTable, 1))).toBe("D2");
  });
});
