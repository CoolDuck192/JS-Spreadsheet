import { describe, expect, it } from "vitest";
import type { CellContent, DataValidationRule, StructuredTable, WorkbookModel } from "../../types";
import { createFormulaEngine } from "../../lib/formulaEngine";
import { addSheet, createBlankWorkbook, getCellContent, setCellContent } from "../../lib/workbook";
import {
  deleteStructuredTableRows,
  insertStructuredTableRows,
  setStructuredTableCalculatedColumn,
  sortStructuredTableRows
} from "./structuredTableRows";
import type { StructuredTableCommandServices } from "./structuredTables";

describe("structured table rows", () => {
  it("inserts before a stable row ID and preserves existing IDs and cell planes", () => {
    const workbook = rowFixture();
    const result = insertStructuredTableRows(workbook, "table-1", {
      beforeRowId: "row-2",
      count: 2
    }, servicesFor(workbook));
    expect(result.status).toBe("committed");
    const next = result.workbook;
    expect(table(next).rowIds).toEqual([
      "row-1", "table-row-1", "table-row-2", "row-2", "row-3"
    ]);
    expect(getCellContent(next, "sheet-1", "A5")).toBe("Grace");
    expect(next.sheets[0].formats.A5).toEqual({ bold: true });
    expect(next.sheets[0].validations.A5).toEqual({ type: "textLength", min: 2 });
    expect(next.sheets[0].comments.A5).toBe("moved comment");
    expect(next.sheets[0].hyperlinks.A5).toBe("https://example.com/grace");
    expect(getCellContent(next, "sheet-1", "D3")).toBe("unrelated-right");
  });

  it.each([
    ["appending", {}],
    ["prepending", { beforeRowId: "row-1" }]
  ])("regenerates generated totals after %s table rows", (_label, anchor) => {
    const workbook = totalsFixture();

    const result = insertStructuredTableRows(
      workbook,
      "table-1",
      { count: 1, ...anchor },
      servicesFor(workbook)
    );

    expect(result.status).toBe("committed");
    expect(getCellContent(result.workbook, "sheet-1", "B6")).toBe("=SUBTOTAL(109,B2:B5)");
  });

  it("deletes non-contiguous stable row IDs atomically", () => {
    const workbook = rowFixture();
    const result = deleteStructuredTableRows(workbook, "table-1", ["row-1", "row-3"], servicesFor(workbook));
    expect(result.status).toBe("committed");
    expect(table(result.workbook).rowIds).toEqual(["row-2"]);
    expect(getCellContent(result.workbook, "sheet-1", "A2")).toBe("Grace");
    expect(result.workbook.sheets[0].formats.A2).toEqual({ bold: true });
    expect(getCellContent(result.workbook, "sheet-1", "A3")).toBeNull();
    expect(getCellContent(result.workbook, "sheet-1", "D3")).toBe("unrelated-right");
  });

  it("rewrites formulas across non-contiguous deletes like sequential bottom-up deletes", () => {
    const workbook = rowFixture({ D1: "=SUM(A4:A6)" });

    const { composite, sequential } = deleteCompositeAndSequential(workbook);

    expect(getCellContent(composite, "sheet-1", "D1")).toBe("=SUM(A4:A6)");
    expect(getCellContent(composite, "sheet-1", "D1"))
      .toBe(getCellContent(sequential, "sheet-1", "D1"));
  });

  it("rewrites named ranges across non-contiguous deletes like sequential bottom-up deletes", () => {
    const workbook: WorkbookModel = {
      ...rowFixture(),
      namedRanges: [{
        name: "CrossesDeletedTableBottom",
        sheetId: "sheet-1",
        range: { start: { row: 3, column: 0 }, end: { row: 5, column: 0 } }
      }]
    };

    const { composite, sequential } = deleteCompositeAndSequential(workbook);

    expect(composite.namedRanges[0].range).toEqual({
      start: { row: 3, column: 0 },
      end: { row: 5, column: 0 }
    });
    expect(composite.namedRanges).toEqual(sequential.namedRanges);
  });

  it("preserves named-range array identity when a table row edit leaves every range unchanged", () => {
    const base = rowFixture();
    const workbook: WorkbookModel = {
      ...base,
      namedRanges: [{
        name: "Unrelated",
        sheetId: "sheet-1",
        range: { start: { row: 10, column: 3 }, end: { row: 11, column: 3 } }
      }]
    };

    const inserted = insertStructuredTableRows(workbook, "table-1", {
      beforeRowId: "row-2",
      count: 1
    }, servicesFor(workbook));

    expect(inserted.status).toBe("committed");
    expect(inserted.workbook.namedRanges).toBe(workbook.namedRanges);
  });

  it("keeps out-of-table references pinned when body formulas move", () => {
    const workbook = rowFixture({ B3: "=A3+D3+1" });

    const inserted = insertStructuredTableRows(workbook, "table-1", {
      beforeRowId: "row-2",
      count: 1
    }, servicesFor(workbook));
    expect(inserted.status).toBe("committed");
    expect(getCellContent(inserted.workbook, "sheet-1", "B4")).toBe("=A4+D3+1");

    const deleted = deleteStructuredTableRows(workbook, "table-1", ["row-1"], servicesFor(workbook));
    expect(deleted.status).toBe("committed");
    expect(getCellContent(deleted.workbook, "sheet-1", "B2")).toBe("=A2+D3+1");

    const sortable = rowFixture({ A3: "Zed", B3: "=A3+D3+1" });
    const sorted = sortStructuredTableRows(sortable, "table-1", [
      { columnId: "column-name", direction: "asc" }
    ], servicesFor(sortable));
    expect(sorted.status).toBe("committed");
    expect(getCellContent(sorted.workbook, "sheet-1", "B4")).toBe("=A4+D3+1");
  });

  it("preserves row locks while translating in-table relative references during sorting", () => {
    const workbook = rowFixture({
      A3: "Zed",
      B3: "=A3+$A3+A$3+$A$3+D3+Other!A3"
    });

    const sorted = sortStructuredTableRows(workbook, "table-1", [
      { columnId: "column-name", direction: "asc" }
    ], servicesFor(workbook));

    expect(sorted.status).toBe("committed");
    expect(getCellContent(sorted.workbook, "sheet-1", "B4")).toBe(
      "=A4+$A4+A$3+$A$3+D3+Other!A3"
    );
  });

  it("preserves mixed table ranges as evaluable formulas during sorting", () => {
    const workbook = rowFixture({
      A3: "=SUM(B3:D3)",
      B3: 20,
      C3: 4,
      D3: 5
    });

    const sorted = sortStructuredTableRows(workbook, "table-1", [
      { columnId: "column-score", direction: "asc" }
    ], servicesFor(workbook));

    expect(sorted.status).toBe("committed");
    expect(getCellContent(sorted.workbook, "sheet-1", "A4")).toBe("=SUM(B3:D3)");
    const engine = createFormulaEngine(sorted.workbook);
    try {
      expect(engine.getComputedValue("sheet-1", "A4")).toBe(12);
    } finally {
      engine.destroy();
    }
  });

  it.each([
    ["mixed-column range", "=SUM(A3:C3)"],
    ["external workbook reference", "='[Book.xlsx]Data'!D3"]
  ])("rejects row insertion atomically for an unsupported body formula with a %s", (_label, formula) => {
    const workbook = rowFixture({ B3: formula });

    const result = insertStructuredTableRows(workbook, "table-1", {
      beforeRowId: "row-2",
      count: 1
    }, servicesFor(workbook));

    expect(result.status).toBe("rejected");
    expect(result.workbook).toBe(workbook);
    expect(result).toMatchObject({
      issues: [{ code: "TABLE_FORMULA_REFERENCE_UNSUPPORTED" }]
    });
    expect(getCellContent(result.workbook, "sheet-1", "B3")).toBe(formula);
  });

  it("rejects unknown or duplicate anchors and row IDs without partial movement", () => {
    const workbook = rowFixture();
    for (const result of [
      insertStructuredTableRows(workbook, "table-1", { beforeRowId: "missing", count: 1 }, servicesFor(workbook)),
      insertStructuredTableRows(workbook, "table-1", { beforeRowId: "row-1", afterRowId: "row-2", count: 1 }, servicesFor(workbook)),
      deleteStructuredTableRows(workbook, "table-1", ["row-1", "row-1"], servicesFor(workbook)),
      deleteStructuredTableRows(workbook, "table-1", ["missing"], servicesFor(workbook))
    ]) {
      expect(result.status).toBe("rejected");
      expect(result.workbook).toBe(workbook);
    }
  });

  it("fills calculated columns with token-aware relative references", () => {
    const workbook = calculatedFixture();
    const result = setStructuredTableCalculatedColumn(
      workbook,
      "table-calc",
      "column-total",
      "=C2*D2+$A$1",
      servicesFor(workbook)
    );
    expect(result.status).toBe("committed");
    expect(getCellContent(result.workbook, "sheet-1", "E2")).toBe("=C2*D2+$A$1");
    expect(getCellContent(result.workbook, "sheet-1", "E3")).toBe("=C3*D3+$A$1");
    expect(getCellContent(result.workbook, "sheet-1", "E4")).toBe("=C4*D4+$A$1");
    expect(table(result.workbook).columns[4].calculatedFormula).toBe("=C2*D2+$A$1");
  });

  it("translates structured current-row references before evaluating a calculated column", () => {
    const workbook = calculatedFixture();
    const result = setStructuredTableCalculatedColumn(
      workbook,
      "table-calc",
      "column-total",
      "=[@Quantity]*[@Price]",
      servicesFor(workbook)
    );

    expect(result.status).toBe("committed");
    expect(table(result.workbook).columns[4].calculatedFormula).toBe("=C2*D2");
    expect(getCellContent(result.workbook, "sheet-1", "E2")).toBe("=C2*D2");
    expect(getCellContent(result.workbook, "sheet-1", "E3")).toBe("=C3*D3");
    expect(getCellContent(result.workbook, "sheet-1", "E4")).toBe("=C4*D4");
    const engine = createFormulaEngine(result.workbook);
    try {
      expect(engine.getComputedValue("sheet-1", "E2")).toBe(8);
      expect(engine.getComputedValue("sheet-1", "E3")).toBe(15);
      expect(engine.getComputedValue("sheet-1", "E4")).toBe(24);
    } finally {
      engine.destroy();
    }
  });

  it("rejects an unknown structured calculated-column reference atomically", () => {
    const workbook = calculatedFixture();

    const result = setStructuredTableCalculatedColumn(
      workbook,
      "table-calc",
      "column-total",
      "=[@Missing]*2",
      servicesFor(workbook)
    );

    expect(result.status).toBe("rejected");
    expect(result.workbook).toBe(workbook);
    expect(result).toMatchObject({
      issues: [{ code: "TABLE_FORMULA_REFERENCE_UNSUPPORTED" }]
    });
  });

  it("normalizes a structured calculated formula before an empty table receives rows", () => {
    const base = calculatedFixture();
    const workbook: WorkbookModel = {
      ...base,
      sheets: [{
        ...base.sheets[0],
        cells: { A1: "A", B1: "B", C1: "Quantity", D1: "Price", E1: "Total" }
      }],
      tables: [{
        ...base.tables[0],
        range: { ...base.tables[0].range, end: { row: 0, column: 4 } },
        rowIds: []
      }]
    };
    const calculated = setStructuredTableCalculatedColumn(
      workbook,
      "table-calc",
      "column-total",
      "=[@Quantity]*[@Price]",
      servicesFor(workbook)
    );
    const inserted = insertStructuredTableRows(
      calculated.workbook,
      "table-calc",
      { count: 1 },
      servicesFor(calculated.workbook)
    );

    expect(calculated.status).toBe("committed");
    expect(table(calculated.workbook).columns[4].calculatedFormula).toBe("=C2*D2");
    expect(inserted.status).toBe("committed");
    expect(getCellContent(inserted.workbook, "sheet-1", "E2")).toBe("=C2*D2");
  });

  it("keeps a calculated-column anchor at the first body row when prepending", () => {
    const workbook = calculatedFixture();
    const calculated = setStructuredTableCalculatedColumn(
      workbook,
      "table-calc",
      "column-total",
      "=C2*D2+$A$1",
      servicesFor(workbook)
    );
    const prepended = insertStructuredTableRows(calculated.workbook, "table-calc", {
      beforeRowId: "row-1",
      count: 1
    }, servicesFor(calculated.workbook));

    expect(prepended.status).toBe("committed");
    expect(table(prepended.workbook).columns[4].calculatedFormula).toBe("=C2*D2+$A$1");
    expect(getCellContent(prepended.workbook, "sheet-1", "E2")).toBe("=C2*D2+$A$1");
  });

  it("preserves single-quoted sheet names while regenerating calculated columns", () => {
    const workbook = calculatedFixture();
    const result = setStructuredTableCalculatedColumn(
      workbook,
      "table-calc",
      "column-total",
      "='Q1 data'!B2+C2",
      servicesFor(workbook)
    );

    expect(result.status).toBe("committed");
    expect(getCellContent(result.workbook, "sheet-1", "E2")).toBe("='Q1 data'!B2+C2");
    expect(getCellContent(result.workbook, "sheet-1", "E3")).toBe("='Q1 data'!B3+C3");
    expect(getCellContent(result.workbook, "sheet-1", "E4")).toBe("='Q1 data'!B4+C4");
  });

  it("rewrites same- and cross-sheet dependents and named ranges only inside table columns", () => {
    let workbook = rowFixture();
    workbook = setCellContent(workbook, "sheet-1", "D1", "=SUM(A2:A4)+D3");
    workbook = {
      ...workbook,
      namedRanges: [{
        name: "PeopleNames", sheetId: "sheet-1",
        range: { start: { row: 1, column: 0 }, end: { row: 3, column: 0 } }
      }]
    };
    workbook = addSheet(workbook, "Summary");
    const summaryId = workbook.activeSheetId;
    workbook = setCellContent(workbook, summaryId, "A1", "=SUM(Sheet1!A2:A4)");
    const result = insertStructuredTableRows(workbook, "table-1", {
      beforeRowId: "row-2",
      count: 1
    }, servicesFor(workbook));
    expect(result.status).toBe("committed");
    expect(getCellContent(result.workbook, "sheet-1", "D1")).toBe("=SUM(A2:A5)+D3");
    expect(getCellContent(result.workbook, summaryId, "A1")).toBe("=SUM(Sheet1!A2:A5)");
    expect(result.workbook.namedRanges[0].range).toEqual({
      start: { row: 1, column: 0 },
      end: { row: 4, column: 0 }
    });
  });

  it("bounds insertion rewrites to rows moved inside the old table", () => {
    let workbook = rowFixture({ B10: 7 });
    workbook = setCellContent(
      workbook,
      "sheet-1",
      "D1",
      "=B10*2+SUM(A3:A10)+SUM(A2:A4)"
    );
    workbook = {
      ...workbook,
      namedRanges: [
        {
          name: "BelowTable",
          sheetId: "sheet-1",
          range: { start: { row: 9, column: 1 }, end: { row: 11, column: 1 } }
        },
        {
          name: "StartsInside",
          sheetId: "sheet-1",
          range: { start: { row: 2, column: 0 }, end: { row: 9, column: 0 } }
        },
        {
          name: "EndsInside",
          sheetId: "sheet-1",
          range: { start: { row: 1, column: 0 }, end: { row: 3, column: 0 } }
        }
      ]
    };

    const result = insertStructuredTableRows(workbook, "table-1", {
      beforeRowId: "row-2",
      count: 1
    }, servicesFor(workbook));

    expect(result.status).toBe("committed");
    expect(getCellContent(result.workbook, "sheet-1", "D1")).toBe(
      "=B10*2+SUM(A4:A10)+SUM(A2:A5)"
    );
    expect(result.workbook.namedRanges.map(({ name, range }) => ({ name, range }))).toEqual([
      {
        name: "BelowTable",
        range: { start: { row: 9, column: 1 }, end: { row: 11, column: 1 } }
      },
      {
        name: "StartsInside",
        range: { start: { row: 3, column: 0 }, end: { row: 9, column: 0 } }
      },
      {
        name: "EndsInside",
        range: { start: { row: 1, column: 0 }, end: { row: 4, column: 0 } }
      }
    ]);
  });

  it("bounds deletion rewrites to rows moved inside the old table", () => {
    let workbook = rowFixture({ B10: 7 });
    workbook = setCellContent(
      workbook,
      "sheet-1",
      "D1",
      "=B10*2+SUM(A4:A10)+SUM(A2:A4)"
    );
    workbook = {
      ...workbook,
      namedRanges: [
        {
          name: "BelowTable",
          sheetId: "sheet-1",
          range: { start: { row: 9, column: 1 }, end: { row: 11, column: 1 } }
        },
        {
          name: "StartsInside",
          sheetId: "sheet-1",
          range: { start: { row: 3, column: 0 }, end: { row: 9, column: 0 } }
        },
        {
          name: "EndsInside",
          sheetId: "sheet-1",
          range: { start: { row: 1, column: 0 }, end: { row: 3, column: 0 } }
        }
      ]
    };

    const result = deleteStructuredTableRows(
      workbook,
      "table-1",
      ["row-2"],
      servicesFor(workbook)
    );

    expect(result.status).toBe("committed");
    expect(getCellContent(result.workbook, "sheet-1", "D1")).toBe(
      "=B10*2+SUM(A3:A10)+SUM(A2:A3)"
    );
    expect(result.workbook.namedRanges.map(({ name, range }) => ({ name, range }))).toEqual([
      {
        name: "BelowTable",
        range: { start: { row: 9, column: 1 }, end: { row: 11, column: 1 } }
      },
      {
        name: "StartsInside",
        range: { start: { row: 2, column: 0 }, end: { row: 9, column: 0 } }
      },
      {
        name: "EndsInside",
        range: { start: { row: 1, column: 0 }, end: { row: 2, column: 0 } }
      }
    ]);
  });

  it("sorts formula results numerically and permutes row IDs with every cell plane", () => {
    const workbook = rowFixture({ B2: "=10", B3: "=2", B4: "=2" });
    const engine = createFormulaEngine(workbook);
    try {
      const result = sortStructuredTableRows(workbook, "table-1", [
        { columnId: "column-score", direction: "asc", nulls: "last" },
        { columnId: "column-name", direction: "asc" }
      ], {
        ...servicesFor(workbook),
        getCellEvaluation: engine.getComputedValue
      });
      expect(result.status).toBe("committed");
      expect(table(result.workbook).rowIds).toEqual(["row-2", "row-3", "row-1"]);
      expect(getCellContent(result.workbook, "sheet-1", "A2")).toBe("Grace");
      expect(result.workbook.sheets[0].formats.A2).toEqual({ bold: true });
    } finally {
      engine.destroy();
    }
  });
});

function rowFixture(overrides: Record<string, CellContent> = {}): WorkbookModel {
  const workbook = createBlankWorkbook();
  const validation: DataValidationRule = { type: "textLength", min: 2 };
  const table: StructuredTable = {
    id: "table-1",
    name: "People",
    sheetId: "sheet-1",
    range: { start: { row: 0, column: 0 }, end: { row: 3, column: 1 } },
    headerRow: true,
    totalsRow: false,
    columns: [
      { id: "column-name", name: "Name", sheetColumn: 0 },
      { id: "column-score", name: "Score", sheetColumn: 1, dataType: "number" }
    ],
    rowIds: ["row-1", "row-2", "row-3"]
  };
  return {
    ...workbook,
    tables: [table],
    sheets: [{
      ...workbook.sheets[0],
      cells: {
        A1: "Name", B1: "Score",
        A2: "Ada", B2: 3,
        A3: "Grace", B3: 2,
        A4: "Linus", B4: 1,
        D3: "unrelated-right",
        ...overrides
      },
      formats: { A3: { bold: true } },
      validations: { A3: validation },
      comments: { A3: "moved comment" },
      hyperlinks: { A3: "https://example.com/grace" }
    }]
  };
}

function calculatedFixture(): WorkbookModel {
  const workbook = createBlankWorkbook();
  return {
    ...workbook,
    tables: [{
      id: "table-calc", name: "Calculated", sheetId: "sheet-1",
      range: { start: { row: 0, column: 0 }, end: { row: 3, column: 4 } },
      headerRow: true, totalsRow: false,
      columns: [
        { id: "column-a", name: "A", sheetColumn: 0 },
        { id: "column-b", name: "B", sheetColumn: 1 },
        { id: "column-c", name: "Quantity", sheetColumn: 2 },
        { id: "column-d", name: "Price", sheetColumn: 3 },
        { id: "column-total", name: "Total", sheetColumn: 4 }
      ],
      rowIds: ["row-1", "row-2", "row-3"]
    }],
    sheets: [{
      ...workbook.sheets[0],
      cells: {
        A1: "A", B1: "B", C1: "Quantity", D1: "Price", E1: "Total",
        C2: 2, D2: 4, C3: 3, D3: 5, C4: 4, D4: 6
      }
    }]
  };
}

function totalsFixture(): WorkbookModel {
  const workbook = rowFixture({ B5: "=SUBTOTAL(109,B2:B4)" });
  const current = workbook.tables[0];
  return {
    ...workbook,
    tables: [{
      ...current,
      range: { ...current.range, end: { ...current.range.end, row: 4 } },
      totalsRow: true,
      columns: current.columns.map((column) => column.id === "column-score"
        ? { ...column, totalsFunction: "sum" }
        : column)
    }]
  };
}

function servicesFor(workbook: WorkbookModel): StructuredTableCommandServices {
  let next = 0;
  return {
    createId(kind) {
      next += 1;
      return `${kind}-${next}`;
    },
    getCellEvaluation(sheetId, address) {
      return workbook.sheets.find((sheet) => sheet.id === sheetId)?.cells[address] ?? null;
    }
  };
}

function deleteCompositeAndSequential(workbook: WorkbookModel): {
  composite: WorkbookModel;
  sequential: WorkbookModel;
} {
  const composite = deleteStructuredTableRows(
    workbook,
    "table-1",
    ["row-1", "row-3"],
    servicesFor(workbook)
  );
  const lower = deleteStructuredTableRows(
    workbook,
    "table-1",
    ["row-3"],
    servicesFor(workbook)
  );
  const sequential = deleteStructuredTableRows(
    lower.workbook,
    "table-1",
    ["row-1"],
    servicesFor(lower.workbook)
  );
  expect(composite.status).toBe("committed");
  expect(lower.status).toBe("committed");
  expect(sequential.status).toBe("committed");
  return { composite: composite.workbook, sequential: sequential.workbook };
}

function table(workbook: WorkbookModel): StructuredTable {
  return workbook.tables[0];
}
