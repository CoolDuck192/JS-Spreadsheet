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

function table(workbook: WorkbookModel): StructuredTable {
  return workbook.tables[0];
}
