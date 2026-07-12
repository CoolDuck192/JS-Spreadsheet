import { describe, expect, it } from "vitest";
import type { CellRange, WorkbookModel } from "../../types";
import { createBlankWorkbook, deleteSheet, duplicateSheet, getCellContent, setCellContent } from "../../lib/workbook";
import {
  getStructuredTable,
  getStructuredTableAtCell,
  getStructuredTableBodyRange,
  getStructuredTableForSelection,
  reduceStructuredTableCommand,
  type StructuredTableCommand,
  type StructuredTableCommandServices
} from "./structuredTables";
import { validateExcelTableName } from "./tableNames";

describe("structured table metadata", () => {
  it("creates an A1:C4 table with stable lookup and body geometry", () => {
    const { workbook, services } = tableFixture(3, 4);
    const result = reduceStructuredTableCommand(workbook, {
      type: "table.create",
      sheetId: workbook.activeSheetId,
      range: range(0, 0, 3, 2),
      name: "Sales",
      headerRow: true,
      totalsRow: false
    }, services);
    expect(result.status).toBe("committed");
    const next = result.workbook;
    const table = next.tables[0];
    expect(table).toMatchObject({
      name: "Sales",
      range: range(0, 0, 3, 2),
      headerRow: true,
      columns: [{ name: "Column 1" }, { name: "Column 2" }, { name: "Column 3" }]
    });
    expect(table.rowIds).toHaveLength(3);
    expect(getStructuredTable(next, table.id)).toBe(table);
    expect(getStructuredTableAtCell(next, table.sheetId, { row: 2, column: 1 })).toBe(table);
    expect(getStructuredTableForSelection(next, table.sheetId, range(5, 5, 2, 1))).toBe(table);
    expect(getStructuredTableBodyRange(table)).toEqual(range(1, 0, 3, 2));
  });

  it.each(["", " Sales", "R", "A1", "XFD1048576", "R1C1", "x".repeat(256)])(
    "rejects invalid table name %s atomically",
    (name) => {
      const { workbook, services } = tableFixture(2, 3);
      expectRejectedUnchanged(workbook, {
        type: "table.create",
        sheetId: workbook.activeSheetId,
        range: range(0, 0, 2, 1),
        name,
        headerRow: true,
        totalsRow: false
      }, services, "TABLE_NAME_INVALID");
    }
  );

  it("uses the same validator for rename and rejects normalized name conflicts", () => {
    const { workbook, services } = tableFixture(4, 6);
    const first = commit(workbook, {
      type: "table.create", sheetId: workbook.activeSheetId, range: range(0, 0, 2, 1),
      name: "Sales", headerRow: true, totalsRow: false
    }, services);
    const second = commit(first, {
      type: "table.create", sheetId: workbook.activeSheetId, range: range(0, 2, 2, 3),
      name: "Costs", headerRow: true, totalsRow: false
    }, services);
    expectRejectedUnchanged(second, {
      type: "table.rename", tableId: second.tables[1].id, name: "ＳＡＬＥＳ"
    }, services, "TABLE_NAME_CONFLICT");
    expectRejectedUnchanged(second, {
      type: "table.rename", tableId: second.tables[1].id, name: "R1C1"
    }, services, "TABLE_NAME_INVALID");
  });

  it("rejects blank or duplicate headers", () => {
    const { workbook, services } = tableFixture(3, 3);
    const blank = setCellContent(workbook, workbook.activeSheetId, "A1", "");
    expectRejectedUnchanged(blank, {
      type: "table.create", sheetId: blank.activeSheetId, range: range(0, 0, 2, 1),
      name: "BlankHeaders", headerRow: true, totalsRow: false
    }, services, "TABLE_HEADER_INVALID");
    let duplicate = setCellContent(workbook, workbook.activeSheetId, "A1", "Same");
    duplicate = setCellContent(duplicate, duplicate.activeSheetId, "B1", "same");
    expectRejectedUnchanged(duplicate, {
      type: "table.create", sheetId: duplicate.activeSheetId, range: range(0, 0, 2, 1),
      name: "DuplicateHeaders", headerRow: true, totalsRow: false
    }, services, "TABLE_HEADER_INVALID");
  });

  it("rejects overlap, merges, protection, and impossible ranges", () => {
    const { workbook, services } = tableFixture(4, 5);
    const first = commit(workbook, {
      type: "table.create", sheetId: workbook.activeSheetId, range: range(0, 0, 2, 1),
      name: "First", headerRow: true, totalsRow: false
    }, services);
    expectRejectedUnchanged(first, {
      type: "table.create", sheetId: first.activeSheetId, range: range(1, 1, 3, 2),
      name: "Overlap", headerRow: false, totalsRow: false
    }, services, "TABLE_RANGE_OVERLAP");

    const withMerge: WorkbookModel = {
      ...workbook,
      sheets: [{ ...workbook.sheets[0], merges: [{ id: "merge-1", range: range(0, 0, 1, 1) }] }]
    };
    expectRejectedUnchanged(withMerge, {
      type: "table.create", sheetId: withMerge.activeSheetId, range: range(0, 0, 2, 1),
      name: "Merged", headerRow: true, totalsRow: false
    }, services, "TABLE_MERGE_CONFLICT");

    const protectedWorkbook: WorkbookModel = {
      ...workbook,
      sheets: [{
        ...workbook.sheets[0],
        protection: { isProtected: true, lockedCells: {}, unlockedCells: {} }
      }]
    };
    expectRejectedUnchanged(protectedWorkbook, {
      type: "table.create", sheetId: protectedWorkbook.activeSheetId, range: range(0, 0, 2, 1),
      name: "Protected", headerRow: true, totalsRow: false
    }, services, "TABLE_PROTECTED");

    for (const invalidRange of [range(-1, 0, 2, 1), range(0, 0, 5_000, 1), range(2, 1, 1, 1)]) {
      expectRejectedUnchanged(workbook, {
        type: "table.create", sheetId: workbook.activeSheetId, range: invalidRange,
        name: "Blocked", headerRow: false, totalsRow: false
      }, services, "TABLE_RANGE_BLOCKED");
    }
  });

  it("resizes while preserving surviving IDs and absorbing populated cells", () => {
    const { workbook, services } = tableFixture(4, 5);
    const created = commit(workbook, {
      type: "table.create", sheetId: workbook.activeSheetId, range: range(0, 0, 2, 1),
      name: "Resizable", headerRow: true, totalsRow: false
    }, services);
    const initial = created.tables[0];
    const expanded = commit(created, {
      type: "table.resize", tableId: initial.id, range: range(0, 0, 3, 2)
    }, services);
    const table = expanded.tables[0];
    expect(table.columns.slice(0, 2).map((column) => column.id)).toEqual(initial.columns.map((column) => column.id));
    expect(table.columns).toHaveLength(3);
    expect(table.rowIds.slice(0, 2)).toEqual(initial.rowIds);
    expect(table.rowIds).toHaveLength(3);
    expectRejectedUnchanged(expanded, {
      type: "table.resize", tableId: table.id, range: range(1, 0, 3, 2)
    }, services, "TABLE_RANGE_BLOCKED");
  });

  it("moves a totals row and its cell metadata when a table grows", () => {
    const workbook = totalsResizeFixture({ A7: "middle", B7: 50, A8: "tail", B8: 99 });
    const services = deterministicServices();

    const resized = commit(workbook, {
      type: "table.resize",
      tableId: "table-totals",
      range: range(0, 0, 7, 1)
    }, services);

    expect(getCellContent(resized, "sheet-1", "A8")).toBe("Total");
    expect(getCellContent(resized, "sheet-1", "B8")).toBe("=SUBTOTAL(109,B2:B7)");
    expect(resized.sheets[0].formats.B8).toEqual({ bold: true });
    expect(resized.sheets[0].validations.B8).toEqual({ type: "number", min: 0 });
    expect(resized.sheets[0].comments.B8).toBe("generated total");
    expect(resized.sheets[0].hyperlinks.B8).toBe("https://example.com/total");
    expect(getCellContent(resized, "sheet-1", "A6")).toBe("middle");
    expect(getCellContent(resized, "sheet-1", "B6")).toBe(50);
    expect(getCellContent(resized, "sheet-1", "A7")).toBe("tail");
    expect(getCellContent(resized, "sheet-1", "B7")).toBe(99);
  });

  it("rewrites moved and external formulas when a totals row moves down during resize", () => {
    const base = totalsResizeFixture({
      D1: "=B6*2",
      A7: "=A7*2",
      B7: 50,
      A8: "tail",
      B8: 99
    });
    const workbook: WorkbookModel = {
      ...base,
      sheets: [
        base.sheets[0],
        {
          ...base.sheets[0],
          id: "sheet-summary",
          name: "Summary",
          cells: { A1: "=Sheet1!B6*2" }
        }
      ],
      namedRanges: [{
        name: "CurrentTotal",
        sheetId: "sheet-1",
        range: range(5, 1, 5, 1)
      }],
      tables: [{
        ...base.tables[0],
        columns: base.tables[0].columns.map((column) =>
          column.id === "column-amount" ? { ...column, calculatedFormula: "=B6*2" } : column
        )
      }]
    };

    const resized = commit(workbook, {
      type: "table.resize",
      tableId: "table-totals",
      range: range(0, 0, 7, 1)
    }, deterministicServices());

    expect(getCellContent(resized, "sheet-1", "D1")).toBe("=B8*2");
    expect(getCellContent(resized, "sheet-1", "A6")).toBe("=A6*2");
    expect(getCellContent(resized, "sheet-summary", "A1")).toBe("=Sheet1!B8*2");
    expect(resized.namedRanges[0].range).toEqual(range(7, 1, 7, 1));
    expect(resized.tables[0].columns[1].calculatedFormula).toBe("=B8*2");
  });

  it("keeps outside-band targets when a formula moves up during growth", () => {
    const workbook = totalsResizeFixture({
      A7: "=A9+$A9+A$9+$A$9+A6+$A$6+Sheet1!B9+A1+SUM(A1:B2)"
    });

    const resized = commit(workbook, {
      type: "table.resize",
      tableId: "table-totals",
      range: range(0, 0, 7, 1)
    }, deterministicServices());

    expect(getCellContent(resized, "sheet-1", "A6")).toBe(
      "=A9+$A9+A$9+$A$9+A8+$A$8+Sheet1!B9+A1+SUM(A1:B2)"
    );
  });

  it("keeps a physically moved external-only body formula and rejects an affected mixed formula atomically", () => {
    const formula = "='[Book.xlsx]Data'!D3";
    const workbook = totalsResizeFixture({ A7: formula });
    const command: StructuredTableCommand = {
      type: "table.resize",
      tableId: "table-totals",
      range: range(0, 0, 7, 1)
    };

    const resized = commit(workbook, command, deterministicServices());
    expect(getCellContent(resized, "sheet-1", "A6")).toBe(formula);

    const mixed = totalsResizeFixture({ A7: `${formula}+A7` });
    expectRejectedUnchanged(
      mixed,
      command,
      deterministicServices(),
      "TABLE_FORMULA_REFERENCE_UNSUPPORTED"
    );
  });

  it.each(["Data:Other", "Other:Data"])(
    "rejects a totals move atomically when the edited sheet is in the %s 3-D span",
    (sheetSpan) => {
      const base = totalsResizeFixture({ D1: `=SUM(${sheetSpan}!B6)` });
      const workbook: WorkbookModel = {
        ...base,
        sheets: [{ ...base.sheets[0], name: "Data" }]
      };

      expectRejectedUnchanged(workbook, {
        type: "table.resize",
        tableId: "table-totals",
        range: range(0, 0, 7, 1)
      }, deterministicServices(), "TABLE_FORMULA_REFERENCE_UNSUPPORTED");
    }
  );

  it("rejects a totals move when the edited sheet is inside a 3-D span", () => {
    const base = totalsResizeFixture({ D1: "=SUM(Jan:Mar!B6)" });
    const dataSheet = { ...base.sheets[0], name: "Data" };
    const workbook: WorkbookModel = {
      ...base,
      sheets: [
        { ...dataSheet, id: "sheet-jan", name: "Jan", cells: {} },
        dataSheet,
        { ...dataSheet, id: "sheet-mar", name: "Mar", cells: {} }
      ]
    };

    expectRejectedUnchanged(workbook, {
      type: "table.resize",
      tableId: "table-totals",
      range: range(0, 0, 7, 1)
    }, deterministicServices(), "TABLE_FORMULA_REFERENCE_UNSUPPORTED");
  });

  it("rejects a discontiguous moved range during growth", () => {
    const workbook = totalsResizeFixture({ A7: "=SUM(A7:B9)" });

    expectRejectedUnchanged(workbook, {
      type: "table.resize",
      tableId: "table-totals",
      range: range(0, 0, 7, 1)
    }, deterministicServices(), "TABLE_FORMULA_REFERENCE_UNSUPPORTED");
  });

  it("rewrites moved and external formulas when a totals row moves up during resize", () => {
    const workbook = totalsResizeFixture({
      D1: "=B6*2",
      A4: "=A4*2"
    });

    const resized = commit(workbook, {
      type: "table.resize",
      tableId: "table-totals",
      range: range(0, 0, 3, 1)
    }, deterministicServices());

    expect(getCellContent(resized, "sheet-1", "D1")).toBe("=B4*2");
    expect(getCellContent(resized, "sheet-1", "A5")).toBe("=A5*2");
  });

  it("keeps outside-band targets when a formula moves down during shrink", () => {
    const workbook = totalsResizeFixture({
      A4: "=A2+$A2+A$2+$A$2+A6+$A$6+Sheet1!B2+A1+SUM(A1:B2)"
    });

    const resized = commit(workbook, {
      type: "table.resize",
      tableId: "table-totals",
      range: range(0, 0, 3, 1)
    }, deterministicServices());

    expect(getCellContent(resized, "sheet-1", "A5")).toBe(
      "=A2+$A2+A$2+$A$2+A4+$A$4+Sheet1!B2+A1+SUM(A1:B2)"
    );
  });

  it("rejects a discontiguous moved range during shrink", () => {
    const workbook = totalsResizeFixture({ A4: "=SUM(A2:B4)" });

    expectRejectedUnchanged(workbook, {
      type: "table.resize",
      tableId: "table-totals",
      range: range(0, 0, 3, 1)
    }, deterministicServices(), "TABLE_FORMULA_REFERENCE_UNSUPPORTED");
  });

  it("rejects a moved range whose locked outside endpoint leaves a gap", () => {
    const workbook = totalsResizeFixture({ A7: "=SUM(A7:B$9)" });

    expectRejectedUnchanged(workbook, {
      type: "table.resize",
      tableId: "table-totals",
      range: range(0, 0, 7, 1)
    }, deterministicServices(), "TABLE_FORMULA_REFERENCE_UNSUPPORTED");
  });

  it("leaves newly added-column cell planes fixed while lengthening a table", () => {
    const base = totalsResizeFixture({ C6: "six", C7: "seven", C8: "eight" });
    const workbook: WorkbookModel = {
      ...base,
      sheets: [{
        ...base.sheets[0],
        formats: { ...base.sheets[0].formats, C7: { italic: true } }
      }]
    };

    const resized = commit(workbook, {
      type: "table.resize",
      tableId: "table-totals",
      range: range(0, 0, 7, 2)
    }, deterministicServices());

    expect(getCellContent(resized, "sheet-1", "C6")).toBe("six");
    expect(getCellContent(resized, "sheet-1", "C7")).toBe("seven");
    expect(getCellContent(resized, "sheet-1", "C8")).toBe("eight");
    expect(resized.sheets[0].formats.C7).toEqual({ italic: true });
  });

  it("rejects an unrepresentable totals-row move atomically after deriving resize metadata", () => {
    const workbook = totalsResizeFixture({ D1: "=SUM(A6:B7)" });

    expectRejectedUnchanged(workbook, {
      type: "table.resize",
      tableId: "table-totals",
      range: range(0, 0, 7, 1)
    }, deterministicServices(), "TABLE_FORMULA_REFERENCE_UNSUPPORTED");
  });

  it("keeps displaced body data when a shrunken table later disables totals", () => {
    const workbook = totalsResizeFixture();
    const services = deterministicServices();

    const resized = commit(workbook, {
      type: "table.resize",
      tableId: "table-totals",
      range: range(0, 0, 3, 1)
    }, services);
    expect(getCellContent(resized, "sheet-1", "A4")).toBe("Total");
    expect(getCellContent(resized, "sheet-1", "B4")).toBe("=SUBTOTAL(109,B2:B3)");

    const withoutTotals = commit(resized, {
      type: "table.setTotalsRow",
      tableId: "table-totals",
      enabled: false
    }, services);

    expect(getCellContent(withoutTotals, "sheet-1", "A5")).toBe("Linus");
    expect(getCellContent(withoutTotals, "sheet-1", "B5")).toBe(30);
    expect(getCellContent(withoutTotals, "sheet-1", "A6")).toBe("Margaret");
    expect(getCellContent(withoutTotals, "sheet-1", "B6")).toBe(40);
    expect(getCellContent(withoutTotals, "sheet-1", "A4")).toBeNull();
    expect(getCellContent(withoutTotals, "sheet-1", "B4")).toBeNull();
  });

  it("rewrites moved and external formulas in both header-toggle directions", () => {
    const workbook = headerFormulaFixture();
    const services = deterministicServices();

    const withHeader = commit(workbook, {
      type: "table.setHeaderRow",
      tableId: "table-formulas",
      enabled: true
    }, services);

    expect(getCellContent(withHeader, "sheet-1", "B2")).toBe("=Z1+A2+Sheet2!B1");
    expect(getCellContent(withHeader, "sheet-1", "D1")).toBe("=SUM(A2:A4)");
    expect(getCellContent(withHeader, "sheet-summary", "A1")).toBe("=SUM(Sheet1!A2:A4)");
    expect(withHeader.tables[0].columns[1].calculatedFormula).toBe("=Z1+A2+Sheet2!B1");

    const withoutHeader = commit(withHeader, {
      type: "table.setHeaderRow",
      tableId: "table-formulas",
      enabled: false
    }, services);

    expect(getCellContent(withoutHeader, "sheet-1", "B1")).toBe("=Z1+A1+Sheet2!B1");
    expect(getCellContent(withoutHeader, "sheet-1", "D1")).toBe("=SUM(A1:A3)");
    expect(getCellContent(withoutHeader, "sheet-summary", "A1")).toBe("=SUM(Sheet1!A1:A3)");
    expect(withoutHeader.tables[0].columns[1].calculatedFormula).toBe("=Z1+A1+Sheet2!B1");
  });

  it("keeps the translated calculated-column anchor during later regeneration", () => {
    const workbook = headerFormulaFixture();
    const services = deterministicServices();
    const withHeader = commit(workbook, {
      type: "table.setHeaderRow",
      tableId: "table-formulas",
      enabled: true
    }, services);

    const inserted = commit(withHeader, {
      type: "table.insertRows",
      tableId: "table-formulas",
      count: 1
    }, services);

    expect(getCellContent(inserted, "sheet-1", "B2")).toBe("=Z1+A2+Sheet2!B1");
    expect(getCellContent(inserted, "sheet-1", "B5")).toBe("=Z4+A5+Sheet2!B4");
  });

  it("toggles headers and totals without replacing body row IDs", () => {
    let workbook = createBlankWorkbook();
    workbook = setCellContent(workbook, workbook.activeSheetId, "A1", "Ada");
    workbook = setCellContent(workbook, workbook.activeSheetId, "A2", "Grace");
    const services = deterministicServices();
    const created = commit(workbook, {
      type: "table.create", sheetId: workbook.activeSheetId, range: range(0, 0, 1, 0),
      name: "People", headerRow: false, totalsRow: false
    }, services);
    const tableId = created.tables[0].id;
    const rowIds = created.tables[0].rowIds;
    const withHeader = commit(created, { type: "table.setHeaderRow", tableId, enabled: true }, services);
    expect(withHeader.tables[0].rowIds).toEqual(rowIds);
    expect(getCellContent(withHeader, workbook.activeSheetId, "A1")).toBe("Column1");
    expect(getCellContent(withHeader, workbook.activeSheetId, "A2")).toBe("Ada");
    const withTotals = commit(withHeader, { type: "table.setTotalsRow", tableId, enabled: true }, services);
    expect(withTotals.tables[0].totalsRow).toBe(true);
    expect(withTotals.tables[0].rowIds).toEqual(rowIds);
    const restored = commit(withTotals, { type: "table.setTotalsRow", tableId, enabled: false }, services);
    expect(restored.tables[0].totalsRow).toBe(false);
    expect(restored.tables[0].rowIds).toEqual(rowIds);
  });

  it("renames columns, applies metadata, and converts to a formatted range", () => {
    const { workbook, services } = tableFixture(3, 3);
    const created = commit(workbook, {
      type: "table.create", sheetId: workbook.activeSheetId, range: range(0, 0, 2, 1),
      name: "Styled", headerRow: true, totalsRow: false
    }, services);
    const table = created.tables[0];
    const renamed = commit(created, {
      type: "table.renameColumn", tableId: table.id, columnId: table.columns[0].id, name: "Account"
    }, services);
    expect(getCellContent(renamed, workbook.activeSheetId, "A1")).toBe("Account");
    const keyed = commit(renamed, {
      type: "table.setKeyColumn", tableId: table.id, columnId: table.columns[0].id
    }, services);
    const styled = commit(keyed, {
      type: "table.setStyle", tableId: table.id,
      style: { theme: "TableStyleLight2", showRowStripes: true }
    }, services);
    const converted = commit(styled, { type: "table.convertToRange", tableId: table.id }, services);
    expect(converted.tables).toEqual([]);
    expect(converted.sheets[0].formats.A1).toMatchObject({ bold: true });
    expect(getCellContent(converted, workbook.activeSheetId, "A2")).not.toBeNull();
  });

  it("duplicates and deletes sheet-owned tables with fresh IDs", () => {
    const { workbook, services } = tableFixture(3, 3);
    const created = commit(workbook, {
      type: "table.create", sheetId: workbook.activeSheetId, range: range(0, 0, 2, 1),
      name: "Source", headerRow: true, totalsRow: false
    }, services);
    const duplicated = duplicateSheet(created, created.activeSheetId, services.createId);
    expect(duplicated.tables).toHaveLength(2);
    expect(duplicated.tables[1].sheetId).toBe(duplicated.activeSheetId);
    expect(duplicated.tables[1].id).not.toBe(duplicated.tables[0].id);
    expect(duplicated.tables[1].columns[0].id).not.toBe(duplicated.tables[0].columns[0].id);
    expect(duplicated.tables[1].name).toBe("Source_Copy");
    const withNormalizedCollision = {
      ...duplicated,
      tables: duplicated.tables.map((table, index) => index === 1
        ? { ...table, name: "ＳＯＵＲＣＥ＿ＣＯＰＹ" }
        : table)
    };
    const duplicatedAgain = duplicateSheet(withNormalizedCollision, created.activeSheetId, services.createId);
    expect(duplicatedAgain.tables.at(-1)?.name).toBe("Source_Copy_2");
    const deleted = deleteSheet(duplicated, created.activeSheetId);
    expect(deleted.tables).toHaveLength(1);
    expect(deleted.tables[0].sheetId).toBe(duplicated.activeSheetId);
  });

  it("keeps duplicated names within Excel's normalized character limit", () => {
    const { workbook, services } = tableFixture(2, 3);
    const sourceName = "ﬃ".repeat(85);
    const created = commit(workbook, {
      type: "table.create", sheetId: workbook.activeSheetId, range: range(0, 0, 2, 1),
      name: sourceName, headerRow: true, totalsRow: false
    }, services);

    const duplicated = duplicateSheet(created, created.activeSheetId, services.createId);
    const duplicateName = duplicated.tables[1].name;

    expect(duplicateName).toBe(duplicateName.normalize("NFKC"));
    expect([...duplicateName]).toHaveLength(255);
    expect(validateExcelTableName(duplicateName).valid).toBe(true);
  });
});

function tableFixture(columnCount: number, rowCount: number) {
  let workbook = createBlankWorkbook();
  for (let column = 0; column < columnCount; column += 1) {
    const letter = String.fromCharCode(65 + column);
    workbook = setCellContent(workbook, workbook.activeSheetId, `${letter}1`, `Column ${column + 1}`);
    for (let row = 2; row <= rowCount; row += 1) {
      workbook = setCellContent(workbook, workbook.activeSheetId, `${letter}${row}`, `${letter}-${row}`);
    }
  }
  return { workbook, services: deterministicServices() };
}

function totalsResizeFixture(overrides: Record<string, string | number> = {}): WorkbookModel {
  const workbook = createBlankWorkbook();
  return {
    ...workbook,
    sheets: [{
      ...workbook.sheets[0],
      cells: {
        A1: "Name", B1: "Amount",
        A2: "Ada", B2: 10,
        A3: "Grace", B3: 20,
        A4: "Linus", B4: 30,
        A5: "Margaret", B5: 40,
        A6: "Total", B6: "=SUBTOTAL(109,B2:B5)",
        ...overrides
      },
      formats: { B6: { bold: true } },
      validations: { B6: { type: "number", min: 0 } },
      comments: { B6: "generated total" },
      hyperlinks: { B6: "https://example.com/total" }
    }],
    tables: [{
      id: "table-totals",
      name: "TotalsTable",
      sheetId: "sheet-1",
      range: range(0, 0, 5, 1),
      headerRow: true,
      totalsRow: true,
      columns: [
        { id: "column-name", name: "Name", sheetColumn: 0, totalsLabel: "Total" },
        { id: "column-amount", name: "Amount", sheetColumn: 1, totalsFunction: "sum" }
      ],
      rowIds: ["row-ada", "row-grace", "row-linus", "row-margaret"]
    }]
  };
}

function headerFormulaFixture(): WorkbookModel {
  const workbook = createBlankWorkbook();
  const sheet = workbook.sheets[0];
  return {
    ...workbook,
    sheets: [
      {
        ...sheet,
        cells: {
          A1: 1, B1: "=Z1+A1+Sheet2!B1",
          A2: 2, B2: "=Z2+A2+Sheet2!B2",
          A3: 3, B3: "=Z3+A3+Sheet2!B3",
          D1: "=SUM(A1:A3)"
        }
      },
      {
        ...sheet,
        id: "sheet-summary",
        name: "Summary",
        cells: { A1: "=SUM(Sheet1!A1:A3)" }
      }
    ],
    tables: [{
      id: "table-formulas",
      name: "FormulaTable",
      sheetId: "sheet-1",
      range: range(0, 0, 2, 1),
      headerRow: false,
      totalsRow: false,
      columns: [
        { id: "column-value", name: "Value", sheetColumn: 0 },
        {
          id: "column-formula",
          name: "Formula",
          sheetColumn: 1,
          calculatedFormula: "=Z1+A1+Sheet2!B1"
        }
      ],
      rowIds: ["row-1", "row-2", "row-3"]
    }]
  };
}

function deterministicServices(): StructuredTableCommandServices {
  let next = 0;
  return {
    createId(kind) {
      next += 1;
      return `${kind}-${next}`;
    },
    getCellEvaluation() {
      return null;
    }
  };
}

function commit(
  workbook: WorkbookModel,
  command: StructuredTableCommand,
  services: StructuredTableCommandServices
): WorkbookModel {
  const result = reduceStructuredTableCommand(workbook, command, services);
  expect(result.status).toBe("committed");
  return result.workbook;
}

function expectRejectedUnchanged(
  workbook: WorkbookModel,
  command: StructuredTableCommand,
  services: StructuredTableCommandServices,
  code: string
): void {
  const before = JSON.stringify(workbook);
  const result = reduceStructuredTableCommand(workbook, command, services);
  expect(result).toMatchObject({ status: "rejected", workbook, issues: [{ code }] });
  expect(result.workbook).toBe(workbook);
  expect(JSON.stringify(workbook)).toBe(before);
}

function range(startRow: number, startColumn: number, endRow: number, endColumn: number): CellRange {
  return {
    start: { row: startRow, column: startColumn },
    end: { row: endRow, column: endColumn }
  };
}
