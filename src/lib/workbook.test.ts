import { describe, expect, it } from "vitest";
import type { CellRange, StructuredTable } from "../types";
import { createFormulaEngine } from "./formulaEngine";
import {
  addSheetChart,
  addSheet,
  addSheetFilter,
  autoFillRange,
  clearRangeAll,
  clearCellComments,
  clearCellFormats,
  clearCellHyperlinks,
  clearRange,
  clearSheetFilter,
  clearSheetFilters,
  commitHistory,
  copyRange,
  copyRichRange,
  createBlankWorkbook,
  createHistory,
  defineNamedRange,
  deleteSheetChart,
  deleteSheet,
  duplicateSheet,
  fillDown,
  fillRight,
  getCellMerge,
  getSheetAutoFilterRange,
  getSheetFilters,
  getSheetCharts,
  getSheetTabColor,
  getCellComment,
  getNamedRange,
  getNamedRangeForSelection,
  getColumnWidth,
  getRowHeight,
  getCellConditionalFormatRules,
  getCellFormat,
  getCellHyperlink,
  getCellReadOnly,
  getActiveSheet,
  getCellContent,
  getCellValidation,
  isColumnHidden,
  isRowHidden,
  isSheetHidden,
  insertColumns,
  insertRows,
  mergeCells,
  deleteColumns,
  deleteRows,
  moveRichRange,
  moveSheet,
  pasteMatrix,
  pasteRichRange,
  previewRichPaste,
  redoHistory,
  removeDuplicateRows,
  removeNamedRange,
  removeConditionalFormatRule,
  renameSheet,
  setSheetHidden,
  setCellFormat,
  setCellBorders,
  addConditionalFormatRule,
  clearConditionalFormatRules,
  setActiveSheet,
  setCellContent,
  setCellComment,
  setColumnWidth,
  setCellValidation,
  setCellHyperlink,
  setRangeReadOnly,
  setSheetFreezePanes,
  setSheetProtection,
  setSheetTabColor,
  setColumnsHidden,
  setRowsHidden,
  setRowHeight,
  sortRange,
  clearHiddenRowsAndColumns,
  unhideAllSheets,
  unmergeCells,
  undoHistory
} from "./workbook";

const range = (start: string, end = start): CellRange => ({
  start: { row: Number(start.slice(1)) - 1, column: start.charCodeAt(0) - 65 },
  end: { row: Number(end.slice(1)) - 1, column: end.charCodeAt(0) - 65 }
});

describe("workbook", () => {
  it("creates a blank workbook with one sheet", () => {
    const workbook = createBlankWorkbook();

    expect(workbook.version).toBe(2);
    expect(workbook.tables).toEqual([]);
    expect(workbook.sheets).toHaveLength(1);
    expect(getActiveSheet(workbook).name).toBe("Sheet1");
    expect(getActiveSheet(workbook).rowCount).toBe(100);
    expect(getActiveSheet(workbook).columnCount).toBe(26);
  });

  it("sets, clears, and duplicates sheet tab colors", () => {
    const workbook = createBlankWorkbook();
    const sheetId = workbook.activeSheetId;

    const colored = setSheetTabColor(workbook, sheetId, "#0f766e");
    expect(getSheetTabColor(colored, sheetId)).toBe("#0f766e");
    expect(setSheetTabColor(colored, sheetId, "#0f766e")).toBe(colored);

    const duplicated = duplicateSheet(colored, sheetId);
    expect(getActiveSheet(duplicated).name).toBe("Sheet1 Copy");
    expect(getSheetTabColor(duplicated, duplicated.activeSheetId)).toBe("#0f766e");

    const cleared = setSheetTabColor(colored, sheetId, "");
    expect(getSheetTabColor(cleared, sheetId)).toBeUndefined();
  });

  it("sets, reads, and clears cell content", () => {
    const workbook = createBlankWorkbook();
    const edited = setCellContent(workbook, workbook.activeSheetId, "A1", "42");

    expect(getCellContent(edited, edited.activeSheetId, "A1")).toBe("42");
    expect(getCellContent(workbook, workbook.activeSheetId, "A1")).toBeNull();

    const cleared = clearRange(edited, edited.activeSheetId, range("A1"));
    expect(getCellContent(cleared, cleared.activeSheetId, "A1")).toBeNull();
  });

  it("does not create new workbook objects for unchanged cell content", () => {
    const workbook = createBlankWorkbook();
    const sheetId = workbook.activeSheetId;

    expect(setCellContent(workbook, sheetId, "A1", "")).toBe(workbook);

    const edited = setCellContent(workbook, sheetId, "A1", "42");
    expect(setCellContent(edited, sheetId, "A1", "42")).toBe(edited);
  });

  it("does not create new workbook objects when clearing empty ranges", () => {
    const workbook = createBlankWorkbook();

    expect(clearRange(workbook, workbook.activeSheetId, range("A1", "B2"))).toBe(workbook);
  });

  it("copies and pastes tabular data", () => {
    let workbook = createBlankWorkbook();
    workbook = setCellContent(workbook, workbook.activeSheetId, "A1", "one");
    workbook = setCellContent(workbook, workbook.activeSheetId, "B1", "two");

    const matrix = copyRange(workbook, workbook.activeSheetId, range("A1", "B1"));
    const pasted = pasteMatrix(workbook, workbook.activeSheetId, "A2", matrix);

    expect(matrix).toEqual([["one", "two"]]);
    expect(getCellContent(pasted, pasted.activeSheetId, "A2")).toBe("one");
    expect(getCellContent(pasted, pasted.activeSheetId, "B2")).toBe("two");
  });

  it("copies and pastes formulas, formats, validations, comments, and hyperlinks as a rich range", () => {
    let workbook = createBlankWorkbook();
    const sheetId = workbook.activeSheetId;
    workbook = setCellContent(workbook, sheetId, "A1", "10");
    workbook = setCellContent(workbook, sheetId, "B1", "=A1+$A$1");
    workbook = setCellFormat(workbook, sheetId, range("B1"), {
      bold: true,
      backgroundColor: "#fff1d6",
      numberFormat: "currency"
    });
    workbook = setCellValidation(workbook, sheetId, range("B1"), { type: "number", min: 1, max: 100 });
    workbook = setCellComment(workbook, sheetId, "B1", "Review this calculation");
    workbook = setCellHyperlink(workbook, sheetId, "B1", "https://example.com/report");

    const richRange = copyRichRange(workbook, sheetId, range("B1"));
    const preview = previewRichPaste(richRange, "C2");
    const pasted = pasteRichRange(workbook, sheetId, "C2", richRange);

    expect(preview).toMatchObject([{ address: "C2", content: "=B2+$A$1" }]);
    expect(getCellContent(pasted, sheetId, "C2")).toBe("=B2+$A$1");
    expect(getCellFormat(pasted, sheetId, "C2")).toEqual({
      bold: true,
      backgroundColor: "#fff1d6",
      numberFormat: "currency"
    });
    expect(getCellValidation(pasted, sheetId, "C2")).toEqual({ type: "number", min: 1, max: 100 });
    expect(getCellComment(pasted, sheetId, "C2")).toBe("Review this calculation");
    expect(getCellHyperlink(pasted, sheetId, "C2")).toBe("https://example.com/report");
  });

  it("moves a rich range without rewriting formulas and clears source metadata", () => {
    let workbook = createBlankWorkbook();
    const sheetId = workbook.activeSheetId;
    workbook = setCellContent(workbook, sheetId, "A1", "10");
    workbook = setCellContent(workbook, sheetId, "B1", "=A1");
    workbook = setCellFormat(workbook, sheetId, range("B1"), {
      bold: true,
      backgroundColor: "#fff1d6"
    });
    workbook = setCellValidation(workbook, sheetId, range("B1"), { type: "number", min: 1, max: 100 });
    workbook = setCellComment(workbook, sheetId, "B1", "Move this formula");
    workbook = setCellHyperlink(workbook, sheetId, "B1", "https://example.com/source");

    const moved = moveRichRange(workbook, sheetId, sheetId, "C1", copyRichRange(workbook, sheetId, range("B1")));

    expect(getCellContent(moved, sheetId, "B1")).toBeNull();
    expect(getCellFormat(moved, sheetId, "B1")).toEqual({});
    expect(getCellValidation(moved, sheetId, "B1")).toBeNull();
    expect(getCellComment(moved, sheetId, "B1")).toBeNull();
    expect(getCellHyperlink(moved, sheetId, "B1")).toBeNull();
    expect(getCellContent(moved, sheetId, "C1")).toBe("=A1");
    expect(getCellFormat(moved, sheetId, "C1")).toEqual({ bold: true, backgroundColor: "#fff1d6" });
    expect(getCellValidation(moved, sheetId, "C1")).toEqual({ type: "number", min: 1, max: 100 });
    expect(getCellComment(moved, sheetId, "C1")).toBe("Move this formula");
    expect(getCellHyperlink(moved, sheetId, "C1")).toBe("https://example.com/source");
  });

  it("sorts comments, hyperlinks, and validations with their row data", () => {
    let workbook = createBlankWorkbook();
    const sheetId = workbook.activeSheetId;
    workbook = setCellContent(workbook, sheetId, "A1", "Delta");
    workbook = setCellContent(workbook, sheetId, "B1", "4");
    workbook = setCellContent(workbook, sheetId, "A2", "Alpha");
    workbook = setCellContent(workbook, sheetId, "B2", "1");
    workbook = setCellComment(workbook, sheetId, "A1", "Delta note");
    workbook = setCellHyperlink(workbook, sheetId, "A1", "https://example.com/delta");
    workbook = setCellValidation(workbook, sheetId, range("B1"), { type: "number", min: 1, max: 10 });

    const sorted = sortRange(workbook, sheetId, range("A1", "B2"), "asc");

    expect(getCellContent(sorted, sheetId, "A1")).toBe("Alpha");
    expect(getCellContent(sorted, sheetId, "A2")).toBe("Delta");
    expect(getCellComment(sorted, sheetId, "A1")).toBeNull();
    expect(getCellHyperlink(sorted, sheetId, "A1")).toBeNull();
    expect(getCellValidation(sorted, sheetId, "B1")).toBeNull();
    expect(getCellComment(sorted, sheetId, "A2")).toBe("Delta note");
    expect(getCellHyperlink(sorted, sheetId, "A2")).toBe("https://example.com/delta");
    expect(getCellValidation(sorted, sheetId, "B2")).toEqual({ type: "number", min: 1, max: 10 });
  });

  it("sorts by a pre-sort evaluated snapshot and translates moved relative formulas", () => {
    let workbook = createBlankWorkbook();
    const sheetId = workbook.activeSheetId;
    workbook = setCellContent(workbook, sheetId, "A1", 20);
    workbook = setCellContent(workbook, sheetId, "B1", "=A1");
    workbook = setCellContent(workbook, sheetId, "A2", 10);
    workbook = setCellContent(workbook, sheetId, "B2", "=A2");
    const beforeSort = createFormulaEngine(workbook);

    const sorted = sortRange(workbook, sheetId, range("A1", "B2"), {
      direction: "asc",
      sortColumn: 1,
      readValue: (address) => beforeSort.getComputedValue(sheetId, address)
    });
    const afterSort = createFormulaEngine(sorted);

    expect(getCellContent(sorted, sheetId, "A1")).toBe(10);
    expect(getCellContent(sorted, sheetId, "B1")).toBe("=A1");
    expect(afterSort.getComputedValue(sheetId, "B1")).toBe(10);
    expect(getCellContent(sorted, sheetId, "A2")).toBe(20);
    expect(getCellContent(sorted, sheetId, "B2")).toBe("=A2");
    expect(afterSort.getComputedValue(sheetId, "B2")).toBe(20);
  });

  it("orders text deterministically without the runtime's default locale", () => {
    let workbook = createBlankWorkbook();
    const sheetId = workbook.activeSheetId;
    workbook = setCellContent(workbook, sheetId, "A1", "ä");
    workbook = setCellContent(workbook, sheetId, "A2", "z");

    const sorted = sortRange(workbook, sheetId, range("A1", "A2"), "asc");

    expect(getCellContent(sorted, sheetId, "A1")).toBe("z");
    expect(getCellContent(sorted, sheetId, "A2")).toBe("ä");
  });

  it("removes duplicate rows from a selected range and shifts row metadata", () => {
    let workbook = createBlankWorkbook();
    const sheetId = workbook.activeSheetId;
    workbook = pasteMatrix(workbook, sheetId, "A1", [
      ["East", "Hardware"],
      ["West", "Software"],
      ["East", "Hardware"],
      ["North", "Services"],
      ["West", "Software"]
    ]);
    workbook = setCellFormat(workbook, sheetId, range("A4", "B4"), { bold: true, backgroundColor: "#eaf7f2" });
    workbook = setCellComment(workbook, sheetId, "A4", "Keep this row");
    workbook = setCellHyperlink(workbook, sheetId, "B4", "https://example.com/services");
    workbook = setCellValidation(workbook, sheetId, range("A4"), { type: "list", values: ["North", "South"] });

    const result = removeDuplicateRows(workbook, sheetId, range("A1", "B5"));

    expect(result.removedCount).toBe(2);
    expect(getCellContent(result.workbook, sheetId, "A1")).toBe("East");
    expect(getCellContent(result.workbook, sheetId, "B1")).toBe("Hardware");
    expect(getCellContent(result.workbook, sheetId, "A2")).toBe("West");
    expect(getCellContent(result.workbook, sheetId, "B2")).toBe("Software");
    expect(getCellContent(result.workbook, sheetId, "A3")).toBe("North");
    expect(getCellContent(result.workbook, sheetId, "B3")).toBe("Services");
    expect(getCellContent(result.workbook, sheetId, "A4")).toBeNull();
    expect(getCellContent(result.workbook, sheetId, "B5")).toBeNull();
    expect(getCellFormat(result.workbook, sheetId, "A3")).toEqual({ bold: true, backgroundColor: "#eaf7f2" });
    expect(getCellComment(result.workbook, sheetId, "A3")).toBe("Keep this row");
    expect(getCellHyperlink(result.workbook, sheetId, "B3")).toBe("https://example.com/services");
    expect(getCellValidation(result.workbook, sheetId, "A3")).toEqual({ type: "list", values: ["North", "South"] });
    expect(getCellFormat(result.workbook, sheetId, "A4")).toEqual({});
    expect(getCellComment(result.workbook, sheetId, "A4")).toBeNull();
    expect(getCellValidation(result.workbook, sheetId, "A4")).toBeNull();
  });

  it("keeps AutoFilter headers in place when sorting table data", () => {
    let workbook = createBlankWorkbook();
    const sheetId = workbook.activeSheetId;
    workbook = setCellContent(workbook, sheetId, "A1", "Region");
    workbook = setCellContent(workbook, sheetId, "B1", "Sales");
    workbook = setCellContent(workbook, sheetId, "A2", "West");
    workbook = setCellContent(workbook, sheetId, "B2", 10);
    workbook = setCellContent(workbook, sheetId, "A3", "East");
    workbook = setCellContent(workbook, sheetId, "B3", 20);
    workbook = setCellContent(workbook, sheetId, "A4", "North");
    workbook = setCellContent(workbook, sheetId, "B4", 15);
    workbook = addSheetFilter(workbook, sheetId, range("A1", "B4"), {
      column: 0,
      operator: "contains",
      value: "",
      hasHeader: true
    });

    const sorted = sortRange(workbook, sheetId, range("A1", "B4"), "asc");

    expect(getCellContent(sorted, sheetId, "A1")).toBe("Region");
    expect(getCellContent(sorted, sheetId, "B1")).toBe("Sales");
    expect(getCellContent(sorted, sheetId, "A2")).toBe("East");
    expect(getCellContent(sorted, sheetId, "B2")).toBe(20);
    expect(getCellContent(sorted, sheetId, "A3")).toBe("North");
    expect(getCellContent(sorted, sheetId, "B3")).toBe(15);
    expect(getCellContent(sorted, sheetId, "A4")).toBe("West");
    expect(getCellContent(sorted, sheetId, "B4")).toBe(10);
  });

  it("stores wrapped text formatting and preserves it through rich paste", () => {
    let workbook = createBlankWorkbook();
    const sheetId = workbook.activeSheetId;
    workbook = setCellContent(workbook, sheetId, "A1", "Long wrapped note");
    workbook = setCellFormat(workbook, sheetId, range("A1"), { wrapText: true });

    expect(getCellFormat(workbook, sheetId, "A1")).toMatchObject({ wrapText: true });

    const richRange = copyRichRange(workbook, sheetId, range("A1"));
    const pasted = pasteRichRange(workbook, sheetId, "B2", richRange);

    expect(getCellContent(pasted, sheetId, "B2")).toBe("Long wrapped note");
    expect(getCellFormat(pasted, sheetId, "B2")).toMatchObject({ wrapText: true });
  });

  it("rich paste clears target metadata when the copied source has none", () => {
    let workbook = createBlankWorkbook();
    const sheetId = workbook.activeSheetId;
    workbook = setCellContent(workbook, sheetId, "A1", "Plain");
    workbook = setCellContent(workbook, sheetId, "B2", "Styled");
    workbook = setCellFormat(workbook, sheetId, range("B2"), { bold: true });
    workbook = setCellValidation(workbook, sheetId, range("B2"), { type: "list", values: ["Styled"] });
    workbook = setCellComment(workbook, sheetId, "B2", "Old note");
    workbook = setCellHyperlink(workbook, sheetId, "B2", "https://example.com/old");

    const pasted = pasteRichRange(workbook, sheetId, "B2", copyRichRange(workbook, sheetId, range("A1")));

    expect(getCellContent(pasted, sheetId, "B2")).toBe("Plain");
    expect(getCellFormat(pasted, sheetId, "B2")).toEqual({});
    expect(getCellValidation(pasted, sheetId, "B2")).toBeNull();
    expect(getCellComment(pasted, sheetId, "B2")).toBeNull();
    expect(getCellHyperlink(pasted, sheetId, "B2")).toBeNull();
  });

  it("paste special can paste displayed values without changing target formatting", () => {
    let workbook = createBlankWorkbook();
    const sheetId = workbook.activeSheetId;
    workbook = setCellContent(workbook, sheetId, "A1", "10");
    workbook = setCellContent(workbook, sheetId, "B1", "=A1");
    workbook = setCellContent(workbook, sheetId, "C1", "Keep style");
    workbook = setCellFormat(workbook, sheetId, range("C1"), { italic: true, backgroundColor: "#eaf7f2" });

    const clipboard = copyRichRange(workbook, sheetId, range("B1"), { getDisplayValue: () => "10" });
    const pasted = pasteRichRange(workbook, sheetId, "C1", clipboard, { mode: "values" });

    expect(getCellContent(pasted, sheetId, "C1")).toBe("10");
    expect(getCellFormat(pasted, sheetId, "C1")).toEqual({ italic: true, backgroundColor: "#eaf7f2" });
  });

  it("paste special can paste formats without changing target content", () => {
    let workbook = createBlankWorkbook();
    const sheetId = workbook.activeSheetId;
    workbook = setCellContent(workbook, sheetId, "A1", "Styled source");
    workbook = setCellFormat(workbook, sheetId, range("A1"), { bold: true, textColor: "#17634a" });
    workbook = setCellContent(workbook, sheetId, "C1", "Keep content");

    const pasted = pasteRichRange(workbook, sheetId, "C1", copyRichRange(workbook, sheetId, range("A1")), {
      mode: "formats"
    });

    expect(getCellContent(pasted, sheetId, "C1")).toBe("Keep content");
    expect(getCellFormat(pasted, sheetId, "C1")).toEqual({ bold: true, textColor: "#17634a" });
  });

  it("paste special can transpose a copied range and adjust formulas", () => {
    let workbook = createBlankWorkbook();
    const sheetId = workbook.activeSheetId;
    workbook = setCellContent(workbook, sheetId, "A1", "A");
    workbook = setCellContent(workbook, sheetId, "B1", "=A1");
    workbook = setCellContent(workbook, sheetId, "A2", "C");
    workbook = setCellContent(workbook, sheetId, "B2", "D");

    const pasted = pasteRichRange(workbook, sheetId, "D1", copyRichRange(workbook, sheetId, range("A1", "B2")), {
      mode: "transpose"
    });

    expect(getCellContent(pasted, sheetId, "D1")).toBe("A");
    expect(getCellContent(pasted, sheetId, "E1")).toBe("C");
    expect(getCellContent(pasted, sheetId, "D2")).toBe("=C2");
    expect(getCellContent(pasted, sheetId, "E2")).toBe("D");
  });

  it("inserts and deletes rows while shifting cells, formats, and formula references", () => {
    let workbook = createBlankWorkbook();
    const sheetId = workbook.activeSheetId;
    workbook = setCellContent(workbook, sheetId, "A1", "Item");
    workbook = setCellContent(workbook, sheetId, "A2", "10");
    workbook = setCellContent(workbook, sheetId, "A3", "20");
    workbook = setCellContent(workbook, sheetId, "B4", "=SUM(A2:A3)");
    workbook = setCellFormat(workbook, sheetId, range("A2"), { bold: true });

    const inserted = insertRows(workbook, sheetId, 1);

    expect(getCellContent(inserted, sheetId, "A1")).toBe("Item");
    expect(getCellContent(inserted, sheetId, "A2")).toBeNull();
    expect(getCellContent(inserted, sheetId, "A3")).toBe("10");
    expect(getCellContent(inserted, sheetId, "A4")).toBe("20");
    expect(getCellContent(inserted, sheetId, "B5")).toBe("=SUM(A3:A4)");
    expect(getCellFormat(inserted, sheetId, "A3")).toEqual({ bold: true });

    const deleted = deleteRows(inserted, sheetId, 1);

    expect(getCellContent(deleted, sheetId, "A2")).toBe("10");
    expect(getCellContent(deleted, sheetId, "A3")).toBe("20");
    expect(getCellContent(deleted, sheetId, "B4")).toBe("=SUM(A2:A3)");
    expect(getCellFormat(deleted, sheetId, "A2")).toEqual({ bold: true });
  });

  it("inserts and deletes columns while shifting cells, formats, and formula references", () => {
    let workbook = createBlankWorkbook();
    const sheetId = workbook.activeSheetId;
    workbook = setCellContent(workbook, sheetId, "A1", "Product");
    workbook = setCellContent(workbook, sheetId, "B1", "10");
    workbook = setCellContent(workbook, sheetId, "C1", "=B1*2");
    workbook = setCellFormat(workbook, sheetId, range("B1"), { backgroundColor: "#eaf7f2" });

    const inserted = insertColumns(workbook, sheetId, 1);

    expect(getCellContent(inserted, sheetId, "A1")).toBe("Product");
    expect(getCellContent(inserted, sheetId, "B1")).toBeNull();
    expect(getCellContent(inserted, sheetId, "C1")).toBe("10");
    expect(getCellContent(inserted, sheetId, "D1")).toBe("=C1*2");
    expect(getCellFormat(inserted, sheetId, "C1")).toEqual({ backgroundColor: "#eaf7f2" });

    const deleted = deleteColumns(inserted, sheetId, 1);

    expect(getCellContent(deleted, sheetId, "B1")).toBe("10");
    expect(getCellContent(deleted, sheetId, "C1")).toBe("=B1*2");
    expect(getCellFormat(deleted, sheetId, "B1")).toEqual({ backgroundColor: "#eaf7f2" });
  });

  it("projects partially surviving metadata across a column deletion boundary", () => {
    const initial = createBlankWorkbook();
    const sheetId = initial.activeSheetId;
    const partial = range("A1", "C3");
    const dropped = range("A10", "A12");
    const workbook = {
      ...initial,
      namedRanges: [
        { name: "Partial", sheetId, range: range("A1", "C1") },
        { name: "Dropped", sheetId, range: range("A10") }
      ],
      sheets: initial.sheets.map((sheet) => ({
        ...sheet,
        autoFilterRange: partial,
        conditionalFormats: [
          { id: "cf-partial", range: partial, condition: { type: "blank" as const }, format: { bold: true } },
          { id: "cf-dropped", range: dropped, condition: { type: "blank" as const }, format: { bold: true } }
        ],
        filters: [
          { id: "filter-partial", range: partial, column: 1, operator: "equals" as const, value: "x" },
          { id: "filter-dropped", range: dropped, column: 0, operator: "equals" as const, value: "x" }
        ],
        charts: [
          { id: "chart-partial", title: "Partial", type: "bar" as const, range: partial, anchor: { row: 6, column: 0 } },
          { id: "chart-dropped", title: "Dropped", type: "bar" as const, range: dropped, anchor: { row: 14, column: 0 } }
        ],
        merges: [
          { id: "merge-partial", range: range("A5", "C5") },
          { id: "merge-dropped", range: range("A20", "A21") }
        ]
      }))
    };

    const deleted = deleteColumns(workbook, sheetId, 0, 1);
    const sheet = deleted.sheets[0];

    expect(deleted.namedRanges).toEqual([
      { name: "Partial", sheetId, range: range("A1", "B1") }
    ]);
    expect(sheet.conditionalFormats).toMatchObject([
      { id: "cf-partial", range: range("A1", "B3") }
    ]);
    expect(sheet.filters).toMatchObject([
      { id: "filter-partial", range: range("A1", "B3"), column: 0 }
    ]);
    expect(sheet.autoFilterRange).toEqual(range("A1", "B3"));
    expect(sheet.charts).toMatchObject([
      { id: "chart-partial", range: range("A1", "B3"), anchor: { row: 6, column: 0 } }
    ]);
    expect(sheet.merges).toEqual([
      { id: "merge-partial", range: range("A5", "B5") }
    ]);
  });

  it("projects partially surviving metadata across a row deletion boundary", () => {
    const initial = createBlankWorkbook();
    const sheetId = initial.activeSheetId;
    const partial = range("A1", "C3");
    const dropped = range("A1", "C1");
    const workbook = {
      ...initial,
      namedRanges: [
        { name: "Partial", sheetId, range: range("A1", "A3") },
        { name: "Dropped", sheetId, range: range("D1") }
      ],
      sheets: initial.sheets.map((sheet) => ({
        ...sheet,
        autoFilterRange: partial,
        conditionalFormats: [
          { id: "cf-partial", range: partial, condition: { type: "blank" as const }, format: { bold: true } },
          { id: "cf-dropped", range: dropped, condition: { type: "blank" as const }, format: { bold: true } }
        ],
        filters: [
          { id: "filter-partial", range: partial, column: 1, operator: "equals" as const, value: "x" },
          { id: "filter-dropped", range: dropped, column: 0, operator: "equals" as const, value: "x" }
        ],
        charts: [
          { id: "chart-partial", title: "Partial", type: "bar" as const, range: partial, anchor: { row: 0, column: 6 } },
          { id: "chart-dropped", title: "Dropped", type: "bar" as const, range: dropped, anchor: { row: 0, column: 10 } }
        ],
        merges: [
          { id: "merge-partial", range: range("E1", "E3") },
          { id: "merge-dropped", range: range("F1", "G1") }
        ]
      }))
    };

    const deleted = deleteRows(workbook, sheetId, 0, 1);
    const sheet = deleted.sheets[0];

    expect(deleted.namedRanges).toEqual([
      { name: "Partial", sheetId, range: range("A1", "A2") }
    ]);
    expect(sheet.conditionalFormats).toMatchObject([
      { id: "cf-partial", range: range("A1", "C2") }
    ]);
    expect(sheet.filters).toMatchObject([
      { id: "filter-partial", range: range("A1", "C2"), column: 1 }
    ]);
    expect(sheet.autoFilterRange).toEqual(range("A1", "C2"));
    expect(sheet.charts).toMatchObject([
      { id: "chart-partial", range: range("A1", "C2"), anchor: { row: 0, column: 6 } }
    ]);
    expect(sheet.merges).toEqual([
      { id: "merge-partial", range: range("E1", "E2") }
    ]);
  });

  it.each([
    ["insertRows", insertRows],
    ["deleteRows", deleteRows],
    ["insertColumns", insertColumns],
    ["deleteColumns", deleteColumns]
  ] as const)("%s refuses structured-table worksheet edits", (_name, edit) => {
    let workbook = createBlankWorkbook();
    const sheetId = workbook.activeSheetId;
    const table: StructuredTable = {
      id: "table-1",
      name: "TableOne",
      sheetId,
      range: { start: { row: 0, column: 0 }, end: { row: 1, column: 0 } },
      headerRow: true,
      totalsRow: false,
      columns: [{ id: "table-column-1", name: "Name", sheetColumn: 0 }],
      rowIds: ["table-row-1"]
    };
    workbook = { ...workbook, tables: [table] };

    expect(() => edit(workbook, sheetId, 0, 1)).toThrow(
      "Use WorkbookSession.dispatch for structured-table worksheet edits"
    );
  });

  it("does not rewrite LOG10, scientific notation, or quoted A1-like text during row insertion", () => {
    let workbook = createBlankWorkbook();
    const sheetId = workbook.activeSheetId;
    workbook = setCellContent(workbook, sheetId, "B1", '=LOG10(A1)+1E5+Q1_TOTAL+"A1 ""B2"""');

    const inserted = insertRows(workbook, sheetId, 0);

    expect(getCellContent(inserted, sheetId, "B2")).toBe('=LOG10(A2)+1E5+Q1_TOTAL+"A1 ""B2"""');
  });

  it("updates formulas on other sheets that reference the structurally edited sheet", () => {
    let workbook = createBlankWorkbook();
    const dataSheetId = workbook.activeSheetId;
    workbook = renameSheet(workbook, dataSheetId, "Director's Plan");
    workbook = addSheet(workbook, "Summary");
    const summarySheetId = workbook.activeSheetId;

    workbook = setCellContent(workbook, dataSheetId, "A2", "10");
    workbook = setCellContent(workbook, dataSheetId, "C1", "=A2+Summary!A2");
    workbook = setCellContent(
      workbook,
      summarySheetId,
      "B2",
      "=SUM('Director''s Plan'!$A$2:$A$4)+A2"
    );
    workbook = setCellFormat(workbook, summarySheetId, range("B2"), { bold: true });
    workbook = setCellComment(workbook, summarySheetId, "B2", "Keep this metadata in place");

    const inserted = insertRows(workbook, dataSheetId, 1);

    expect(getCellContent(inserted, dataSheetId, "A2")).toBeNull();
    expect(getCellContent(inserted, dataSheetId, "A3")).toBe("10");
    expect(getCellContent(inserted, dataSheetId, "C1")).toBe("=A3+Summary!A2");
    expect(getCellContent(inserted, summarySheetId, "B2")).toBe(
      "=SUM('Director''s Plan'!$A$3:$A$5)+A2"
    );
    expect(getCellContent(inserted, summarySheetId, "B3")).toBeNull();
    expect(getCellFormat(inserted, summarySheetId, "B2")).toEqual({ bold: true });
    expect(getCellComment(inserted, summarySheetId, "B2")).toBe("Keep this metadata in place");
    expect(getCellComment(inserted, summarySheetId, "B3")).toBeNull();
  });

  it("stores, clamps, and shifts row heights and column widths", () => {
    let workbook = createBlankWorkbook();
    const sheetId = workbook.activeSheetId;
    workbook = setColumnWidth(workbook, sheetId, 1, 144);
    workbook = setRowHeight(workbook, sheetId, 2, 42);

    expect(getColumnWidth(workbook, sheetId, 1)).toBe(144);
    expect(getRowHeight(workbook, sheetId, 2)).toBe(42);

    const inserted = insertColumns(insertRows(workbook, sheetId, 1), sheetId, 1);
    expect(getColumnWidth(inserted, sheetId, 2)).toBe(144);
    expect(getRowHeight(inserted, sheetId, 3)).toBe(42);

    const clamped = setRowHeight(setColumnWidth(workbook, sheetId, 0, 8), sheetId, 0, 8);
    expect(getColumnWidth(clamped, sheetId, 0)).toBe(56);
    expect(getRowHeight(clamped, sheetId, 0)).toBe(22);
  });

  it("hides, unhides, duplicates, and shifts rows and columns", () => {
    let workbook = createBlankWorkbook();
    const sheetId = workbook.activeSheetId;

    workbook = setRowsHidden(workbook, sheetId, 1, 2, true);
    workbook = setColumnsHidden(workbook, sheetId, 1, 1, true);

    expect(isRowHidden(workbook, sheetId, 0)).toBe(false);
    expect(isRowHidden(workbook, sheetId, 1)).toBe(true);
    expect(isRowHidden(workbook, sheetId, 2)).toBe(true);
    expect(isColumnHidden(workbook, sheetId, 1)).toBe(true);

    const shifted = insertRows(insertColumns(workbook, sheetId, 1), sheetId, 1);
    expect(isRowHidden(shifted, sheetId, 1)).toBe(false);
    expect(isRowHidden(shifted, sheetId, 2)).toBe(true);
    expect(isColumnHidden(shifted, sheetId, 1)).toBe(false);
    expect(isColumnHidden(shifted, sheetId, 2)).toBe(true);

    const duplicated = duplicateSheet(shifted, sheetId);
    expect(isRowHidden(duplicated, duplicated.activeSheetId, 2)).toBe(true);
    expect(isColumnHidden(duplicated, duplicated.activeSheetId, 2)).toBe(true);

    const cleared = clearHiddenRowsAndColumns(duplicated, duplicated.activeSheetId);
    expect(isRowHidden(cleared, cleared.activeSheetId, 2)).toBe(false);
    expect(isColumnHidden(cleared, cleared.activeSheetId, 2)).toBe(false);

    const removed = deleteRows(setRowsHidden(workbook, sheetId, 4, 4, true), sheetId, 4);
    expect(isRowHidden(removed, sheetId, 4)).toBe(false);
  });

  it("hides sheet tabs while keeping at least one visible active sheet", () => {
    let workbook = createBlankWorkbook();
    const firstSheetId = workbook.activeSheetId;
    workbook = addSheet(workbook, "Data");
    const secondSheetId = workbook.activeSheetId;
    workbook = addSheet(workbook, "Archive");
    const thirdSheetId = workbook.activeSheetId;

    workbook = setSheetHidden(workbook, thirdSheetId, true);

    expect(isSheetHidden(workbook, thirdSheetId)).toBe(true);
    expect(workbook.activeSheetId).toBe(firstSheetId);

    workbook = setSheetHidden(workbook, firstSheetId, true);

    expect(isSheetHidden(workbook, firstSheetId)).toBe(true);
    expect(workbook.activeSheetId).toBe(secondSheetId);

    expect(setSheetHidden(workbook, secondSheetId, true)).toBe(workbook);
    expect(isSheetHidden(workbook, secondSheetId)).toBe(false);

    const unhidden = unhideAllSheets(workbook);
    expect(isSheetHidden(unhidden, firstSheetId)).toBe(false);
    expect(isSheetHidden(unhidden, thirdSheetId)).toBe(false);
    expect(unhideAllSheets(unhidden)).toBe(unhidden);
  });

  it("adds, clears, duplicates, and shifts cell hyperlinks", () => {
    let workbook = createBlankWorkbook();
    const sheetId = workbook.activeSheetId;

    workbook = setCellHyperlink(workbook, sheetId, "B2", "https://example.com/report");
    expect(getCellHyperlink(workbook, sheetId, "B2")).toBe("https://example.com/report");

    const shifted = insertColumns(insertRows(workbook, sheetId, 1), sheetId, 1);
    expect(getCellHyperlink(shifted, sheetId, "B2")).toBeNull();
    expect(getCellHyperlink(shifted, sheetId, "C3")).toBe("https://example.com/report");

    const duplicated = duplicateSheet(shifted, sheetId);
    expect(getCellHyperlink(duplicated, duplicated.activeSheetId, "C3")).toBe("https://example.com/report");

    const cleared = clearRange(shifted, sheetId, range("C3"));
    expect(getCellHyperlink(cleared, sheetId, "C3")).toBeNull();
    expect(setCellHyperlink(cleared, sheetId, "C3", "")).toBe(cleared);
  });

  it("clears hyperlinks from a range without clearing content or metadata", () => {
    let workbook = createBlankWorkbook();
    const sheetId = workbook.activeSheetId;

    workbook = setCellContent(workbook, sheetId, "A1", "Report");
    workbook = setCellFormat(workbook, sheetId, range("A1"), { bold: true, backgroundColor: "#eaf7f2" });
    workbook = setCellValidation(workbook, sheetId, range("A1"), { type: "list", values: ["Report"] });
    workbook = setCellComment(workbook, sheetId, "A1", "Keep the note");
    workbook = setCellHyperlink(workbook, sheetId, "A1", "https://example.com/report");
    workbook = setCellHyperlink(workbook, sheetId, "B1", "https://example.com/keep");

    const cleared = clearCellHyperlinks(workbook, sheetId, range("A1"));

    expect(getCellContent(cleared, sheetId, "A1")).toBe("Report");
    expect(getCellFormat(cleared, sheetId, "A1")).toEqual({ bold: true, backgroundColor: "#eaf7f2" });
    expect(getCellValidation(cleared, sheetId, "A1")).toEqual({ type: "list", values: ["Report"] });
    expect(getCellComment(cleared, sheetId, "A1")).toBe("Keep the note");
    expect(getCellHyperlink(cleared, sheetId, "A1")).toBeNull();
    expect(getCellHyperlink(cleared, sheetId, "B1")).toBe("https://example.com/keep");
    expect(clearCellHyperlinks(cleared, sheetId, range("A1"))).toBe(cleared);
  });

  it("clears all cell content, formats, conditional formats, and hyperlinks in a range", () => {
    let workbook = createBlankWorkbook();
    const sheetId = workbook.activeSheetId;

    workbook = setCellContent(workbook, sheetId, "A1", "Report");
    workbook = setCellFormat(workbook, sheetId, range("A1"), { bold: true, backgroundColor: "#eaf7f2" });
    workbook = setCellValidation(workbook, sheetId, range("A1"), { type: "list", values: ["Report"] });
    workbook = setCellComment(workbook, sheetId, "A1", "Keep the note");
    workbook = setCellHyperlink(workbook, sheetId, "A1", "https://example.com/report");
    workbook = addConditionalFormatRule(workbook, sheetId, range("A1"), {
      condition: { type: "textContains", value: "Report" },
      format: { backgroundColor: "#fff1d6", textColor: "#8a4b00" }
    });
    workbook = setCellContent(workbook, sheetId, "B1", "Keep");
    workbook = setCellHyperlink(workbook, sheetId, "B1", "https://example.com/keep");

    const cleared = clearRangeAll(workbook, sheetId, range("A1"));

    expect(getCellContent(cleared, sheetId, "A1")).toBeNull();
    expect(getCellFormat(cleared, sheetId, "A1")).toEqual({});
    expect(getCellHyperlink(cleared, sheetId, "A1")).toBeNull();
    expect(getCellConditionalFormatRules(cleared, sheetId, "A1")).toHaveLength(0);
    expect(getCellValidation(cleared, sheetId, "A1")).toEqual({ type: "list", values: ["Report"] });
    expect(getCellComment(cleared, sheetId, "A1")).toBe("Keep the note");
    expect(getCellContent(cleared, sheetId, "B1")).toBe("Keep");
    expect(getCellHyperlink(cleared, sheetId, "B1")).toBe("https://example.com/keep");
    expect(clearRangeAll(cleared, sheetId, range("A1"))).toBe(cleared);
  });

  it("fills down and right from the selected edge", () => {
    let workbook = createBlankWorkbook();
    workbook = setCellContent(workbook, workbook.activeSheetId, "A1", "seed");
    workbook = setCellContent(workbook, workbook.activeSheetId, "B1", "=A1");

    const down = fillDown(workbook, workbook.activeSheetId, range("A1", "A3"));
    expect(getCellContent(down, down.activeSheetId, "A2")).toBe("seed");
    expect(getCellContent(down, down.activeSheetId, "A3")).toBe("seed");

    const right = fillRight(workbook, workbook.activeSheetId, range("B1", "D1"));
    expect(getCellContent(right, right.activeSheetId, "C1")).toBe("=B1");
    expect(getCellContent(right, right.activeSheetId, "D1")).toBe("=C1");
  });

  it("fills formulas with relative references adjusted for each target cell", () => {
    let workbook = createBlankWorkbook();
    const sheetId = workbook.activeSheetId;
    workbook = setCellContent(workbook, sheetId, "A1", "10");
    workbook = setCellContent(workbook, sheetId, "A2", "20");
    workbook = setCellContent(workbook, sheetId, "A3", "30");
    workbook = setCellContent(workbook, sheetId, "B1", "=A1+$A$1");

    const down = fillDown(workbook, sheetId, range("B1", "B3"));
    expect(getCellContent(down, sheetId, "B2")).toBe("=A2+$A$1");
    expect(getCellContent(down, sheetId, "B3")).toBe("=A3+$A$1");

    const right = fillRight(workbook, sheetId, range("B1", "D1"));
    expect(getCellContent(right, sheetId, "C1")).toBe("=B1+$A$1");
    expect(getCellContent(right, sheetId, "D1")).toBe("=C1+$A$1");
  });

  it("auto-fills numeric series and formulas down from the selected range", () => {
    let workbook = createBlankWorkbook();
    const sheetId = workbook.activeSheetId;
    workbook = setCellContent(workbook, sheetId, "A1", "1");
    workbook = setCellContent(workbook, sheetId, "A2", "3");
    workbook = setCellContent(workbook, sheetId, "B1", "=A1");

    const series = autoFillRange(workbook, sheetId, range("A1", "A2"), range("A1", "A5"));
    expect(getCellContent(series, sheetId, "A3")).toBe(5);
    expect(getCellContent(series, sheetId, "A4")).toBe(7);
    expect(getCellContent(series, sheetId, "A5")).toBe(9);

    const formulas = autoFillRange(workbook, sheetId, range("B1"), range("B1", "B3"));
    expect(getCellContent(formulas, sheetId, "B2")).toBe("=A2");
    expect(getCellContent(formulas, sheetId, "B3")).toBe("=A3");
  });

  it("auto-fills numeric series right from the selected range", () => {
    let workbook = createBlankWorkbook();
    const sheetId = workbook.activeSheetId;
    workbook = setCellContent(workbook, sheetId, "A1", 2);
    workbook = setCellContent(workbook, sheetId, "B1", 5);

    const filled = autoFillRange(workbook, sheetId, range("A1", "B1"), range("A1", "E1"));
    expect(getCellContent(filled, sheetId, "C1")).toBe(8);
    expect(getCellContent(filled, sheetId, "D1")).toBe(11);
    expect(getCellContent(filled, sheetId, "E1")).toBe(14);
  });

  it("auto-fills ISO date series by day intervals", () => {
    let workbook = createBlankWorkbook();
    const sheetId = workbook.activeSheetId;
    workbook = setCellContent(workbook, sheetId, "A1", "2026-06-28");
    workbook = setCellContent(workbook, sheetId, "A2", "2026-06-30");
    workbook = setCellContent(workbook, sheetId, "B1", "2026-07-04");

    const stepped = autoFillRange(workbook, sheetId, range("A1", "A2"), range("A1", "A4"));
    expect(getCellContent(stepped, sheetId, "A3")).toBe("2026-07-02");
    expect(getCellContent(stepped, sheetId, "A4")).toBe("2026-07-04");

    const single = autoFillRange(workbook, sheetId, range("B1"), range("B1", "B3"));
    expect(getCellContent(single, sheetId, "B2")).toBe("2026-07-05");
    expect(getCellContent(single, sheetId, "B3")).toBe("2026-07-06");
  });

  it("auto-fills source formats and validation rules into target cells", () => {
    let workbook = createBlankWorkbook();
    const sheetId = workbook.activeSheetId;
    workbook = setCellContent(workbook, sheetId, "A1", "Approved");
    workbook = setCellContent(workbook, sheetId, "A2", "Rejected");
    workbook = setCellFormat(workbook, sheetId, range("A1"), { bold: true, backgroundColor: "#eaf7f2" });
    workbook = setCellFormat(workbook, sheetId, range("A2"), { italic: true, textColor: "#d44949" });
    workbook = setCellValidation(workbook, sheetId, range("A1"), { type: "list", values: ["Approved", "Rejected"] });
    workbook = setCellFormat(workbook, sheetId, range("A3"), { wrapText: true, textColor: "#111827" });
    workbook = setCellFormat(workbook, sheetId, range("A4"), { bold: true, backgroundColor: "#1f6feb" });
    workbook = setCellValidation(workbook, sheetId, range("A3", "A4"), { type: "number", min: 1, max: 10 });

    const filled = autoFillRange(workbook, sheetId, range("A1", "A2"), range("A1", "A4"));

    expect(getCellContent(filled, sheetId, "A3")).toBe("Approved");
    expect(getCellContent(filled, sheetId, "A4")).toBe("Rejected");
    expect(getCellFormat(filled, sheetId, "A3")).toEqual({ bold: true, backgroundColor: "#eaf7f2" });
    expect(getCellFormat(filled, sheetId, "A4")).toEqual({ italic: true, textColor: "#d44949" });
    expect(getCellValidation(filled, sheetId, "A3")).toEqual({ type: "list", values: ["Approved", "Rejected"] });
    expect(getCellValidation(filled, sheetId, "A4")).toBeNull();
  });

  it("auto-fills month and weekday name series", () => {
    let workbook = createBlankWorkbook();
    const sheetId = workbook.activeSheetId;
    workbook = setCellContent(workbook, sheetId, "A1", "Jan");
    workbook = setCellContent(workbook, sheetId, "A2", "Feb");
    workbook = setCellContent(workbook, sheetId, "B1", "Monday");

    const months = autoFillRange(workbook, sheetId, range("A1", "A2"), range("A1", "A5"));
    expect(getCellContent(months, sheetId, "A3")).toBe("Mar");
    expect(getCellContent(months, sheetId, "A4")).toBe("Apr");
    expect(getCellContent(months, sheetId, "A5")).toBe("May");

    const weekdays = autoFillRange(workbook, sheetId, range("B1"), range("B1", "B4"));
    expect(getCellContent(weekdays, sheetId, "B2")).toBe("Tuesday");
    expect(getCellContent(weekdays, sheetId, "B3")).toBe("Wednesday");
    expect(getCellContent(weekdays, sheetId, "B4")).toBe("Thursday");
  });

  it("auto-fills numeric series and formulas upward from the selected range", () => {
    let workbook = createBlankWorkbook();
    const sheetId = workbook.activeSheetId;
    workbook = setCellContent(workbook, sheetId, "A4", "10");
    workbook = setCellContent(workbook, sheetId, "A5", "12");
    workbook = setCellContent(workbook, sheetId, "B4", "=A4");

    const series = autoFillRange(workbook, sheetId, range("A4", "A5"), range("A1", "A5"));
    expect(getCellContent(series, sheetId, "A3")).toBe(8);
    expect(getCellContent(series, sheetId, "A2")).toBe(6);
    expect(getCellContent(series, sheetId, "A1")).toBe(4);

    const formulas = autoFillRange(workbook, sheetId, range("B4"), range("B2", "B4"));
    expect(getCellContent(formulas, sheetId, "B3")).toBe("=A3");
    expect(getCellContent(formulas, sheetId, "B2")).toBe("=A2");
  });

  it("auto-fills numeric series leftward from the selected range", () => {
    let workbook = createBlankWorkbook();
    const sheetId = workbook.activeSheetId;
    workbook = setCellContent(workbook, sheetId, "D1", 20);
    workbook = setCellContent(workbook, sheetId, "E1", 25);

    const filled = autoFillRange(workbook, sheetId, range("D1", "E1"), range("A1", "E1"));
    expect(getCellContent(filled, sheetId, "C1")).toBe(15);
    expect(getCellContent(filled, sheetId, "B1")).toBe(10);
    expect(getCellContent(filled, sheetId, "A1")).toBe(5);
  });

  it("auto-fills month names backwards when dragging up", () => {
    let workbook = createBlankWorkbook();
    const sheetId = workbook.activeSheetId;
    workbook = setCellContent(workbook, sheetId, "A3", "Mar");
    workbook = setCellContent(workbook, sheetId, "A4", "Apr");

    const months = autoFillRange(workbook, sheetId, range("A3", "A4"), range("A1", "A4"));
    expect(getCellContent(months, sheetId, "A2")).toBe("Feb");
    expect(getCellContent(months, sheetId, "A1")).toBe("Jan");
  });

  it("auto-fills text-with-number series like Excel", () => {
    let workbook = createBlankWorkbook();
    const sheetId = workbook.activeSheetId;
    workbook = setCellContent(workbook, sheetId, "A1", "Item 1");
    workbook = setCellContent(workbook, sheetId, "A2", "Item 2");
    workbook = setCellContent(workbook, sheetId, "B1", "Q3");
    workbook = setCellContent(workbook, sheetId, "C1", "Task 05");

    const stepped = autoFillRange(workbook, sheetId, range("A1", "A2"), range("A1", "A4"));
    expect(getCellContent(stepped, sheetId, "A3")).toBe("Item 3");
    expect(getCellContent(stepped, sheetId, "A4")).toBe("Item 4");

    const single = autoFillRange(workbook, sheetId, range("B1"), range("B1", "B3"));
    expect(getCellContent(single, sheetId, "B2")).toBe("Q4");
    expect(getCellContent(single, sheetId, "B3")).toBe("Q5");

    const padded = autoFillRange(workbook, sheetId, range("C1"), range("C1", "C2"));
    expect(getCellContent(padded, sheetId, "C2")).toBe("Task 06");
  });

  it("translates formulas ending in digits instead of treating them as text series", () => {
    let workbook = createBlankWorkbook();
    const sheetId = workbook.activeSheetId;
    workbook = setCellContent(workbook, sheetId, "A1", "=B2*2");

    const filled = autoFillRange(workbook, sheetId, range("A1"), range("A1", "A3"));
    expect(getCellContent(filled, sheetId, "A2")).toBe("=B3*2");
    expect(getCellContent(filled, sheetId, "A3")).toBe("=B4*2");
  });

  it("copies numeric-looking single cells instead of inventing a text series", () => {
    let workbook = createBlankWorkbook();
    const sheetId = workbook.activeSheetId;
    workbook = setCellContent(workbook, sheetId, "A1", "1.5");
    workbook = setCellContent(workbook, sheetId, "B1", "+3");

    const decimal = autoFillRange(workbook, sheetId, range("A1"), range("A1", "A3"));
    expect(getCellContent(decimal, sheetId, "A2")).toBe("1.5");
    expect(getCellContent(decimal, sheetId, "A3")).toBe("1.5");

    const signed = autoFillRange(workbook, sheetId, range("B1"), range("B1", "B2"));
    expect(getCellContent(signed, sheetId, "B2")).toBe("+3");
  });

  it("keeps zero-padding stable when a padded text series crosses zero", () => {
    let workbook = createBlankWorkbook();
    const sheetId = workbook.activeSheetId;
    workbook = setCellContent(workbook, sheetId, "A3", "Task 01");

    const filled = autoFillRange(workbook, sheetId, range("A3"), range("A1", "A3"));
    expect(getCellContent(filled, sheetId, "A2")).toBe("Task 00");
    expect(getCellContent(filled, sheetId, "A1")).toBe("Task -01");
  });

  it("adds, renames, duplicates, switches, and deletes sheets", () => {
    let workbook = createBlankWorkbook();
    workbook = addSheet(workbook, "Budget");
    expect(workbook.sheets.map((sheet) => sheet.name)).toEqual(["Sheet1", "Budget"]);

    const budget = workbook.sheets[1];
    workbook = setActiveSheet(workbook, budget.id);
    workbook = renameSheet(workbook, budget.id, "Forecast");
    workbook = setCellContent(workbook, budget.id, "A1", "100");
    workbook = duplicateSheet(workbook, budget.id);

    expect(getActiveSheet(workbook).name).toBe("Forecast Copy");
    expect(getCellContent(workbook, workbook.activeSheetId, "A1")).toBe("100");

    workbook = deleteSheet(workbook, budget.id);
    expect(workbook.sheets.map((sheet) => sheet.name)).toEqual(["Sheet1", "Forecast Copy"]);
  });

  it("moves sheets while preserving the active sheet and sheet contents", () => {
    let workbook = createBlankWorkbook();
    const firstSheetId = workbook.activeSheetId;
    workbook = setCellContent(workbook, firstSheetId, "A1", "First");
    workbook = addSheet(workbook, "Budget");
    const budgetSheetId = workbook.activeSheetId;
    workbook = setCellContent(workbook, budgetSheetId, "A1", "Budget");
    workbook = addSheet(workbook, "Archive");
    const archiveSheetId = workbook.activeSheetId;
    workbook = setCellContent(workbook, archiveSheetId, "A1", "Archive");

    workbook = moveSheet(workbook, archiveSheetId, 0);

    expect(workbook.sheets.map((sheet) => sheet.name)).toEqual(["Archive", "Sheet1", "Budget"]);
    expect(workbook.activeSheetId).toBe(archiveSheetId);
    expect(getCellContent(workbook, archiveSheetId, "A1")).toBe("Archive");
    expect(getCellContent(workbook, budgetSheetId, "A1")).toBe("Budget");

    workbook = moveSheet(workbook, archiveSheetId, 99);

    expect(workbook.sheets.map((sheet) => sheet.name)).toEqual(["Sheet1", "Budget", "Archive"]);
    expect(moveSheet(workbook, archiveSheetId, 2)).toBe(workbook);
  });

  it("does not create new workbook objects for unchanged sheet actions", () => {
    const workbook = createBlankWorkbook();

    expect(setActiveSheet(workbook, workbook.activeSheetId)).toBe(workbook);
    expect(renameSheet(workbook, workbook.activeSheetId, "Sheet1")).toBe(workbook);
  });

  it("does not delete the last remaining sheet", () => {
    const workbook = createBlankWorkbook();
    expect(deleteSheet(workbook, workbook.activeSheetId)).toEqual(workbook);
  });

  it("tracks undo and redo history", () => {
    const first = createBlankWorkbook();
    const second = setCellContent(first, first.activeSheetId, "A1", "10");
    const third = setCellContent(second, second.activeSheetId, "A1", "20");

    let history = createHistory(first);
    history = commitHistory(history, second);
    history = commitHistory(history, third);

    history = undoHistory(history);
    expect(getCellContent(history.present, history.present.activeSheetId, "A1")).toBe("10");

    history = redoHistory(history);
    expect(getCellContent(history.present, history.present.activeSheetId, "A1")).toBe("20");
  });

  it("never retains more than 100 undo snapshots", () => {
    const initial = createBlankWorkbook();
    const sheetId = initial.activeSheetId;
    let history = createHistory(initial);
    let maximumDepth = 0;

    for (let value = 1; value <= 105; value += 1) {
      history = commitHistory(history, setCellContent(history.present, sheetId, "A1", value));
      maximumDepth = Math.max(maximumDepth, history.past.length);
    }

    expect({
      maximumDepth,
      retainedDepth: history.past.length,
      oldestRetainedValue: getCellContent(history.past[0], sheetId, "A1"),
      newestRetainedValue: getCellContent(history.past.at(-1)!, sheetId, "A1")
    }).toEqual({
      maximumDepth: 100,
      retainedDepth: 100,
      oldestRetainedValue: 5,
      newestRetainedValue: 104
    });
  });

  it("applies text and color formatting to a range", () => {
    const workbook = createBlankWorkbook();
    const formatted = setCellFormat(workbook, workbook.activeSheetId, range("A1", "B2"), {
      bold: true,
      italic: true,
      textColor: "#ffffff",
      backgroundColor: "#1f6feb",
      verticalAlign: "bottom"
    });

    expect(getCellFormat(formatted, formatted.activeSheetId, "A1")).toEqual({
      bold: true,
      italic: true,
      textColor: "#ffffff",
      backgroundColor: "#1f6feb",
      verticalAlign: "bottom"
    });
    expect(getCellFormat(formatted, formatted.activeSheetId, "B2").bold).toBe(true);
    expect(getCellFormat(workbook, workbook.activeSheetId, "A1")).toEqual({});
  });

  it("clears cell formats without clearing content or metadata", () => {
    let workbook = createBlankWorkbook();
    const sheetId = workbook.activeSheetId;

    workbook = setCellContent(workbook, sheetId, "A1", "Keep value");
    workbook = setCellFormat(workbook, sheetId, range("A1"), {
      bold: true,
      backgroundColor: "#eaf7f2",
      numberFormat: "currency",
      wrapText: true
    });
    workbook = setCellValidation(workbook, sheetId, range("A1"), { type: "list", values: ["Keep value"] });
    workbook = setCellComment(workbook, sheetId, "A1", "Keep note");
    workbook = setCellHyperlink(workbook, sheetId, "A1", "https://example.com/report");
    workbook = addConditionalFormatRule(workbook, sheetId, range("A1"), {
      condition: { type: "textContains", value: "Keep" },
      format: { backgroundColor: "#fff1d6", textColor: "#8a4b00" }
    });

    const cleared = clearCellFormats(workbook, sheetId, range("A1"));

    expect(getCellContent(cleared, sheetId, "A1")).toBe("Keep value");
    expect(getCellFormat(cleared, sheetId, "A1")).toEqual({});
    expect(getCellValidation(cleared, sheetId, "A1")).toEqual({ type: "list", values: ["Keep value"] });
    expect(getCellComment(cleared, sheetId, "A1")).toBe("Keep note");
    expect(getCellHyperlink(cleared, sheetId, "A1")).toBe("https://example.com/report");
    expect(getCellConditionalFormatRules(cleared, sheetId, "A1")).toHaveLength(0);
    expect(clearCellFormats(cleared, sheetId, range("A1"))).toBe(cleared);
  });

  it("applies and clears cell border formatting without removing other formats", () => {
    let workbook = createBlankWorkbook();
    const sheetId = workbook.activeSheetId;

    workbook = setCellFormat(workbook, sheetId, range("A1", "B2"), { backgroundColor: "#eaf7f2" });
    const bordered = setCellBorders(workbook, sheetId, range("A1", "B2"), "all");

    expect(getCellFormat(bordered, sheetId, "A1").borders).toEqual({
      top: { style: "thin", color: "#64748b" },
      right: { style: "thin", color: "#64748b" },
      bottom: { style: "thin", color: "#64748b" },
      left: { style: "thin", color: "#64748b" }
    });
    expect(getCellFormat(bordered, sheetId, "B2").borders?.bottom).toEqual({ style: "thin", color: "#64748b" });

    const cleared = setCellBorders(bordered, sheetId, range("A1", "B2"), "none");
    expect(getCellFormat(cleared, sheetId, "A1").borders).toBeUndefined();
    expect(getCellFormat(cleared, sheetId, "A1").backgroundColor).toBe("#eaf7f2");
    expect(setCellBorders(cleared, sheetId, range("A1", "B2"), "none")).toBe(cleared);
  });

  it("applies outer borders only to the outside edge of a selected range", () => {
    const workbook = createBlankWorkbook();
    const sheetId = workbook.activeSheetId;

    const bordered = setCellBorders(workbook, sheetId, range("A1", "B2"), "outer", "#334155");

    expect(getCellFormat(bordered, sheetId, "A1").borders).toEqual({
      top: { style: "thin", color: "#334155" },
      left: { style: "thin", color: "#334155" }
    });
    expect(getCellFormat(bordered, sheetId, "B1").borders).toEqual({
      top: { style: "thin", color: "#334155" },
      right: { style: "thin", color: "#334155" }
    });
    expect(getCellFormat(bordered, sheetId, "A2").borders).toEqual({
      bottom: { style: "thin", color: "#334155" },
      left: { style: "thin", color: "#334155" }
    });
    expect(getCellFormat(bordered, sheetId, "B2").borders).toEqual({
      right: { style: "thin", color: "#334155" },
      bottom: { style: "thin", color: "#334155" }
    });
  });

  it("locks, unlocks, duplicates, and shifts read-only cells and sheet protection", () => {
    let workbook = createBlankWorkbook();
    const sheetId = workbook.activeSheetId;

    workbook = setRangeReadOnly(workbook, sheetId, range("B2"), true);
    expect(getCellReadOnly(workbook, sheetId, "B2")).toBe(true);
    expect(getCellReadOnly(workbook, sheetId, "A1")).toBe(false);

    workbook = setSheetProtection(workbook, sheetId, true);
    expect(getCellReadOnly(workbook, sheetId, "A1")).toBe(true);

    workbook = setRangeReadOnly(workbook, sheetId, range("A1"), false);
    expect(getCellReadOnly(workbook, sheetId, "A1")).toBe(false);
    expect(getCellReadOnly(workbook, sheetId, "B2")).toBe(true);

    const shifted = insertRows(insertColumns(workbook, sheetId, 1), sheetId, 1);
    expect(getCellReadOnly(shifted, sheetId, "A1")).toBe(false);
    expect(getCellReadOnly(shifted, sheetId, "C3")).toBe(true);

    const duplicated = duplicateSheet(shifted, sheetId);
    expect(getCellReadOnly(duplicated, duplicated.activeSheetId, "A1")).toBe(false);
    expect(getCellReadOnly(duplicated, duplicated.activeSheetId, "C3")).toBe(true);

    const unprotected = setSheetProtection(shifted, sheetId, false);
    expect(getCellReadOnly(unprotected, sheetId, "A1")).toBe(false);
    expect(getCellReadOnly(unprotected, sheetId, "C3")).toBe(true);
    expect(setRangeReadOnly(unprotected, sheetId, range("C3"), false)).not.toBe(unprotected);
  });

  it("sets, duplicates, and clears sheet freeze panes", () => {
    let workbook = createBlankWorkbook();
    const sheetId = workbook.activeSheetId;

    workbook = setSheetFreezePanes(workbook, sheetId, { freezeTopRow: true, freezeFirstColumn: true });
    expect(getActiveSheet(workbook)).toMatchObject({ freezeTopRow: true, freezeFirstColumn: true });

    const duplicated = duplicateSheet(workbook, sheetId);
    expect(getActiveSheet(duplicated)).toMatchObject({ freezeTopRow: true, freezeFirstColumn: true });

    const cleared = setSheetFreezePanes(workbook, sheetId, { freezeTopRow: false, freezeFirstColumn: false });
    expect(getActiveSheet(cleared)).toMatchObject({ freezeTopRow: false, freezeFirstColumn: false });
    expect(setSheetFreezePanes(cleared, sheetId, { freezeTopRow: false, freezeFirstColumn: false })).toBe(cleared);
  });

  it("adds, clears, duplicates, and shifts cell comments", () => {
    let workbook = createBlankWorkbook();
    const sheetId = workbook.activeSheetId;
    workbook = setCellComment(workbook, sheetId, "B2", "Check the forecast input");

    expect(getCellComment(workbook, sheetId, "B2")).toBe("Check the forecast input");

    const inserted = insertRows(insertColumns(workbook, sheetId, 1), sheetId, 1);
    expect(getCellComment(inserted, sheetId, "C3")).toBe("Check the forecast input");
    expect(getCellComment(inserted, sheetId, "B2")).toBeNull();

    const duplicated = duplicateSheet(inserted, sheetId);
    expect(getCellComment(duplicated, duplicated.activeSheetId, "C3")).toBe("Check the forecast input");

    const cleared = setCellComment(inserted, sheetId, "C3", "");
    expect(getCellComment(cleared, sheetId, "C3")).toBeNull();
  });

  it("clears comments from a range without clearing content or metadata", () => {
    let workbook = createBlankWorkbook();
    const sheetId = workbook.activeSheetId;

    workbook = setCellContent(workbook, sheetId, "A1", "Forecast");
    workbook = setCellFormat(workbook, sheetId, range("A1"), { bold: true });
    workbook = setCellValidation(workbook, sheetId, range("A1"), { type: "list", values: ["Forecast"] });
    workbook = setCellHyperlink(workbook, sheetId, "A1", "https://example.com/forecast");
    workbook = setCellComment(workbook, sheetId, "A1", "Review the forecast");
    workbook = setCellComment(workbook, sheetId, "B1", "Keep nearby note");

    const cleared = clearCellComments(workbook, sheetId, range("A1"));

    expect(getCellContent(cleared, sheetId, "A1")).toBe("Forecast");
    expect(getCellFormat(cleared, sheetId, "A1")).toEqual({ bold: true });
    expect(getCellValidation(cleared, sheetId, "A1")).toEqual({ type: "list", values: ["Forecast"] });
    expect(getCellHyperlink(cleared, sheetId, "A1")).toBe("https://example.com/forecast");
    expect(getCellComment(cleared, sheetId, "A1")).toBeNull();
    expect(getCellComment(cleared, sheetId, "B1")).toBe("Keep nearby note");
    expect(clearCellComments(cleared, sheetId, range("A1"))).toBe(cleared);
  });

  it("merges, unmerges, duplicates, and shifts merged cells", () => {
    let workbook = createBlankWorkbook();
    const sheetId = workbook.activeSheetId;
    workbook = setCellContent(workbook, sheetId, "A1", "Quarterly report");
    workbook = setCellContent(workbook, sheetId, "B1", "hidden");
    workbook = setCellContent(workbook, sheetId, "A2", "also hidden");

    workbook = mergeCells(workbook, sheetId, range("A1", "B2"));

    expect(getCellMerge(workbook, sheetId, "A1")).toMatchObject({
      role: "anchor",
      range: range("A1", "B2"),
      rowSpan: 2,
      columnSpan: 2
    });
    expect(getCellMerge(workbook, sheetId, "B1")).toMatchObject({
      role: "covered",
      anchor: { row: 0, column: 0 },
      range: range("A1", "B2")
    });
    expect(getCellContent(workbook, sheetId, "A1")).toBe("Quarterly report");
    expect(getCellContent(workbook, sheetId, "B1")).toBeNull();
    expect(getCellContent(workbook, sheetId, "A2")).toBeNull();

    const shifted = insertRows(insertColumns(workbook, sheetId, 0), sheetId, 0);
    expect(getCellMerge(shifted, sheetId, "B2")).toMatchObject({
      role: "anchor",
      range: range("B2", "C3")
    });

    const duplicated = duplicateSheet(shifted, sheetId);
    expect(getCellMerge(duplicated, duplicated.activeSheetId, "B2")).toMatchObject({
      role: "anchor",
      range: range("B2", "C3")
    });

    const unmerged = unmergeCells(shifted, sheetId, range("B2", "C3"));
    expect(getCellMerge(unmerged, sheetId, "B2")).toBeNull();
  });

  it("defines, updates, and shifts named ranges", () => {
    let workbook = createBlankWorkbook();
    const sheetId = workbook.activeSheetId;
    workbook = defineNamedRange(workbook, sheetId, " Sales_Total ", range("A1", "A3"));

    expect(getNamedRange(workbook, "sales_total")).toMatchObject({
      name: "Sales_Total",
      sheetId,
      range: range("A1", "A3")
    });
    expect(getNamedRangeForSelection(workbook, sheetId, range("A1", "A3"))?.name).toBe("Sales_Total");

    workbook = defineNamedRange(workbook, sheetId, "Sales_Total", range("B1", "B3"));
    expect(workbook.namedRanges).toHaveLength(1);
    expect(getNamedRange(workbook, "SALES_TOTAL")?.range).toEqual(range("B1", "B3"));

    const shifted = insertRows(insertColumns(workbook, sheetId, 1), sheetId, 0);
    expect(getNamedRange(shifted, "Sales_Total")?.range).toEqual(range("C2", "C4"));
  });

  it("removes named ranges by name", () => {
    let workbook = createBlankWorkbook();
    const sheetId = workbook.activeSheetId;
    workbook = defineNamedRange(workbook, sheetId, "Sales_Total", range("A1", "A3"));

    const unchanged = removeNamedRange(workbook, "Missing");
    expect(unchanged).toBe(workbook);

    const removed = removeNamedRange(workbook, "sales_total");
    expect(removed.namedRanges).toEqual([]);
    expect(getNamedRange(removed, "Sales_Total")).toBeNull();
  });

  it("applies, clears, and shifts sheet filters", () => {
    let workbook = createBlankWorkbook();
    const sheetId = workbook.activeSheetId;
    workbook = addSheetFilter(workbook, sheetId, range("A1", "B4"), {
      column: 0,
      operator: "equals",
      value: "West",
      hasHeader: true
    });

    expect(getSheetFilters(workbook, sheetId)).toHaveLength(1);
    expect(getSheetFilters(workbook, sheetId)[0]).toMatchObject({
      range: range("A1", "B4"),
      column: 0,
      operator: "equals",
      value: "West",
      hasHeader: true
    });
    expect(getSheetAutoFilterRange(workbook, sheetId)).toEqual(range("A1", "B4"));

    const shifted = insertRows(workbook, sheetId, 0);
    expect(getSheetFilters(shifted, sheetId)[0]).toMatchObject({
      range: range("A2", "B5"),
      column: 0
    });
    expect(getSheetAutoFilterRange(shifted, sheetId)).toEqual(range("A2", "B5"));

    const cleared = clearSheetFilters(shifted, sheetId);
    expect(getSheetFilters(cleared, sheetId)).toHaveLength(0);
    expect(getSheetAutoFilterRange(cleared, sheetId)).toBeNull();
  });

  it("replaces filters for the same range column", () => {
    let workbook = createBlankWorkbook();
    const sheetId = workbook.activeSheetId;
    workbook = addSheetFilter(workbook, sheetId, range("A1", "B4"), {
      column: 0,
      operator: "equals",
      value: "West",
      hasHeader: true
    });

    workbook = addSheetFilter(workbook, sheetId, range("A1", "B4"), {
      column: 0,
      operator: "equals",
      value: "East",
      hasHeader: true
    });

    expect(getSheetFilters(workbook, sheetId)).toHaveLength(1);
    expect(getSheetFilters(workbook, sheetId)[0]).toMatchObject({
      range: range("A1", "B4"),
      column: 0,
      operator: "equals",
      value: "East",
      hasHeader: true
    });
  });

  it("clears one filter column while keeping the AutoFilter range active", () => {
    let workbook = createBlankWorkbook();
    const sheetId = workbook.activeSheetId;
    workbook = addSheetFilter(workbook, sheetId, range("A1", "B4"), {
      column: 0,
      operator: "equals",
      value: "West",
      hasHeader: true
    });
    workbook = addSheetFilter(workbook, sheetId, range("A1", "B4"), {
      column: 1,
      operator: "greaterThan",
      value: "9",
      hasHeader: true
    });

    const cleared = clearSheetFilter(workbook, sheetId, range("A1", "B4"), 0);

    expect(getSheetAutoFilterRange(cleared, sheetId)).toEqual(range("A1", "B4"));
    expect(getSheetFilters(cleared, sheetId)).toHaveLength(1);
    expect(getSheetFilters(cleared, sheetId)[0]).toMatchObject({
      column: 1,
      operator: "greaterThan",
      value: "9"
    });
  });

  it("adds, deletes, and shifts embedded charts", () => {
    let workbook = createBlankWorkbook();
    const sheetId = workbook.activeSheetId;
    workbook = addSheetChart(workbook, sheetId, range("A1", "B3"), {
      anchor: { row: 0, column: 3 },
      title: "Sales overview",
      type: "bar"
    });

    const chart = getSheetCharts(workbook, sheetId)[0];
    expect(chart).toMatchObject({
      id: "chart-1",
      title: "Sales overview",
      type: "bar",
      range: range("A1", "B3"),
      anchor: { row: 0, column: 3 }
    });

    const shifted = insertRows(workbook, sheetId, 0);
    expect(getSheetCharts(shifted, sheetId)[0]).toMatchObject({
      range: range("A2", "B4"),
      anchor: { row: 1, column: 3 }
    });

    const removed = deleteSheetChart(shifted, sheetId, chart.id);
    expect(getSheetCharts(removed, sheetId)).toHaveLength(0);
  });

  it("applies, clears, and shifts conditional format rules", () => {
    let workbook = createBlankWorkbook();
    const sheetId = workbook.activeSheetId;
    workbook = addConditionalFormatRule(workbook, sheetId, range("A2", "A3"), {
      condition: { type: "greaterThan", value: "10" },
      format: { backgroundColor: "#fff1d6", textColor: "#8a4b00", bold: true }
    });

    expect(getCellConditionalFormatRules(workbook, sheetId, "A2")).toHaveLength(1);
    expect(getCellConditionalFormatRules(workbook, sheetId, "A1")).toHaveLength(0);

    const shifted = insertRows(workbook, sheetId, 1);
    expect(getCellConditionalFormatRules(shifted, sheetId, "A3")).toHaveLength(1);
    expect(getCellConditionalFormatRules(shifted, sheetId, "A2")).toHaveLength(0);

    const cleared = clearConditionalFormatRules(shifted, sheetId, range("A3", "A4"));
    expect(getCellConditionalFormatRules(cleared, sheetId, "A3")).toHaveLength(0);
    expect(getCellConditionalFormatRules(cleared, sheetId, "A4")).toHaveLength(0);
  });

  it("removes one conditional format rule without clearing the other rules", () => {
    let workbook = createBlankWorkbook();
    const sheetId = workbook.activeSheetId;
    workbook = addConditionalFormatRule(workbook, sheetId, range("A1", "A2"), {
      condition: { type: "greaterThan", value: "10" },
      format: { backgroundColor: "#fff1d6", textColor: "#8a4b00", bold: true }
    });
    workbook = addConditionalFormatRule(workbook, sheetId, range("B1", "B2"), {
      condition: { type: "textContains", value: "Risk" },
      format: { backgroundColor: "#fee2e2", textColor: "#991b1b", bold: false }
    });

    const firstRuleId = getCellConditionalFormatRules(workbook, sheetId, "A1")[0].id;
    const removed = removeConditionalFormatRule(workbook, sheetId, firstRuleId);

    expect(getCellConditionalFormatRules(removed, sheetId, "A1")).toHaveLength(0);
    expect(getCellConditionalFormatRules(removed, sheetId, "B1")).toHaveLength(1);
    expect(removeConditionalFormatRule(removed, sheetId, firstRuleId)).toBe(removed);
  });

  it("applies, clears, and shifts data validation rules", () => {
    let workbook = createBlankWorkbook();
    const sheetId = workbook.activeSheetId;
    workbook = setCellValidation(workbook, sheetId, range("A2", "A3"), {
      type: "list",
      values: ["Open", "Closed"]
    });

    expect(getCellValidation(workbook, sheetId, "A2")).toEqual({ type: "list", values: ["Open", "Closed"] });
    expect(getCellValidation(workbook, sheetId, "A3")).toEqual({ type: "list", values: ["Open", "Closed"] });

    const shifted = insertRows(workbook, sheetId, 1);
    expect(getCellValidation(shifted, sheetId, "A3")).toEqual({ type: "list", values: ["Open", "Closed"] });

    const cleared = setCellValidation(shifted, sheetId, range("A3"), null);
    expect(getCellValidation(cleared, sheetId, "A3")).toBeNull();
    expect(getCellValidation(cleared, sheetId, "A4")).toEqual({ type: "list", values: ["Open", "Closed"] });
  });

  it("does not create new workbook objects for unchanged formatting", () => {
    const workbook = createBlankWorkbook();
    const formatted = setCellFormat(workbook, workbook.activeSheetId, range("A1"), { bold: true });

    expect(setCellFormat(formatted, formatted.activeSheetId, range("A1"), { bold: true })).toBe(formatted);
  });
});
