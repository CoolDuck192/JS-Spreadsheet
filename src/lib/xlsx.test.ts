import { describe, expect, it } from "vitest";
import { strFromU8, unzipSync } from "fflate";
import type { WorkbookModel } from "../types";
import {
  addSheet,
  addConditionalFormatRule,
  createBlankWorkbook,
  defineNamedRange,
  getCellComment,
  getCellConditionalFormatRules,
  getCellContent,
  getCellFormat,
  getCellHyperlink,
  getCellReadOnly,
  getNamedRange,
  getCellValidation,
  getCellMerge,
  getColumnWidth,
  getRowHeight,
  getSheetAutoFilterRange,
  getSheetTabColor,
  isColumnHidden,
  isRowHidden,
  isSheetHidden,
  mergeCells,
  setColumnsHidden,
  setCellComment,
  setCellContent,
  setCellFormat,
  setCellHyperlink,
  setCellValidation,
  addSheetFilter,
  setColumnWidth,
  setRangeReadOnly,
  setSheetFreezePanes,
  setSheetHidden,
  setSheetProtection,
  setSheetTabColor,
  setRowsHidden,
  setRowHeight,
  setCellBorders
} from "./workbook";
import { exportStructuredTableToXlsx, exportWorkbookToXlsx, importWorkbookFromXlsx } from "./xlsx";

describe("xlsx", () => {
  it("round-trips sheets, text, formulas, comments, hyperlinks, merges, and dimensions", async () => {
    let workbook = createBlankWorkbook();
    const firstSheetId = workbook.activeSheetId;
    workbook = setCellContent(workbook, firstSheetId, "A1", "Project");
    workbook = setCellContent(workbook, firstSheetId, "D4", "=SUM(B2:B2)");
    workbook = setCellFormat(
      workbook,
      firstSheetId,
      {
        start: { row: 0, column: 0 },
        end: { row: 0, column: 0 }
      },
      { wrapText: true }
    );
    workbook = setCellComment(workbook, firstSheetId, "A1", "Imported note");
    workbook = setCellHyperlink(workbook, firstSheetId, "A1", "https://example.com/report");
    workbook = mergeCells(workbook, firstSheetId, {
      start: { row: 0, column: 4 },
      end: { row: 1, column: 5 }
    });
    workbook = setColumnWidth(workbook, firstSheetId, 1, 136);
    workbook = setRowHeight(workbook, firstSheetId, 2, 44);
    workbook = setColumnsHidden(workbook, firstSheetId, 3, 3, true);
    workbook = setRowsHidden(workbook, firstSheetId, 4, 4, true);
    workbook = addSheet(workbook, "Forecast");
    workbook = setCellContent(workbook, workbook.activeSheetId, "A1", "Second sheet");

    const data = await exportWorkbookToXlsx(workbook);
    const imported = await importWorkbookFromXlsx(data);
    const importedFirstSheetId = imported.sheets[0].id;
    const importedSecondSheetId = imported.sheets[1].id;

    expect(imported.sheets.map((sheet) => sheet.name)).toEqual(["Sheet1", "Forecast"]);
    expect(getCellContent(imported, importedFirstSheetId, "A1")).toBe("Project");
    expect(getCellContent(imported, importedFirstSheetId, "D4")).toBe("=SUM(B2:B2)");
    expect(getCellFormat(imported, importedFirstSheetId, "A1")).toMatchObject({ wrapText: true });
    expect(getCellComment(imported, importedFirstSheetId, "A1")).toBe("Imported note");
    expect(getCellHyperlink(imported, importedFirstSheetId, "A1")).toBe("https://example.com/report");
    expect(getCellMerge(imported, importedFirstSheetId, "E1")).toMatchObject({
      role: "anchor",
      rowSpan: 2,
      columnSpan: 2
    });
    expect(getColumnWidth(imported, importedFirstSheetId, 1)).toBe(136);
    expect(getRowHeight(imported, importedFirstSheetId, 2)).toBe(44);
    expect(isColumnHidden(imported, importedFirstSheetId, 3)).toBe(true);
    expect(isRowHidden(imported, importedFirstSheetId, 4)).toBe(true);
    expect(getCellContent(imported, importedSecondSheetId, "A1")).toBe("Second sheet");
  });

  it("preserves native model number and boolean types through XLSX", async () => {
    let workbook = createBlankWorkbook();
    const sheetId = workbook.activeSheetId;
    workbook = setCellContent(workbook, sheetId, "A1", 42.5);
    workbook = setCellContent(workbook, sheetId, "A2", -7);
    workbook = setCellContent(workbook, sheetId, "B1", true);
    workbook = setCellContent(workbook, sheetId, "B2", false);

    const data = await exportWorkbookToXlsx(workbook);
    const imported = await importWorkbookFromXlsx(data);
    const importedSheetId = imported.sheets[0].id;

    expect(["A1", "A2", "B1", "B2"].map((address) => getCellContent(imported, importedSheetId, address))).toEqual([
      42.5,
      -7,
      true,
      false
    ]);
  });

  it("imports and re-exports native XLSX dates as typed serials with matching formats", async () => {
    const ExcelJS = (await import("exceljs")).default;
    const nativeWorkbook = new ExcelJS.Workbook();
    const nativeSheet = nativeWorkbook.addWorksheet("Dates");
    nativeSheet.getCell("A1").value = new Date(Date.UTC(2026, 0, 15));
    nativeSheet.getCell("A1").numFmt = "mm-dd-yy";
    nativeSheet.getCell("A2").value = new Date(Date.UTC(2026, 0, 15, 12));
    nativeSheet.getCell("A2").numFmt = "m/d/yy h:mm";

    const nativeData = await nativeWorkbook.xlsx.writeBuffer();
    const imported = await importWorkbookFromXlsx(nativeData as ArrayBuffer);
    const importedSheetId = imported.sheets[0].id;

    expect(getCellContent(imported, importedSheetId, "A1")).toBe(46037);
    expect(getCellContent(imported, importedSheetId, "A2")).toBe(46037.5);
    expect(getCellFormat(imported, importedSheetId, "A1")).toMatchObject({ numberFormat: "date" });
    expect(getCellFormat(imported, importedSheetId, "A2")).toMatchObject({ numberFormat: "dateTime" });

    const exportedData = await exportWorkbookToXlsx(imported);
    const nativeReadback = new ExcelJS.Workbook();
    await nativeReadback.xlsx.load(exportedData as Parameters<typeof nativeReadback.xlsx.load>[0]);
    const nativeReadbackSheet = nativeReadback.getWorksheet("Dates");
    const nativeDate = nativeReadbackSheet?.getCell("A1");
    const nativeDateTime = nativeReadbackSheet?.getCell("A2");

    expect(nativeDate?.value).toBeInstanceOf(Date);
    expect((nativeDate?.value as Date).toISOString()).toBe("2026-01-15T00:00:00.000Z");
    expect(nativeDate?.numFmt).toBe("mmm d, yyyy");
    expect(nativeDateTime?.value).toBeInstanceOf(Date);
    expect((nativeDateTime?.value as Date).toISOString()).toBe("2026-01-15T12:00:00.000Z");
    expect(nativeDateTime?.numFmt).toBe("mmm d, yyyy h:mm AM/PM");

    const reimported = await importWorkbookFromXlsx(exportedData);
    const reimportedSheetId = reimported.sheets[0].id;

    expect(getCellContent(reimported, reimportedSheetId, "A1")).toBe(46037);
    expect(getCellContent(reimported, reimportedSheetId, "A2")).toBe(46037.5);
    expect(getCellFormat(reimported, reimportedSheetId, "A1")).toMatchObject({ numberFormat: "date" });
    expect(getCellFormat(reimported, reimportedSheetId, "A2")).toMatchObject({ numberFormat: "dateTime" });
  });

  it("classifies temporal number formats from actual tokens instead of literals and colors", async () => {
    const ExcelJS = (await import("exceljs")).default;
    const nativeWorkbook = new ExcelJS.Workbook();
    const nativeSheet = nativeWorkbook.addWorksheet("Formats");
    const formats = {
      A1: "[Red]0.00",
      A2: "h:mm",
      A3: '0.00 "days"',
      A4: "0.00\\d",
      A5: "[$USD-409]#,##0.00"
    } as const;

    for (const [address, numFmt] of Object.entries(formats)) {
      nativeSheet.getCell(address).value = address === "A2" ? 0.5 : 12.5;
      nativeSheet.getCell(address).numFmt = numFmt;
    }

    const nativeData = await nativeWorkbook.xlsx.writeBuffer();
    const imported = await importWorkbookFromXlsx(nativeData as ArrayBuffer);
    const importedSheetId = imported.sheets[0].id;

    expect(getCellContent(imported, importedSheetId, "A1")).toBe(12.5);
    expect(getCellFormat(imported, importedSheetId, "A1").numberFormat).toBeUndefined();
    expect(getCellContent(imported, importedSheetId, "A2")).toBe(0.5);
    expect(getCellFormat(imported, importedSheetId, "A2").numberFormat).toBe("dateTime");
    expect(["A3", "A4", "A5"].map((address) => getCellFormat(imported, importedSheetId, address).numberFormat)).toEqual([
      undefined,
      undefined,
      undefined
    ]);
  });

  it("round-trips the active sheet through XLSX", async () => {
    let workbook = createBlankWorkbook();
    workbook = setCellContent(workbook, workbook.activeSheetId, "A1", "First sheet");
    workbook = addSheet(workbook, "Forecast");
    workbook = setCellContent(workbook, workbook.activeSheetId, "A1", "Selected sheet");

    const data = await exportWorkbookToXlsx(workbook);
    const imported = await importWorkbookFromXlsx(data);

    expect(imported.sheets.map((sheet) => sheet.name)).toEqual(["Sheet1", "Forecast"]);
    expect(imported.activeSheetId).toBe(imported.sheets[1].id);
    expect(getCellContent(imported, imported.activeSheetId, "A1")).toBe("Selected sheet");
  });

  it("round-trips font family and size, ignoring Excel's defaults", async () => {
    let workbook = createBlankWorkbook();
    const sheetId = workbook.activeSheetId;
    workbook = setCellContent(workbook, sheetId, "A1", "Styled");
    workbook = setCellContent(workbook, sheetId, "A2", "Plain");
    workbook = setCellFormat(
      workbook,
      sheetId,
      { start: { row: 0, column: 0 }, end: { row: 0, column: 0 } },
      { fontFamily: "Georgia", fontSize: 18 }
    );

    const data = await exportWorkbookToXlsx(workbook);
    const imported = await importWorkbookFromXlsx(data);
    const importedSheetId = imported.sheets[0].id;

    expect(getCellFormat(imported, importedSheetId, "A1")).toMatchObject({ fontFamily: "Georgia", fontSize: 18 });
    // Unstyled cells must not pick up Excel's default font as an explicit format.
    const plainFormat = getCellFormat(imported, importedSheetId, "A2");
    expect(plainFormat.fontFamily).toBeUndefined();
    expect(plainFormat.fontSize).toBeUndefined();
  });

  it("keeps deliberate fonts that partially match Excel's default stamp", async () => {
    let workbook = createBlankWorkbook();
    const sheetId = workbook.activeSheetId;
    workbook = setCellContent(workbook, sheetId, "A1", "calibri large");
    workbook = setCellContent(workbook, sheetId, "A2", "arial default size");
    // Only the FULL default signature (Calibri/Aptos AND 11) is the stamp;
    // a deviation in either field marks a deliberate choice that must survive.
    workbook = setCellFormat(
      workbook,
      sheetId,
      { start: { row: 0, column: 0 }, end: { row: 0, column: 0 } },
      { fontFamily: "Calibri", fontSize: 14 }
    );
    workbook = setCellFormat(
      workbook,
      sheetId,
      { start: { row: 1, column: 0 }, end: { row: 1, column: 0 } },
      { fontFamily: "Arial", fontSize: 11 }
    );

    const data = await exportWorkbookToXlsx(workbook);
    const imported = await importWorkbookFromXlsx(data);
    const importedSheetId = imported.sheets[0].id;

    expect(getCellFormat(imported, importedSheetId, "A1")).toMatchObject({ fontFamily: "Calibri", fontSize: 14 });
    expect(getCellFormat(imported, importedSheetId, "A2")).toMatchObject({ fontFamily: "Arial", fontSize: 11 });
  });

  it("round-trips hidden sheets through XLSX", async () => {
    let workbook = createBlankWorkbook();
    const visibleSheetId = workbook.activeSheetId;
    workbook = setCellContent(workbook, visibleSheetId, "A1", "Visible sheet");
    workbook = addSheet(workbook, "Archive");
    const hiddenSheetId = workbook.activeSheetId;
    workbook = setCellContent(workbook, hiddenSheetId, "A1", "Hidden sheet");
    workbook = setSheetHidden(workbook, hiddenSheetId, true);

    const data = await exportWorkbookToXlsx(workbook);
    const imported = await importWorkbookFromXlsx(data);

    expect(imported.sheets.map((sheet) => sheet.name)).toEqual(["Sheet1", "Archive"]);
    expect(imported.activeSheetId).toBe(imported.sheets[0].id);
    expect(isSheetHidden(imported, imported.sheets[1].id)).toBe(true);
    expect(getCellContent(imported, imported.sheets[1].id, "A1")).toBe("Hidden sheet");
  });

  it("round-trips sheet tab colors through XLSX", async () => {
    let workbook = createBlankWorkbook();
    const sheetId = workbook.activeSheetId;
    workbook = setSheetTabColor(workbook, sheetId, "#0f766e");
    workbook = addSheet(workbook, "Forecast");
    workbook = setSheetTabColor(workbook, workbook.activeSheetId, "#7c3aed");

    const data = await exportWorkbookToXlsx(workbook);
    const imported = await importWorkbookFromXlsx(data);

    expect(getSheetTabColor(imported, imported.sheets[0].id)).toBe("#0f766e");
    expect(getSheetTabColor(imported, imported.sheets[1].id)).toBe("#7c3aed");
  });

  it("round-trips blank sheets through XLSX", async () => {
    const data = await exportWorkbookToXlsx(createBlankWorkbook());
    const imported = await importWorkbookFromXlsx(data);

    expect(imported.sheets).toHaveLength(1);
    expect(imported.sheets[0]).toMatchObject({
      name: "Sheet1",
      rowCount: 100,
      columnCount: 26,
      cells: {}
    });
    expect(imported.activeSheetId).toBe(imported.sheets[0].id);
  });

  it("round-trips worksheet auto filter ranges through XLSX", async () => {
    let workbook = createBlankWorkbook();
    const sheetId = workbook.activeSheetId;
    workbook = setCellContent(workbook, sheetId, "A1", "Region");
    workbook = setCellContent(workbook, sheetId, "B1", "Sales");
    workbook = setCellContent(workbook, sheetId, "A2", "West");
    workbook = setCellContent(workbook, sheetId, "B2", 120);
    workbook = addSheetFilter(workbook, sheetId, {
      start: { row: 0, column: 0 },
      end: { row: 1, column: 1 }
    }, {
      column: 0,
      operator: "equals",
      value: "West",
      hasHeader: true
    });

    const data = await exportWorkbookToXlsx(workbook);
    const imported = await importWorkbookFromXlsx(data);
    const importedSheetId = imported.sheets[0].id;

    expect(getSheetAutoFilterRange(imported, importedSheetId)).toEqual({
      start: { row: 0, column: 0 },
      end: { row: 1, column: 1 }
    });
  });

  it("round-trips common cell styles through XLSX", async () => {
    let workbook = createBlankWorkbook();
    const sheetId = workbook.activeSheetId;
    workbook = setCellContent(workbook, sheetId, "A1", "Styled total");
    workbook = setCellFormat(
      workbook,
      sheetId,
      {
        start: { row: 0, column: 0 },
        end: { row: 0, column: 0 }
      },
      {
        bold: true,
        italic: true,
        textColor: "#17634a",
        backgroundColor: "#eaf7f2",
        numberFormat: "currency",
        horizontalAlign: "center",
        verticalAlign: "middle",
        wrapText: true
      }
    );
    workbook = setCellBorders(
      workbook,
      sheetId,
      {
        start: { row: 0, column: 0 },
        end: { row: 0, column: 0 }
      },
      "all",
      "#334155"
    );

    const data = await exportWorkbookToXlsx(workbook);
    const imported = await importWorkbookFromXlsx(data);
    const importedSheetId = imported.sheets[0].id;

    expect(getCellFormat(imported, importedSheetId, "A1")).toEqual({
      bold: true,
      italic: true,
      textColor: "#17634a",
      backgroundColor: "#eaf7f2",
      numberFormat: "currency",
      horizontalAlign: "center",
      verticalAlign: "middle",
      wrapText: true,
      borders: {
        top: { style: "thin", color: "#334155" },
        right: { style: "thin", color: "#334155" },
        bottom: { style: "thin", color: "#334155" },
        left: { style: "thin", color: "#334155" }
      }
    });
  });

  it("round-trips data validation rules through XLSX", async () => {
    let workbook = createBlankWorkbook();
    const sheetId = workbook.activeSheetId;
    workbook = setCellContent(workbook, sheetId, "A1", "Open");
    workbook = setCellContent(workbook, sheetId, "B1", "5");
    workbook = setCellContent(workbook, sheetId, "C1", "Code");
    workbook = setCellValidation(
      workbook,
      sheetId,
      {
        start: { row: 0, column: 0 },
        end: { row: 0, column: 0 }
      },
      { type: "list", values: ["Open", "Closed"], allowBlank: false }
    );
    workbook = setCellValidation(
      workbook,
      sheetId,
      {
        start: { row: 0, column: 1 },
        end: { row: 0, column: 1 }
      },
      { type: "number", min: 1, max: 10 }
    );
    workbook = setCellValidation(
      workbook,
      sheetId,
      {
        start: { row: 0, column: 2 },
        end: { row: 0, column: 2 }
      },
      { type: "textLength", min: 2, max: 5 } as Parameters<typeof setCellValidation>[3]
    );

    const data = await exportWorkbookToXlsx(workbook);
    const imported = await importWorkbookFromXlsx(data);
    const importedSheetId = imported.sheets[0].id;

    expect(getCellValidation(imported, importedSheetId, "A1")).toEqual({
      type: "list",
      values: ["Open", "Closed"],
      allowBlank: false
    });
    expect(getCellValidation(imported, importedSheetId, "B1")).toEqual({ type: "number", min: 1, max: 10 });
    expect(getCellValidation(imported, importedSheetId, "C1")).toEqual({ type: "textLength", min: 2, max: 5 });
  });

  it("round-trips named ranges through XLSX", async () => {
    let workbook = createBlankWorkbook();
    const sheetId = workbook.activeSheetId;
    workbook = setCellContent(workbook, sheetId, "A1", "10");
    workbook = setCellContent(workbook, sheetId, "A2", "20");
    workbook = setCellContent(workbook, sheetId, "B1", "=SUM(Sales)");
    workbook = defineNamedRange(workbook, sheetId, "Sales", {
      start: { row: 0, column: 0 },
      end: { row: 1, column: 0 }
    });

    const data = await exportWorkbookToXlsx(workbook);
    const imported = await importWorkbookFromXlsx(data);
    const importedSheetId = imported.sheets[0].id;

    expect(getNamedRange(imported, "Sales")).toMatchObject({
      name: "Sales",
      sheetId: importedSheetId,
      range: {
        start: { row: 0, column: 0 },
        end: { row: 1, column: 0 }
      }
    });
    expect(getCellContent(imported, importedSheetId, "B1")).toBe("=SUM(Sales)");
  });

  it("round-trips conditional formatting rules through XLSX", async () => {
    let workbook = createBlankWorkbook();
    const sheetId = workbook.activeSheetId;
    workbook = setCellContent(workbook, sheetId, "A1", "5");
    workbook = setCellContent(workbook, sheetId, "A2", "15");
    workbook = setCellContent(workbook, sheetId, "C1", "West");
    workbook = setCellContent(workbook, sheetId, "C2", "East");
    workbook = setCellContent(workbook, sheetId, "C3", "West");
    workbook = setCellContent(workbook, sheetId, "D1", "10");
    workbook = setCellContent(workbook, sheetId, "D2", "30");
    workbook = setCellContent(workbook, sheetId, "D3", "20");
    workbook = setCellContent(workbook, sheetId, "E1", "10");
    workbook = setCellContent(workbook, sheetId, "E2", "20");
    workbook = setCellContent(workbook, sheetId, "E3", "30");
    workbook = addConditionalFormatRule(
      workbook,
      sheetId,
      {
        start: { row: 0, column: 0 },
        end: { row: 1, column: 0 }
      },
      {
        condition: { type: "greaterThan", value: "10" },
        format: { backgroundColor: "#fff1d6", textColor: "#8a4b00", bold: true }
      }
    );
    workbook = addConditionalFormatRule(
      workbook,
      sheetId,
      {
        start: { row: 0, column: 1 },
        end: { row: 1, column: 1 }
      },
      {
        condition: { type: "blank" },
        format: { backgroundColor: "#fde8e8", textColor: "#9b1c1c" }
      }
    );
    workbook = addConditionalFormatRule(
      workbook,
      sheetId,
      {
        start: { row: 0, column: 2 },
        end: { row: 2, column: 2 }
      },
      {
        condition: { type: "duplicate" },
        format: { backgroundColor: "#fde8e8", textColor: "#9b1c1c" }
      }
    );
    workbook = addConditionalFormatRule(
      workbook,
      sheetId,
      {
        start: { row: 0, column: 3 },
        end: { row: 2, column: 3 }
      },
      {
        condition: { type: "top", count: 2 } as Parameters<typeof addConditionalFormatRule>[3]["condition"],
        format: { backgroundColor: "#fff1d6", textColor: "#8a4b00" }
      }
    );
    workbook = addConditionalFormatRule(
      workbook,
      sheetId,
      {
        start: { row: 0, column: 4 },
        end: { row: 2, column: 4 }
      },
      {
        condition: {
          type: "colorScale",
          minColor: "#ffffff",
          maxColor: "#000000"
        } as Parameters<typeof addConditionalFormatRule>[3]["condition"],
        format: {}
      }
    );

    const data = await exportWorkbookToXlsx(workbook);
    const imported = await importWorkbookFromXlsx(data);
    const importedSheetId = imported.sheets[0].id;

    expect(getCellConditionalFormatRules(imported, importedSheetId, "A2")).toMatchObject([
      {
        range: {
          start: { row: 0, column: 0 },
          end: { row: 1, column: 0 }
        },
        condition: { type: "greaterThan", value: "10" },
        format: { backgroundColor: "#fff1d6", textColor: "#8a4b00", bold: true }
      }
    ]);
    expect(getCellConditionalFormatRules(imported, importedSheetId, "B1")).toMatchObject([
      {
        range: {
          start: { row: 0, column: 1 },
          end: { row: 1, column: 1 }
        },
        condition: { type: "blank" },
        format: { backgroundColor: "#fde8e8", textColor: "#9b1c1c" }
      }
    ]);
    expect(getCellConditionalFormatRules(imported, importedSheetId, "C1")).toMatchObject([
      {
        range: {
          start: { row: 0, column: 2 },
          end: { row: 2, column: 2 }
        },
        condition: { type: "duplicate" },
        format: { backgroundColor: "#fde8e8", textColor: "#9b1c1c" }
      }
    ]);
    expect(getCellConditionalFormatRules(imported, importedSheetId, "D1")).toMatchObject([
      {
        range: {
          start: { row: 0, column: 3 },
          end: { row: 2, column: 3 }
        },
        condition: { type: "top", count: 2 },
        format: { backgroundColor: "#fff1d6", textColor: "#8a4b00" }
      }
    ]);
    expect(getCellConditionalFormatRules(imported, importedSheetId, "E1")).toMatchObject([
      {
        range: {
          start: { row: 0, column: 4 },
          end: { row: 2, column: 4 }
        },
        condition: { type: "colorScale", minColor: "#ffffff", maxColor: "#000000" },
        format: {}
      }
    ]);
  });

  it("round-trips protected sheets and unlocked cells through XLSX", async () => {
    let workbook = createBlankWorkbook();
    const sheetId = workbook.activeSheetId;
    workbook = setCellContent(workbook, sheetId, "A1", "Formula");
    workbook = setCellContent(workbook, sheetId, "B1", "Input");
    workbook = setRangeReadOnly(
      workbook,
      sheetId,
      {
        start: { row: 0, column: 1 },
        end: { row: 0, column: 1 }
      },
      false
    );
    workbook = setSheetProtection(workbook, sheetId, true);

    const data = await exportWorkbookToXlsx(workbook);
    const imported = await importWorkbookFromXlsx(data);
    const importedSheetId = imported.sheets[0].id;

    expect(getCellReadOnly(imported, importedSheetId, "A1")).toBe(true);
    expect(getCellReadOnly(imported, importedSheetId, "B1")).toBe(false);
    expect(imported.sheets[0].protection).toEqual({
      isProtected: true,
      lockedCells: {},
      unlockedCells: { B1: true }
    });
  });

  it("round-trips freeze panes through XLSX", async () => {
    let workbook = createBlankWorkbook();
    const sheetId = workbook.activeSheetId;
    workbook = setCellContent(workbook, sheetId, "A1", "Header");
    workbook = setSheetFreezePanes(workbook, sheetId, { freezeTopRow: true, freezeFirstColumn: true });

    const data = await exportWorkbookToXlsx(workbook);
    const imported = await importWorkbookFromXlsx(data);

    expect(imported.sheets[0]).toMatchObject({
      freezeTopRow: true,
      freezeFirstColumn: true
    });
  });
});

describe("native structured table XLSX", () => {
  function nativeTableWorkbook(): WorkbookModel {
    const base = createBlankWorkbook();
    const sheet = base.sheets[0];
    return {
      ...base,
      sheets: [{
        ...sheet,
        name: "Sales",
        cells: {
          A1: "Label", B1: "Standard", C1: "Custom Value", D1: "Custom Formula",
          A2: "East", B2: 2, C2: 5, D2: "=B2*C2",
          A3: "West", B3: 4, C3: 3, D3: "=B3*C3",
          A4: "Grand Total", B4: "=SUM(B2:B3)", C4: 99, D4: "=SUM(D2:D3)"
        }
      }],
      tables: [{
        id: "private-table-id",
        name: "SalesTable",
        sheetId: sheet.id,
        range: { start: { row: 0, column: 0 }, end: { row: 3, column: 3 } },
        headerRow: true,
        totalsRow: true,
        columns: [
          { id: "private-label", name: "Label", sheetColumn: 0, totalsLabel: "Grand Total" },
          { id: "private-standard", name: "Standard", sheetColumn: 1, totalsFunction: "sum" },
          { id: "private-custom-value", name: "Custom Value", sheetColumn: 2 },
          { id: "private-custom-formula", name: "Custom Formula", sheetColumn: 3, calculatedFormula: "=B2*C2" }
        ],
        rowIds: ["private-row-one", "private-row-two"],
        keyColumnId: "private-label",
        style: { theme: "TableStyleLight9", showRowStripes: true },
        filter: {
          kind: "set",
          columnId: "private-label",
          operator: "in",
          values: [{ type: "string", value: "East" }]
        }
      }]
    };
  }

  it("exports actual Excel tables plus the OOXML metadata ExcelJS drops", async () => {
    const workbook = nativeTableWorkbook();
    const data = await exportWorkbookToXlsx(workbook);
    const ExcelJS = (await import("exceljs")).default;
    const native = new ExcelJS.Workbook();
    await native.xlsx.load(data as Parameters<typeof native.xlsx.load>[0]);
    const worksheet = native.getWorksheet("Sales")!;

    expect(worksheet.getTables()).toHaveLength(1);
    expect(worksheet.getTable("SalesTable").name).toBe("SalesTable");
    expect(worksheet.getCell("B4").formula).toBe("SUM(B2:B3)");
    expect(worksheet.getCell("D4").formula).toBe("SUM(D2:D3)");
    expect(worksheet.getCell("C4").value).toBe(99);

    const entries = unzipSync(new Uint8Array(data));
    const xml = strFromU8(entries["xl/tables/table1.xml"]);
    expect(Object.keys(entries)).toContain("xl/worksheets/_rels/sheet1.xml.rels");
    expect(xml).toContain('totalsRowLabel="Grand Total"');
    expect(xml).toContain('totalsRowFunction="sum"');
    expect(xml).toContain('totalsRowFunction="custom"');
    expect(xml).toContain("<totalsRowFormula>");
    expect(xml).toContain("<calculatedColumnFormula>");
    expect(xml).toContain('<filter val="East"/>');
    expect(xml).toContain('name="TableStyleLight9"');
    for (const privateId of [
      workbook.tables[0].id,
      workbook.tables[0].sheetId,
      workbook.tables[0].keyColumnId,
      ...workbook.tables[0].columns.map((column) => column.id),
      ...workbook.tables[0].rowIds
    ]) {
      expect(xml).not.toContain(privateId);
    }
  });

  it("preserves standard, custom formula, custom value, label, calculation, and filter semantics", async () => {
    const first = await importWorkbookFromXlsx(await exportWorkbookToXlsx(nativeTableWorkbook()));
    const persisted = JSON.parse(JSON.stringify(first)) as WorkbookModel;
    const second = await importWorkbookFromXlsx(await exportWorkbookToXlsx(persisted));
    const table = second.tables[0];
    const sheet = second.sheets[0];

    expect(table.columns.find((column) => column.name === "Label")?.totalsLabel).toBe("Grand Total");
    expect(table.columns.find((column) => column.name === "Standard")?.totalsFunction).toBe("sum");
    expect(table.columns.find((column) => column.name === "Custom Formula")?.calculatedFormula).toBe("=B2*C2");
    expect(sheet.cells.D4).toBe("=SUM(D2:D3)");
    expect(sheet.cells.C4).toBe(99);
    expect(table.filter).toMatchObject({ kind: "set", operator: "in" });
  });

  it("imports table-column formula metadata even when ExcelJS cannot parse its native column position", async () => {
    const workbook = nativeTableWorkbook();
    const sheet = workbook.sheets[0];
    const adjusted: WorkbookModel = {
      ...workbook,
      sheets: [{
        ...sheet,
        cells: { ...sheet.cells, B2: "=C2*2", B3: "=C3*2" }
      }],
      tables: [{
        ...workbook.tables[0],
        columns: workbook.tables[0].columns.map((column) =>
          column.name === "Standard" ? { ...column, calculatedFormula: "=C2*2" } : column
        )
      }]
    };

    const imported = await importWorkbookFromXlsx(await exportWorkbookToXlsx(adjusted));
    expect(imported.tables[0].columns.find((column) => column.name === "Standard")?.calculatedFormula).toBe(
      "=C2*2"
    );
    expect(imported.sheets[0].cells.B3).toBe("=C3*2");
  });

  it("creates a narrow native table artifact and rejects external dependencies", async () => {
    const workbook = nativeTableWorkbook();
    const bytes = await exportStructuredTableToXlsx(workbook, workbook.tables[0].id);
    expect(bytes).toBeInstanceOf(Uint8Array);
    expect(bytes.byteLength).toBeGreaterThan(0);

    const external: WorkbookModel = {
      ...workbook,
      sheets: [{
        ...workbook.sheets[0],
        cells: { ...workbook.sheets[0].cells, D2: "=Z99" }
      }]
    };
    await expect(exportStructuredTableToXlsx(external, external.tables[0].id)).rejects.toMatchObject({
      code: "TABLE_EXPORT_EXTERNAL_DEPENDENCY"
    });
    const named: WorkbookModel = {
      ...workbook,
      namedRanges: [{
        name: "TaxRate",
        sheetId: workbook.sheets[0].id,
        range: { start: { row: 20, column: 20 }, end: { row: 20, column: 20 } }
      }],
      sheets: [{
        ...workbook.sheets[0],
        cells: { ...workbook.sheets[0].cells, D2: "=TaxRate*B2" }
      }]
    };
    await expect(exportStructuredTableToXlsx(named, named.tables[0].id)).rejects.toMatchObject({
      code: "TABLE_EXPORT_EXTERNAL_DEPENDENCY"
    });
    const otherTableReference: WorkbookModel = {
      ...workbook,
      sheets: [{
        ...workbook.sheets[0],
        cells: { ...workbook.sheets[0].cells, D2: "=OtherTable[@Amount]" }
      }]
    };
    await expect(
      exportStructuredTableToXlsx(otherTableReference, otherTableReference.tables[0].id)
    ).rejects.toMatchObject({ code: "TABLE_EXPORT_EXTERNAL_DEPENDENCY" });
  });
});

describe("shared formula import", () => {
  it("translates shared-formula slave cells to their own formulas", async () => {
    // Build a file with a REAL shared formula (master B1, slaves B2:B3 via si refs)
    // the way third-party producers write them — not via our own export, which
    // writes full formulas per cell.
    const ExcelJS = (await import("exceljs")).default;
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("Data");
    worksheet.getCell("A1").value = 1;
    worksheet.getCell("A2").value = 2;
    worksheet.getCell("A3").value = 3;
    worksheet.fillFormula("B1:B3", "A1*2", [2, 4, 6]);
    const buffer = await workbook.xlsx.writeBuffer();

    // Fixture sanity: the file must genuinely contain shared-formula slaves.
    const reread = new ExcelJS.Workbook();
    await reread.xlsx.load(buffer);
    const slaveValue = reread.getWorksheet("Data")?.getCell("B2").value;
    expect(slaveValue).toHaveProperty("sharedFormula");

    const imported = await importWorkbookFromXlsx(buffer as ArrayBuffer);
    const cells = imported.sheets[0].cells;

    expect(cells["B1"]).toBe("=A1*2");
    // Slaves must get the TRANSLATED formula, not "=B1" (the master's address).
    expect(cells["B2"]).toBe("=A2*2");
    expect(cells["B3"]).toBe("=A3*2");
  });
});
