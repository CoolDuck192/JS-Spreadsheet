import { describe, expect, it } from "vitest";
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
import { exportWorkbookToXlsx, importWorkbookFromXlsx } from "./xlsx";

describe("xlsx", () => {
  it("round-trips sheets, values, formulas, comments, hyperlinks, merges, and dimensions", async () => {
    let workbook = createBlankWorkbook();
    const firstSheetId = workbook.activeSheetId;
    workbook = setCellContent(workbook, firstSheetId, "A1", "Project");
    workbook = setCellContent(workbook, firstSheetId, "B2", 42);
    workbook = setCellContent(workbook, firstSheetId, "C3", true);
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
    expect(getCellContent(imported, importedFirstSheetId, "B2")).toBe(42);
    expect(getCellContent(imported, importedFirstSheetId, "C3")).toBe(true);
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
