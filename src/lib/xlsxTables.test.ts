import { readFile } from "node:fs/promises";
import { resolve } from "node:path";
import ExcelJS from "exceljs";
import { describe, expect, it } from "vitest";
import type { IdGenerator } from "../core/ids";
import type { SheetModel, StructuredTable, WorkbookModel } from "../types";
import { readNativeTableXml } from "./xlsxTableXml";
import {
  addStructuredTablesToWorksheet,
  importStructuredTablesFromWorksheet,
  listWorksheetTables
} from "./xlsxTables";

const generatedFixturePath = resolve("src/test/fixtures/xlsx/generated-sales-structured-table.xlsx");
const realFixturePath = resolve("src/test/fixtures/xlsx/exceljs-issue-1669.xlsx");

function sequencedIds(prefix: string): IdGenerator {
  let index = 0;
  return (kind) => `${prefix}-${kind}-${++index}`;
}

async function loadFixture(path: string) {
  const bytes = new Uint8Array(await readFile(path));
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(bytes as unknown as ExcelJS.Buffer);
  return { bytes, workbook };
}

function blankSheet(id = "sheet-sales", name = "Sales"): SheetModel {
  return {
    id,
    name,
    rowCount: 100,
    columnCount: 26,
    cells: {},
    formats: {},
    columnWidths: {},
    rowHeights: {},
    comments: {},
    hyperlinks: {},
    validations: {},
    conditionalFormats: [],
    filters: [],
    charts: [],
    merges: [],
    protection: { isProtected: false, lockedCells: {}, unlockedCells: {} }
  };
}

describe("xlsxTables", () => {
  it("imports the generated native table with formulas, totals, styles, and a remapped filter", async () => {
    const { bytes, workbook } = await loadFixture(generatedFixturePath);
    const worksheet = workbook.getWorksheet("Sales")!;
    const tables = await importStructuredTablesFromWorksheet(
      worksheet,
      "sheet-sales",
      readNativeTableXml(bytes),
      { idGenerator: sequencedIds("generated") }
    );

    expect(tables).toHaveLength(1);
    const table = tables[0];
    expect(table).toMatchObject({
      name: "SalesTable",
      sheetId: "sheet-sales",
      range: { start: { row: 0, column: 0 }, end: { row: 5, column: 5 } },
      headerRow: true,
      totalsRow: true,
      style: {
        theme: "TableStyleLight9",
        showFirstColumn: false,
        showLastColumn: false,
        showRowStripes: true,
        showColumnStripes: false
      }
    });
    expect(table.columns.map((column) => column.name)).toEqual([
      "Order ID",
      "Order Date",
      "Region",
      "Units",
      "Unit Price",
      "Amount"
    ]);
    expect(table.columns.find((column) => column.name === "Order Date")?.dataType).toBe("date");
    expect(table.columns.find((column) => column.name === "Amount")).toMatchObject({
      calculatedFormula: "=D2*E2",
      totalsFunction: "sum"
    });
    expect(table.columns.find((column) => column.name === "Order ID")?.totalsLabel).toBe("Total");
    expect(table.rowIds).toHaveLength(4);
    const region = table.columns.find((column) => column.name === "Region")!;
    expect(table.filter).toEqual({
      kind: "set",
      columnId: region.id,
      operator: "in",
      values: [
        { type: "string", value: "East" },
        { type: "string", value: "West" }
      ]
    });
    expect([2, 3, 4, 5].map((row) => worksheet.getCell(`F${row}`).formula)).toEqual([
      "D2*E2",
      "D3*E3",
      "D4*E4",
      "D5*E5"
    ]);
    expect(worksheet.getCell("B2").value).toBeInstanceOf(Date);
  });

  it("imports a custom totals formula with a single structured item specifier", async () => {
    const { bytes, workbook } = await loadFixture(generatedFixturePath);
    const worksheet = workbook.getWorksheet("Sales")!;
    const metadata = readNativeTableXml(bytes).map((entry) => entry.name === "SalesTable"
      ? {
          ...entry,
          totals: {
            ...entry.totals,
            Amount: {
              ...entry.totals.Amount,
              formula: "=COUNTA(SalesTable[#Data])"
            }
          }
        }
      : entry);

    await expect(importStructuredTablesFromWorksheet(
      worksheet,
      "sheet-sales",
      metadata,
      { idGenerator: sequencedIds("single-selector") }
    )).resolves.toHaveLength(1);
    expect(worksheet.getCell("F6").formula).toBe("COUNTA(A2:F5)");
  });

  it("degrades cross-table calculated-column metadata without aborting import", async () => {
    const { bytes, workbook } = await loadFixture(generatedFixturePath);
    const worksheet = workbook.getWorksheet("Sales")!;
    const metadata = readNativeTableXml(bytes).map((entry) => entry.name === "SalesTable"
      ? {
          ...entry,
          calculatedColumns: {
            ...entry.calculatedColumns,
            Amount: "=VLOOKUP(SalesTable[[#This Row],[Order ID]],Dim[#All],2,0)"
          }
        }
      : entry);

    const tables = await importStructuredTablesFromWorksheet(
      worksheet,
      "sheet-sales",
      metadata,
      { idGenerator: sequencedIds("cross-table") }
    );

    expect(tables[0].columns.find((column) => column.name === "Amount")?.calculatedFormula).toBeUndefined();
    expect(worksheet.getCell("F2").formula).toBe("D2*E2");
    expect(worksheet.getCell("F5").formula).toBe("D5*E5");
  });

  it("imports both independently authored tables with default headers and native filters", async () => {
    const { bytes, workbook } = await loadFixture(realFixturePath);
    const metadata = readNativeTableXml(bytes);
    const table1 = (await importStructuredTablesFromWorksheet(
      workbook.getWorksheet("Sheet1")!,
      "sheet-1",
      metadata,
      { idGenerator: sequencedIds("one") }
    ))[0];
    const table2 = (await importStructuredTablesFromWorksheet(
      workbook.getWorksheet("Sheet2")!,
      "sheet-2",
      metadata,
      { idGenerator: sequencedIds("two") }
    ))[0];

    expect(table1).toMatchObject({
      name: "Table1",
      range: { start: { row: 0, column: 0 }, end: { row: 5, column: 1 } },
      headerRow: true,
      totalsRow: false,
      style: { theme: "TableStyleMedium2" }
    });
    expect(table1.filter).toEqual({
      kind: "comparison",
      columnId: table1.columns[0].id,
      operator: "neq",
      value: { type: "number", value: 4 }
    });
    expect(table2).toMatchObject({
      name: "Table2",
      range: { start: { row: 0, column: 0 }, end: { row: 5, column: 1 } },
      headerRow: true,
      totalsRow: false,
      style: { theme: "TableStyleLight9" }
    });
    expect(table2.columns.map((column) => column.name)).toEqual(["DK", "Name"]);
    expect(table2.filter).toEqual({
      kind: "set",
      columnId: table2.columns[0].id,
      operator: "in",
      values: [
        { type: "string", value: "T123456789" },
        { type: "string", value: "T123456791" },
        { type: "string", value: "T123456793" }
      ]
    });
  });

  it("derives stable keyed row IDs while regenerating table and column identities", async () => {
    const { bytes, workbook } = await loadFixture(generatedFixturePath);
    const worksheet = workbook.getWorksheet("Sales")!;
    const metadata = readNativeTableXml(bytes);
    const original = await importStructuredTablesFromWorksheet(worksheet, "sheet", metadata, {
      idGenerator: sequencedIds("first"),
      tableKeys: { salestable: { columnName: "order id" } }
    });
    const originalOrderIds = [2, 3, 4, 5].map((row) => String(worksheet.getCell(`A${row}`).value));
    const row2 = worksheet.getRow(2).values;
    worksheet.getRow(2).values = worksheet.getRow(5).values;
    worksheet.getRow(5).values = row2;
    for (const row of [2, 3, 4, 5]) worksheet.getCell(`F${row}`).value = { formula: `D${row}*E${row}` };
    const reordered = await importStructuredTablesFromWorksheet(worksheet, "sheet", metadata, {
      idGenerator: sequencedIds("second"),
      tableKeys: { SalesTable: { columnName: "Order ID" } }
    });

    expect(original[0].id).not.toBe(reordered[0].id);
    expect(original[0].columns.map((column) => column.id)).not.toEqual(
      reordered[0].columns.map((column) => column.id)
    );
    const originalByOrder = new Map(originalOrderIds.map((orderId, index) => [
      orderId,
      original[0].rowIds[index]
    ]));
    const reorderedByOrder = new Map([2, 3, 4, 5].map((row, index) => [
      String(worksheet.getCell(`A${row}`).value),
      reordered[0].rowIds[index]
    ]));
    expect(reorderedByOrder).toEqual(originalByOrder);

    const randomA = await importStructuredTablesFromWorksheet(worksheet, "sheet", metadata, {
      idGenerator: sequencedIds("random-a")
    });
    const randomB = await importStructuredTablesFromWorksheet(worksheet, "sheet", metadata, {
      idGenerator: sequencedIds("random-b")
    });
    expect(randomA[0].rowIds).not.toEqual(randomB[0].rowIds);
  });

  it("rejects blank and duplicate configured keys as one failed import", async () => {
    const { bytes, workbook } = await loadFixture(generatedFixturePath);
    const worksheet = workbook.getWorksheet("Sales")!;
    const metadata = readNativeTableXml(bytes);
    worksheet.getCell("A3").value = "";
    await expect(importStructuredTablesFromWorksheet(worksheet, "sheet", metadata, {
      idGenerator: sequencedIds("blank"),
      tableKeys: { SalesTable: { columnName: "Order ID" } }
    })).rejects.toMatchObject({ code: "XLSX_TABLE_KEY_INVALID" });

    worksheet.getCell("A3").value = worksheet.getCell("A2").value;
    await expect(importStructuredTablesFromWorksheet(worksheet, "sheet", metadata, {
      idGenerator: sequencedIds("duplicate"),
      tableKeys: { SalesTable: { columnName: "Order ID" } }
    })).rejects.toMatchObject({ code: "XLSX_TABLE_KEY_INVALID" });
  });

  it("adds native ExcelJS tables without overwriting typed worksheet cells", () => {
    const worksheet = new ExcelJS.Workbook().addWorksheet("Sales");
    const sheet = blankSheet();
    sheet.cells = {
      A1: "Order ID", B1: "Amount",
      A2: "SO-1", B2: "=2*5",
      A3: "SO-2", B3: 12,
      A4: "Total", B4: "=SUM(B2:B3)"
    };
    for (const [address, value] of Object.entries(sheet.cells)) {
      worksheet.getCell(address).value = typeof value === "string" && value.startsWith("=")
        ? { formula: value.slice(1) }
        : value;
    }
    const table: StructuredTable = {
      id: "table-id",
      name: "SalesTable",
      sheetId: sheet.id,
      range: { start: { row: 0, column: 0 }, end: { row: 3, column: 1 } },
      headerRow: true,
      totalsRow: true,
      columns: [
        { id: "order", name: "Order ID", sheetColumn: 0, totalsLabel: "Total" },
        { id: "amount", name: "Amount", sheetColumn: 1, totalsFunction: "sum" }
      ],
      rowIds: ["row-1", "row-2"],
      style: { theme: "TableStyleLight9", showRowStripes: true }
    };
    const workbook: WorkbookModel = {
      version: 2,
      activeSheetId: sheet.id,
      sheets: [sheet],
      namedRanges: [],
      tables: [table]
    };

    addStructuredTablesToWorksheet(workbook, sheet, worksheet);

    expect(listWorksheetTables(worksheet)).toHaveLength(1);
    expect(worksheet.getTable("SalesTable").name).toBe("SalesTable");
    expect(worksheet.getCell("B2").formula).toBe("2*5");
    expect(worksheet.getCell("B4").formula).toBe("SUM(B2:B3)");
  });

  it("rejects malformed runtime table arrays and invalid native names", async () => {
    expect(() => listWorksheetTables({ getTables: () => [[{}, undefined]] } as unknown as ExcelJS.Worksheet))
      .toThrow(/invalid native table list/i);

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("Bad");
    worksheet.addTable({
      name: "ValidName",
      ref: "A1",
      headerRow: true,
      totalsRow: false,
      columns: [{ name: "A" }],
      rows: [[1]]
    });
    const native = listWorksheetTables(worksheet)[0] as unknown as { table: { name: string; displayName: string } };
    native.table.name = "A1";
    native.table.displayName = "A1";
    await expect(importStructuredTablesFromWorksheet(worksheet, "sheet", [], {})).rejects.toMatchObject({
      code: "XLSX_TABLE_NAME_INVALID"
    });
  });
});
