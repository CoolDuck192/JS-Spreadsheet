import { readFile } from "node:fs/promises";
import { resolve } from "node:path";
import {
  DOMParser,
  type Document as XmlDocument,
  type Element as XmlElement
} from "@xmldom/xmldom";
import { strFromU8, unzipSync } from "fflate";
import { describe, expect, it } from "vitest";
import { makeDenseWorksheetPackage } from "../test/xlsxSecurityFixtures";
import type { StructuredTable, WorkbookModel } from "../types";
import { validateXlsxArchive } from "./xlsxSecurity";
import {
  patchNativeTableXml,
  prepareNativeTableXmlForExcelJs,
  readNativeTableXml
} from "./xlsxTableXml";

const fixturePath = resolve("src/test/fixtures/xlsx/generated-sales-structured-table.xlsx");
const realFixturePath = resolve("src/test/fixtures/xlsx/exceljs-issue-1669.xlsx");

async function fixture(name: "generated" | "real" = "generated") {
  const bytes = await readFile(name === "generated" ? fixturePath : realFixturePath);
  return new Uint8Array(bytes.buffer, bytes.byteOffset, bytes.byteLength);
}

const salesTable: StructuredTable = {
  id: "private-table-id",
  name: "SalesTable",
  sheetId: "private-sheet-id",
  range: { start: { row: 0, column: 0 }, end: { row: 5, column: 5 } },
  headerRow: true,
  totalsRow: true,
  columns: [
    { id: "private-order-id", name: "Order ID", sheetColumn: 0, totalsLabel: "Grand Total" },
    { id: "private-order-date", name: "Order Date", sheetColumn: 1 },
    { id: "private-region", name: "Region", sheetColumn: 2 },
    { id: "private-units", name: "Units", sheetColumn: 3 },
    { id: "private-unit-price", name: "Unit Price", sheetColumn: 4 },
    {
      id: "private-amount",
      name: "Amount",
      sheetColumn: 5,
      calculatedFormula: "=D2+E2",
      totalsFunction: "average"
    }
  ],
  rowIds: ["private-row-1", "private-row-2", "private-row-3", "private-row-4"],
  keyColumnId: "private-order-id",
  style: {
    theme: "TableStyleMedium4",
    showFirstColumn: true,
    showLastColumn: false,
    showRowStripes: false,
    showColumnStripes: true
  },
  filter: {
    kind: "set",
    columnId: "private-region",
    operator: "in",
    values: [
      { type: "string", value: "East" },
      { type: "string", value: "North" }
    ]
  }
};

const denseTable: StructuredTable = {
  id: "dense-table-id",
  name: "DenseTable",
  sheetId: "dense-sheet-id",
  range: { start: { row: 0, column: 0 }, end: { row: 99_999, column: 9 } },
  headerRow: true,
  totalsRow: false,
  columns: Array.from({ length: 10 }, (_, index) => ({
    id: `dense-column-${index + 1}`,
    name: `Column${index + 1}`,
    sheetColumn: index
  })),
  rowIds: [],
  keyColumnId: "dense-column-1",
  style: {
    theme: "TableStyleMedium9",
    showFirstColumn: false,
    showLastColumn: true,
    showRowStripes: true,
    showColumnStripes: false
  }
};

function tableXml(data: Uint8Array, name = "xl/tables/table1.xml") {
  return strFromU8(unzipSync(data)[name]);
}

function readU32(bytes: Uint8Array, offset: number): number {
  return (
    (bytes[offset] |
      (bytes[offset + 1] << 8) |
      (bytes[offset + 2] << 16) |
      (bytes[offset + 3] << 24)) >>>
    0
  );
}

function writeU16(bytes: Uint8Array, offset: number, value: number): void {
  bytes[offset] = value & 0xff;
  bytes[offset + 1] = (value >>> 8) & 0xff;
}

function writeU32(bytes: Uint8Array, offset: number, value: number): void {
  bytes[offset] = value & 0xff;
  bytes[offset + 1] = (value >>> 8) & 0xff;
  bytes[offset + 2] = (value >>> 16) & 0xff;
  bytes[offset + 3] = (value >>> 24) & 0xff;
}

function withEocdLikeZipComment(data: Uint8Array): Uint8Array {
  let eocdOffset = data.length - 22;
  while (eocdOffset >= 0 && readU32(data, eocdOffset) !== 0x06054b50) {
    eocdOffset -= 1;
  }
  if (eocdOffset < 0) throw new Error("EOCD not found in test archive");

  const comment = new Uint8Array(24);
  writeU32(comment, 0, 0x06054b50);
  const result = new Uint8Array(data.length + comment.length);
  result.set(data);
  result.set(comment, data.length);
  writeU16(result, eocdOffset + 20, comment.length);
  return result;
}

function parseXml(xml: string) {
  const errors: string[] = [];
  const document = new DOMParser({ onError: (_level, message) => errors.push(message) }).parseFromString(
    xml,
    "application/xml"
  );
  expect(errors).toEqual([]);
  return document;
}

function firstByLocalName(document: XmlDocument | XmlElement, name: string) {
  return Array.from(document.getElementsByTagName("*")).find((node) => node.localName === name);
}

describe("xlsxTableXml", () => {
  it("patches and revalidates a dense A1:J100000 worksheet archive", () => {
    const source = makeDenseWorksheetPackage();
    const sourceEntries = unzipSync(source);
    const patched = patchNativeTableXml(source, [denseTable]);
    const patchedEntries = unzipSync(patched);
    const sourceTable = sourceEntries["xl/tables/table1.xml"];
    const patchedTable = patchedEntries["xl/tables/table1.xml"];

    expect(sourceTable).toBeDefined();
    expect(patchedTable).toBeDefined();
    expect(strFromU8(patchedTable)).not.toEqual(strFromU8(sourceTable));
    expect(strFromU8(patchedTable)).toContain('name="TableStyleMedium9"');
    expect(strFromU8(patchedEntries["xl/worksheets/sheet1.xml"]))
      .toContain('dimension ref="A1:J100000"');
    expect(validateXlsxArchive(patched)).toEqual({ ok: true });
  });

  it("reads native table XML from the validated archive view when a ZIP comment resembles an EOCD", async () => {
    const source = await fixture();
    const commented = withEocdLikeZipComment(source);

    expect(readNativeTableXml(commented)).toEqual(readNativeTableXml(source));
  });

  it("patches the validated archive view when a ZIP comment resembles an EOCD", async () => {
    const source = await fixture();
    const commented = withEocdLikeZipComment(source);

    expect(readNativeTableXml(patchNativeTableXml(commented, []))).toEqual(
      readNativeTableXml(source)
    );
  });

  it("prepares the validated archive view when a ZIP comment resembles an EOCD", async () => {
    const source = await fixture();
    const commented = withEocdLikeZipComment(source);
    const prepared = prepareNativeTableXmlForExcelJs(commented);

    expect(unzipSync(prepared)["xl/workbook.xml"]).toBeDefined();
  });

  it("reads calculated columns, filters, and totals from the deterministic fixture", async () => {
    expect(readNativeTableXml(await fixture())).toEqual([
      {
        name: "SalesTable",
        calculatedColumns: { Amount: "=[@Units]*[@[Unit Price]]" },
        filter: {
          kind: "set",
          columnId: "Region",
          operator: "in",
          values: [
            { type: "string", value: "East" },
            { type: "string", value: "West" }
          ]
        },
        totals: {
          "Order ID": { label: "Total" },
          Amount: { function: "sum" }
        }
      }
    ]);
  });

  it("reads custom and multi-value filters from the independent Excel fixture", async () => {
    expect(readNativeTableXml(await fixture("real"))).toEqual([
      {
        name: "Table1",
        calculatedColumns: {},
        filter: {
          kind: "comparison",
          columnId: "Column1",
          operator: "neq",
          value: { type: "string", value: "4" }
        },
        totals: {}
      },
      {
        name: "Table2",
        calculatedColumns: {},
        filter: {
          kind: "set",
          columnId: "DK",
          operator: "in",
          values: [
            { type: "string", value: "T123456789" },
            { type: "string", value: "T123456791" },
            { type: "string", value: "T123456793" }
          ]
        },
        totals: {}
      }
    ]);
  });

  it("patches only native table metadata and never leaks application IDs", async () => {
    const source = await fixture();
    const patched = patchNativeTableXml(source, [salesTable]);
    const xml = tableXml(patched);
    const document = parseXml(xml);
    const style = firstByLocalName(document, "tableStyleInfo");
    const columns = Array.from(document.getElementsByTagName("*")).filter(
      (node) => node.localName === "tableColumn"
    );
    const amount = columns.find((node) => node.getAttribute("name") === "Amount");
    const orderId = columns.find((node) => node.getAttribute("name") === "Order ID");

    expect(style?.getAttribute("name")).toBe("TableStyleMedium4");
    expect(style?.getAttribute("showFirstColumn")).toBe("1");
    expect(style?.getAttribute("showRowStripes")).toBe("0");
    expect(style?.getAttribute("showColumnStripes")).toBe("1");
    expect(amount?.getAttribute("totalsRowFunction")).toBe("average");
    expect(firstByLocalName(amount as XmlElement, "calculatedColumnFormula")?.textContent).toBe(
      "SalesTable[@Units]+SalesTable[@[Unit Price]]"
    );
    expect(orderId?.getAttribute("totalsRowLabel")).toBe("Grand Total");
    expect(xml).toContain('<filter val="East"/>');
    expect(xml).toContain('<filter val="North"/>');
    expect(xml).not.toContain('<filter val="West"/>');
    for (const secret of [
      salesTable.id,
      salesTable.sheetId,
      salesTable.keyColumnId,
      ...salesTable.columns.map((column) => column.id),
      ...salesTable.rowIds
    ]) {
      expect(xml).not.toContain(secret);
    }

    const sourceEntries = unzipSync(source);
    const patchedEntries = unzipSync(patched);
    expect(patchedEntries["xl/workbook.xml"]).toEqual(sourceEntries["xl/workbook.xml"]);
  });

  it("preserves quoted sheet qualifiers in calculated and custom totals formulas", async () => {
    const source = await fixture();
    const table: StructuredTable = {
      ...salesTable,
      columns: salesTable.columns.map((column) =>
        column.name === "Amount"
          ? { ...column, calculatedFormula: "='A1'!B2+C2" }
          : column
      )
    };
    const workbook: WorkbookModel = {
      version: 2,
      activeSheetId: table.sheetId,
      sheets: [{
        id: table.sheetId,
        name: "Sales",
        rowCount: 6,
        columnCount: 6,
        cells: { D6: "='Sheet''s'!B2+C2" },
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
      }],
      namedRanges: [],
      tables: [table]
    };

    const patched = patchNativeTableXml(source, [table], workbook);
    const document = parseXml(tableXml(patched));
    const columns = Array.from(document.getElementsByTagName("*")).filter(
      (node) => node.localName === "tableColumn"
    );
    const amount = columns.find((node) => node.getAttribute("name") === "Amount")!;
    const units = columns.find((node) => node.getAttribute("name") === "Units")!;

    expect(firstByLocalName(amount, "calculatedColumnFormula")?.textContent).toBe(
      "'A1'!B2+SalesTable[@Region]"
    );
    expect(units.getAttribute("totalsRowFunction")).toBe("custom");
    expect(firstByLocalName(units, "totalsRowFormula")?.textContent).toBe(
      "'Sheet''s'!B2+SalesTable[@Region]"
    );
    expect(readNativeTableXml(patched)[0]).toMatchObject({
      calculatedColumns: { Amount: "='A1'!B2+SalesTable[@Region]" },
      totals: {
        Units: {
          function: "custom",
          formula: "='Sheet''s'!B2+SalesTable[@Region]"
        }
      }
    });
  });

  it("writes custom filter comparisons and produces stable bytes", async () => {
    const source = await fixture();
    const customTable: StructuredTable = {
      ...salesTable,
      filter: {
        kind: "comparison",
        columnId: "private-units",
        operator: "neq",
        value: { type: "number", value: 4 }
      }
    };
    const first = patchNativeTableXml(source, [customTable]);
    const second = patchNativeTableXml(source, [customTable]);

    expect(first).toEqual(second);
    expect(tableXml(first)).toContain('<customFilter operator="notEqual" val="4"/>');
  });

  it("writes same-column OR comparisons as native custom filters", async () => {
    const source = await fixture();
    const patched = patchNativeTableXml(source, [{
      ...salesTable,
      filter: {
        kind: "logical",
        operator: "or",
        operands: [
          {
            kind: "comparison",
            columnId: "private-units",
            operator: "eq",
            value: { type: "number", value: 5 }
          },
          {
            kind: "comparison",
            columnId: "private-units",
            operator: "eq",
            value: { type: "number", value: 20 }
          }
        ]
      }
    }]);
    const document = parseXml(tableXml(patched));
    const customFilters = firstByLocalName(document, "customFilters")!;
    const filters = Array.from(document.getElementsByTagName("*")).filter(
      (node) => node.localName === "customFilter"
    );

    expect(customFilters.getAttribute("and")).toBe("0");
    expect(filters.map((filter) => filter.getAttribute("val"))).toEqual(["5", "20"]);
  });

  it.each([
    ["logical comparisons", {
      kind: "logical" as const,
      operator: "and" as const,
      operands: [
        {
          kind: "comparison" as const,
          columnId: "private-units",
          operator: "gte" as const,
          value: { type: "number" as const, value: 5 }
        },
        {
          kind: "comparison" as const,
          columnId: "private-units",
          operator: "lte" as const,
          value: { type: "number" as const, value: 10 }
        }
      ]
    }],
    ["a between range", {
      kind: "range" as const,
      columnId: "private-units",
      operator: "between" as const,
      lower: { type: "number" as const, value: 5 },
      upper: { type: "number" as const, value: 10 }
    }]
  ])("writes %s as ANDed native custom filters", async (_label, filter) => {
    const source = await fixture();
    const patched = patchNativeTableXml(source, [{ ...salesTable, filter }]);
    const document = parseXml(tableXml(patched));
    const customFilters = firstByLocalName(document, "customFilters")!;
    const filters = Array.from(document.getElementsByTagName("*")).filter(
      (node) => node.localName === "customFilter"
    );

    expect(customFilters.getAttribute("and")).toBe("1");
    expect(filters.map((item) => [item.getAttribute("operator"), item.getAttribute("val")])).toEqual([
      ["greaterThanOrEqual", "5"],
      ["lessThanOrEqual", "10"]
    ]);
  });

  it("rejects unsupported filter expressions instead of silently dropping them", async () => {
    const source = await fixture();
    expect(() => patchNativeTableXml(source, [{
      ...salesTable,
      filter: {
        kind: "comparison",
        columnId: "private-region",
        operator: "contains",
        value: { type: "string", value: "East" }
      }
    }])).toThrow(/unsupported native table filter/i);
  });
});
