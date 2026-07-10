import { readFile } from "node:fs/promises";
import { resolve } from "node:path";
import {
  DOMParser,
  type Document as XmlDocument,
  type Element as XmlElement
} from "@xmldom/xmldom";
import { strFromU8, unzipSync } from "fflate";
import { describe, expect, it } from "vitest";
import type { StructuredTable } from "../types";
import { patchNativeTableXml, readNativeTableXml } from "./xlsxTableXml";

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

function tableXml(data: Uint8Array, name = "xl/tables/table1.xml") {
  return strFromU8(unzipSync(data)[name]);
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
