import { readFile } from "node:fs/promises";
import { resolve } from "node:path";
import {
  DOMParser,
  type Document as XmlDocument,
  type Element as XmlElement
} from "@xmldom/xmldom";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { describe, expect, it } from "vitest";
import { makeAboveFloorWorksheetPackage } from "../test/xlsxSecurityFixtures";
import { addSyntheticVbaProject } from "../test/xlsxImportFixtures";
import type { StructuredTable, WorkbookModel } from "../types";
import {
  patchNativeTableXml,
  prepareNativeTableXmlForExcelJs,
  readNativeCommentsXml,
  readNativeTableXml
} from "./xlsxTableXml";

const fixturePath = resolve("src/test/fixtures/xlsx/generated-sales-structured-table.xlsx");
const realFixturePath = resolve("src/test/fixtures/xlsx/exceljs-issue-1669.xlsx");
const chartFixturePath = resolve("src/test/fixtures/xlsx/variant-chart.xlsx");
const commentFixturePath = resolve("src/test/fixtures/xlsx/variant-comment.xlsx");
const tableFixturePath = resolve("src/test/fixtures/xlsx/variant-table.xlsx");
const ALIASED_COMMENT_SHEET_COUNT = 3_000;
const ALIASED_COMMENT_PARSE_BUDGET_MS = 250;

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
  range: { start: { row: 0, column: 0 }, end: { row: 47_619, column: 9 } },
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
  it("removes unsupported drawing and chart parts from the ExcelJS derivative", async () => {
    const source = await readFile(chartFixturePath);
    const sourceEntries = unzipSync(
      new Uint8Array(source.buffer, source.byteOffset, source.byteLength)
    );
    const relationshipsName = "xl/worksheets/_rels/sheet1.xml.rels";
    sourceEntries[relationshipsName] = strToU8(
      strFromU8(sourceEntries[relationshipsName]).replace(
        "</Relationships>",
        '<Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/xl/charts/help" TargetMode="External" Id="rIdExternal"/></Relationships>'
      )
    );
    const prepared = prepareNativeTableXmlForExcelJs(
      zipSync(sourceEntries)
    );
    const entries = unzipSync(prepared);
    const names = Object.keys(entries);

    expect(names.filter((name) => /^xl\/(?:drawings|charts)\//i.test(name))).toEqual([]);
    expect(strFromU8(entries["xl/worksheets/sheet1.xml"])).not.toMatch(/<drawing\b/i);
    expect(strFromU8(entries["xl/worksheets/_rels/sheet1.xml.rels"])).not.toContain(
      "/relationships/drawing"
    );
    expect(strFromU8(entries["xl/worksheets/_rels/sheet1.xml.rels"])).toContain(
      "https://example.com/xl/charts/help"
    );
    expect(strFromU8(entries["[Content_Types].xml"])).not.toMatch(
      /PartName="\/xl\/(?:drawings|charts)\//i
    );
  });

  it("removes foreign comments from the ExcelJS derivative after native extraction", async () => {
    const source = await readFile(commentFixturePath);
    const prepared = prepareNativeTableXmlForExcelJs(
      new Uint8Array(source.buffer, source.byteOffset, source.byteLength)
    );
    const entries = unzipSync(prepared);

    expect(Object.keys(entries).filter((name) => /^xl\/comments(?:\/|\d)/i.test(name))).toEqual([]);
    expect(strFromU8(entries["xl/worksheets/_rels/sheet1.xml.rels"])).not.toContain(
      "/relationships/comments"
    );
    expect(strFromU8(entries["[Content_Types].xml"])).not.toContain("/xl/comments/");
  });

  it("memoizes aliased native-comment parts within a bounded parse duration", async () => {
    const source = await readFile(commentFixturePath);
    const entries = unzipSync(new Uint8Array(source.buffer, source.byteOffset, source.byteLength));
    const workbookXml = strFromU8(entries["xl/workbook.xml"]);
    const sheetXml = workbookXml.match(/<sheet\b[^>]*\/>/)?.[0];
    expect(sheetXml).toBeDefined();
    entries["xl/workbook.xml"] = strToU8(
      workbookXml.replace(
        sheetXml!,
        Array.from(
          { length: ALIASED_COMMENT_SHEET_COUNT },
          (_, index) => sheetXml!
            .replace('name="S"', `name="Alias ${index + 1}"`)
            .replace('sheetId="1"', `sheetId="${index + 1}"`)
        ).join("")
      )
    );
    const adversarial = zipSync(entries);

    const startedAt = performance.now();
    const comments = readNativeCommentsXml(adversarial);
    const durationMs = performance.now() - startedAt;

    expect(comments).toHaveLength(ALIASED_COMMENT_SHEET_COUNT);
    expect(comments[ALIASED_COMMENT_SHEET_COUNT - 1]).toEqual({
      sheetIndex: ALIASED_COMMENT_SHEET_COUNT - 1,
      sheetName: `Alias ${ALIASED_COMMENT_SHEET_COUNT}`,
      comments: { A1: "hello" }
    });
    expect(durationMs).toBeLessThan(ALIASED_COMMENT_PARSE_BUDGET_MS);
  });

  it("numbers native comments by worksheet order when a chartsheet comes first", async () => {
    const source = await readFile(commentFixturePath);
    const entries = unzipSync(new Uint8Array(source.buffer, source.byteOffset, source.byteLength));
    entries["xl/workbook.xml"] = strToU8(
      strFromU8(entries["xl/workbook.xml"]).replace(
        "<sheets>",
        '<sheets><sheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" name="Chart" sheetId="2" state="visible" r:id="rId4"/>'
      )
    );
    entries["xl/_rels/workbook.xml.rels"] = strToU8(
      strFromU8(entries["xl/_rels/workbook.xml.rels"]).replace(
        "</Relationships>",
        '<Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet" Target="/xl/chartsheets/sheet1.xml" Id="rId4"/></Relationships>'
      )
    );
    entries["xl/chartsheets/sheet1.xml"] = strToU8(
      '<chartsheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetViews><sheetView workbookViewId="0"/></sheetViews></chartsheet>'
    );
    entries["[Content_Types].xml"] = strToU8(
      strFromU8(entries["[Content_Types].xml"]).replace(
        "</Types>",
        '<Override PartName="/xl/chartsheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml"/></Types>'
      )
    );

    expect(readNativeCommentsXml(zipSync(entries))).toEqual([{
      sheetIndex: 0,
      sheetName: "S",
      comments: { A1: "hello" }
    }]);
  });

  it.each(["xl/notes.dat", "xl/notes.xml"])(
    "ignores comment relationships to unrecognized part %s",
    async (commentPart) => {
      const source = await readFile(commentFixturePath);
      const entries = unzipSync(new Uint8Array(source.buffer, source.byteOffset, source.byteLength));
      entries[commentPart] = entries["xl/comments/comment1.xml"];
      delete entries["xl/comments/comment1.xml"];
      entries["xl/worksheets/_rels/sheet1.xml.rels"] = strToU8(
        strFromU8(entries["xl/worksheets/_rels/sheet1.xml.rels"])
          .replace("/xl/comments/comment1.xml", `/${commentPart}`)
      );
      entries["[Content_Types].xml"] = strToU8(
        strFromU8(entries["[Content_Types].xml"])
          .replace("/xl/comments/comment1.xml", `/${commentPart}`)
      );

      expect(readNativeCommentsXml(zipSync(entries))).toEqual([{
        sheetIndex: 0,
        sheetName: "S",
        comments: {}
      }]);
    }
  );

  it.each(["A1048577", "XFE1", "A9007199254740992"])(
    "ignores out-of-grid comment reference %s",
    async (reference) => {
      const source = await readFile(commentFixturePath);
      const entries = unzipSync(new Uint8Array(source.buffer, source.byteOffset, source.byteLength));
      entries["xl/comments/comment1.xml"] = strToU8(
        strFromU8(entries["xl/comments/comment1.xml"]).replace('ref="A1"', `ref="${reference}"`)
      );

      expect(readNativeCommentsXml(zipSync(entries))).toEqual([{
        sheetIndex: 0,
        sheetName: "S",
        comments: {}
      }]);
    }
  );
  it("canonicalizes package-root table targets for ExcelJS", async () => {
    const source = await readFile(tableFixturePath);
    const prepared = prepareNativeTableXmlForExcelJs(
      new Uint8Array(source.buffer, source.byteOffset, source.byteLength)
    );
    const relationships = strFromU8(
      unzipSync(prepared)["xl/worksheets/_rels/sheet1.xml.rels"]
    );

    expect(relationships).toContain('Target="../tables/table1.xml"');
    expect(relationships).not.toContain('Target="/xl/tables/table1.xml"');
  });

  it("removes VBA projects from the ExcelJS derivative", async () => {
    const prepared = prepareNativeTableXmlForExcelJs(
      addSyntheticVbaProject(await fixture())
    );
    const entries = unzipSync(prepared);

    expect(entries["xl/vbaProject.bin"]).toBeUndefined();
    expect(strFromU8(entries["xl/_rels/workbook.xml.rels"])).not.toContain("vbaProject");
    expect(strFromU8(entries["[Content_Types].xml"])).not.toContain("macroEnabled");
    expect(strFromU8(entries["[Content_Types].xml"])).not.toContain("vbaProject.bin");
  });

  it("patches and revalidates an above-floor A1:J47620 worksheet archive", () => {
    const source = makeAboveFloorWorksheetPackage();
    const patched = patchNativeTableXml(source, [denseTable]);
    const patchedEntries = unzipSync(patched);
    const patchedTable = patchedEntries["xl/tables/table1.xml"];

    expect(patchedTable).toBeDefined();
    expect(strFromU8(patchedTable)).toContain('name="TableStyleMedium9"');
    expect(strFromU8(patchedEntries["xl/worksheets/sheet1.xml"]))
      .toContain('dimension ref="A1:J47620"');
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

  it.each([
    ["contains", "*Ea~*st~?~~*"],
    ["startsWith", "Ea~*st~?~~*"],
    ["endsWith", "*Ea~*st~?~~"]
  ] as const)("writes and reads %s filters with escaped Excel wildcards", async (operator, criterion) => {
    const source = await fixture();
    const patched = patchNativeTableXml(source, [{
      ...salesTable,
      filter: {
        kind: "comparison",
        columnId: "private-region",
        operator,
        value: { type: "string", value: "Ea*st?~" }
      }
    }]);
    const document = parseXml(tableXml(patched));
    const custom = firstByLocalName(document, "customFilter")!;

    expect(custom.getAttribute("operator")).toBeNull();
    expect(custom.getAttribute("val")).toBe(criterion);
    expect(readNativeTableXml(patched)[0].filter).toEqual({
      kind: "comparison",
      columnId: "Region",
      operator,
      value: { type: "string", value: "Ea*st?~" }
    });
  });

  it("writes and reads two-value notIn filters as ANDed not-equal comparisons", async () => {
    const source = await fixture();
    const patched = patchNativeTableXml(source, [{
      ...salesTable,
      filter: {
        kind: "set",
        columnId: "private-region",
        operator: "notIn",
        values: [
          { type: "string", value: "East" },
          { type: "string", value: "West" }
        ]
      }
    }]);
    const document = parseXml(tableXml(patched));
    const customFilters = firstByLocalName(document, "customFilters")!;
    const filters = Array.from(document.getElementsByTagName("*")).filter(
      (node) => node.localName === "customFilter"
    );

    expect(customFilters.getAttribute("and")).toBe("1");
    expect(filters.map((item) => [item.getAttribute("operator"), item.getAttribute("val")])).toEqual([
      ["notEqual", "East"],
      ["notEqual", "West"]
    ]);
    expect(readNativeTableXml(patched)[0].filter).toEqual({
      kind: "set",
      columnId: "Region",
      operator: "notIn",
      values: [
        { type: "string", value: "East" },
        { type: "string", value: "West" }
      ]
    });
  });
});
