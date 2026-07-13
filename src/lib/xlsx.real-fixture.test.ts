import { readFile } from "node:fs/promises";
import { resolve } from "node:path";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { describe, expect, it } from "vitest";
import type { FilterExpression } from "../table/core/query";
import type { StructuredTable, WorkbookModel } from "../types";
import { migrateWorkbookModel } from "../core/workbook/migrateWorkbook";
import { formatCellAddress } from "./addressing";
import { createBlankWorkbook, setCellComment, setCellContent } from "./workbook";
import { exportWorkbookToXlsx, importWorkbookFromXlsx } from "./xlsx";

const fixtureDirectory = resolve("src/test/fixtures/xlsx");

async function fixture(name: string) {
  const bytes = await readFile(resolve(fixtureDirectory, name));
  return new Uint8Array(bytes.buffer, bytes.byteOffset, bytes.byteLength);
}

function summarizeTable(table: StructuredTable) {
  return {
    name: table.name,
    range: table.range,
    headerRow: table.headerRow,
    totalsRow: table.totalsRow,
    columns: table.columns.map((column) => column.name),
    style: table.style,
    filter: table.filter ? summarizeFilter(table.filter, table) : undefined
  };
}

function summarizeFilter(filter: FilterExpression, table: StructuredTable): unknown {
  if (filter.kind === "logical") {
    return { ...filter, operands: filter.operands.map((operand) => summarizeFilter(operand, table)) };
  }
  if (filter.kind === "not") return { ...filter, operand: summarizeFilter(filter.operand, table) };
  return {
    ...filter,
    columnId: table.columns.find((column) => column.id === filter.columnId)?.name
  };
}

function tableXmlStrings(bytes: Uint8Array): string[] {
  const entries = unzipSync(bytes);
  return Object.keys(entries)
    .filter((name) => /^xl\/tables\/table\d+\.xml$/i.test(name))
    .sort()
    .map((name) => strFromU8(entries[name]));
}

describe("real native XLSX fixtures", () => {
  it("imports openpyxl worksheets containing unsupported charts", async () => {
    const workbook = await importWorkbookFromXlsx(await fixture("variant-chart.xlsx"));

    expect(workbook.sheets).toHaveLength(1);
    expect(workbook.sheets[0].cells).toMatchObject({ A1: "Q", B5: 200 });
    expect(workbook.sheets[0].charts).toEqual([]);
  });

  it("imports openpyxl comments from foreign comment-part layouts", async () => {
    const workbook = await importWorkbookFromXlsx(await fixture("variant-comment.xlsx"));

    expect(workbook.sheets[0].comments.A1).toBe("hello");
  });

  it("preserves application-authored comments through the native XML path", async () => {
    let source = createBlankWorkbook();
    source = setCellContent(source, source.activeSheetId, "A1", "commented");
    source = setCellComment(source, source.activeSheetId, "A1", "round-trip note");

    const workbook = await importWorkbookFromXlsx(await exportWorkbookToXlsx(source));

    expect(workbook.sheets[0].comments.A1).toBe("round-trip note");
  });

  it("drops an XFD1048576 orphan comment without expanding a tiny sheet", async () => {
    const entries = unzipSync(await fixture("variant-comment.xlsx"));
    entries["xl/comments/comment1.xml"] = strToU8(
      strFromU8(entries["xl/comments/comment1.xml"])
        .replace('ref="A1"', 'ref="XFD1048576"')
    );

    const workbook = await importWorkbookFromXlsx(zipSync(entries));

    expect(workbook.sheets[0]).toMatchObject({
      rowCount: 100,
      columnCount: 26
    });
    expect(workbook.sheets[0].comments).toEqual({});
  });

  it("drops an AZ200 comment-only cell outside the default import grid", async () => {
    let source = createBlankWorkbook();
    source = setCellComment(source, source.activeSheetId, "AZ200", "orphaned note");

    const workbook = await importWorkbookFromXlsx(await exportWorkbookToXlsx(source));

    expect(workbook.sheets[0]).toMatchObject({
      rowCount: 100,
      columnCount: 26
    });
    expect(workbook.sheets[0].comments).toEqual({});
  });

  it("imports openpyxl root-relative table relationships without renaming the table", async () => {
    const workbook = await importWorkbookFromXlsx(await fixture("variant-table.xlsx"));
    const table = workbook.tables[0];

    expect(table.name).toBe("T1");
    expect(table.range).toEqual({ start: { row: 0, column: 0 }, end: { row: 4, column: 1 } });
    expect(table.columns.map((column) => column.name)).toEqual(["Q", "V"]);
    expect(table.rowIds).toHaveLength(4);
    expect(table.style).toEqual({
      theme: "TableStyleMedium9",
      showFirstColumn: false,
      showLastColumn: false,
      showRowStripes: true,
      showColumnStripes: false
    });
    expect(migrateWorkbookModel(workbook)).not.toBeNull();

    const roundTripped = await importWorkbookFromXlsx(await exportWorkbookToXlsx(workbook));
    expect(roundTripped.tables[0].name).toBe("T1");
  });

  it("imports a trimmed openpyxl kitchen sink without losing supported workbook data", async () => {
    const source = await fixture("real-kitchen-sink-trimmed.xlsx");
    expect(source.byteLength).toBeLessThanOrEqual(100 * 1024);

    const workbook = await importWorkbookFromXlsx(source);
    const sales = workbook.sheets.find((sheet) => sheet.name === "Sales")!;
    const dashboard = workbook.sheets.find((sheet) => sheet.name === "Dashboard")!;
    const dense = workbook.sheets.find((sheet) => sheet.name === "Data10k")!;
    const table = workbook.tables.find((candidate) => candidate.name === "SalesTable")!;

    expect(workbook.sheets).toHaveLength(3);
    expect(table.range).toEqual({ start: { row: 0, column: 0 }, end: { row: 50, column: 5 } });
    expect(table.rowIds).toHaveLength(50);
    expect(Array.from({ length: 50 }, (_, index) => sales.cells[`F${index + 2}`])).toEqual(
      Array.from({ length: 50 }, (_, index) => `=D${index + 2}*E${index + 2}`)
    );
    expect(Array.from({ length: 50 }, (_, index) => sales.cells[`A${index + 2}`])).toEqual(
      Array.from({ length: 50 }, (_, index) => 46_055 + index)
    );
    expect(Array.from({ length: 50 }, (_, index) => {
      const row = index + 2;
      return [
        sales.cells[`B${row}`],
        sales.cells[`C${row}`],
        sales.cells[`D${row}`],
        sales.cells[`E${row}`]
      ];
    })).toEqual(Array.from({ length: 50 }, (_, index) => [
      `Rep ${index % 5 + 1}`,
      ["South", "East", "West", "North"][index % 4],
      (index + 1) * 2,
      Number((9.99 + ((index + 1) % 7) * 5).toFixed(2))
    ]));
    expect(sales.formats.A2).toMatchObject({ numberFormat: "date" });
    expect(sales.freezeTopRow).toBe(true);
    expect(sales.charts).toEqual([]);
    expect(dashboard.comments.A10).toBe("Checked by finance");
    expect(dashboard.cells.A9).toBe("Docs");
    expect(dashboard.merges[0].range).toEqual({
      start: { row: 0, column: 0 },
      end: { row: 0, column: 3 }
    });
    expect(Object.keys(dense.cells)).toHaveLength(2_000);
    expect([dense.cells.A1, dense.cells.H250]).toEqual([1, 1]);
    expect(dense.rowCount).toBe(250);
  });

  it("imports the deterministic sales matrix exactly and removes filter-derived hidden rows", async () => {
    const expected = JSON.parse(await readFile(
      resolve(fixtureDirectory, "generated-sales-structured-table.expected.json"),
      "utf8"
    )) as {
      sheetName: string;
      tableName: string;
      range: string;
      headerRow: boolean;
      totalsRow: boolean;
      columns: string[];
      bodyRowCount: number;
      keyColumn: string;
      calculatedColumn: string;
      filteredOutOrderIds: string[];
    };
    const workbook = await importWorkbookFromXlsx(
      await fixture("generated-sales-structured-table.xlsx"),
      { tableKeys: { SalesTable: { columnName: expected.keyColumn } } }
    );
    const table = workbook.tables[0];
    const sheet = workbook.sheets[0];

    expect(sheet.name).toBe(expected.sheetName);
    expect(table.name).toBe(expected.tableName);
    expect(table.range).toEqual({ start: { row: 0, column: 0 }, end: { row: 5, column: 5 } });
    expect(table.headerRow).toBe(expected.headerRow);
    expect(table.totalsRow).toBe(expected.totalsRow);
    expect(table.columns.map((column) => column.name)).toEqual(expected.columns);
    expect(table.rowIds).toHaveLength(expected.bodyRowCount);
    expect(table.keyColumnId).toBe(table.columns.find((column) => column.name === expected.keyColumn)?.id);
    expect(table.columns.find((column) => column.name === expected.calculatedColumn)?.calculatedFormula).toBe("=D2*E2");
    expect(["F2", "F3", "F4", "F5"].map((address) => sheet.cells[address])).toEqual([
      "=D2*E2", "=D3*E3", "=D4*E4", "=D5*E5"
    ]);
    expect(sheet.cells.B2).toBe(46037);
    expect(sheet.formats.B2).toMatchObject({ numberFormat: "date" });
    expect(sheet.hiddenRows?.["3"]).toBeUndefined();

    const regionColumn = table.columns.find((column) => column.name === "Region")!;
    const allowed = table.filter?.kind === "set"
      ? new Set(table.filter.values.map((value) => value.type === "string" ? value.value : ""))
      : new Set<string>();
    const filteredOut = table.rowIds.flatMap((_rowId, index) => {
      const row = table.range.start.row + 1 + index;
      const region = String(sheet.cells[formatCellAddress({ row, column: regionColumn.sheetColumn })] ?? "");
      return allowed.has(region) ? [] : [String(sheet.cells[formatCellAddress({ row, column: 0 })])];
    });
    expect(filteredOut).toEqual(expected.filteredOutOrderIds);
  });

  it("round-trips the Microsoft Excel issue-1669 tables, styles, headers, and filters", async () => {
    const source = await fixture("exceljs-issue-1669.xlsx");
    const first = await importWorkbookFromXlsx(source);

    expect(first.tables.map(summarizeTable)).toEqual([
      {
        name: "Table1",
        range: { start: { row: 0, column: 0 }, end: { row: 5, column: 1 } },
        headerRow: true,
        totalsRow: false,
        columns: ["Column1", "Column2"],
        style: {
          theme: "TableStyleMedium2",
          showFirstColumn: false,
          showLastColumn: false,
          showRowStripes: true,
          showColumnStripes: false
        },
        filter: {
          kind: "comparison",
          columnId: "Column1",
          operator: "neq",
          value: { type: "number", value: 4 }
        }
      },
      {
        name: "Table2",
        range: { start: { row: 0, column: 0 }, end: { row: 5, column: 1 } },
        headerRow: true,
        totalsRow: false,
        columns: ["DK", "Name"],
        style: {
          theme: "TableStyleLight9",
          showFirstColumn: false,
          showLastColumn: false,
          showRowStripes: true,
          showColumnStripes: false
        },
        filter: {
          kind: "set",
          columnId: "DK",
          operator: "in",
          values: [
            { type: "string", value: "T123456789" },
            { type: "string", value: "T123456791" },
            { type: "string", value: "T123456793" }
          ]
        }
      }
    ]);

    const exported = new Uint8Array(await exportWorkbookToXlsx(first));
    const second = await importWorkbookFromXlsx(exported);
    expect(second.tables.map(summarizeTable)).toEqual(first.tables.map(summarizeTable));
    const xml = tableXmlStrings(exported).join("\n");
    expect(xml).toContain('name="TableStyleMedium2"');
    expect(xml).toContain('operator="notEqual" val="4"');
    expect(xml).toContain('name="TableStyleLight9"');
    expect(xml).toContain('<filter val="T123456789"/>');
  });

  it("does not persist generated workbook identities into OOXML", async () => {
    const imported = await importWorkbookFromXlsx(await fixture("generated-sales-structured-table.xlsx"), {
      tableKeys: { SalesTable: { columnName: "Order ID" } }
    });
    const exported = new Uint8Array(await exportWorkbookToXlsx(imported));
    const xml = tableXmlStrings(exported).join("\n");
    const identities = imported.tables.flatMap((table) => [
      table.id,
      table.keyColumnId,
      ...table.columns.map((column) => column.id),
      ...table.rowIds
    ]).filter((value): value is string => typeof value === "string");
    for (const identity of identities) expect(xml).not.toContain(identity);
  });
});
