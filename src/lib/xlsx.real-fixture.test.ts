import { readFile } from "node:fs/promises";
import { resolve } from "node:path";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { describe, expect, it } from "vitest";
import type { FilterExpression } from "../table/core/query";
import type { StructuredTable, WorkbookModel } from "../types";
import { migrateWorkbookModel } from "../core/workbook/migrateWorkbook";
import {
  reduceStructuredTableCommand,
  type StructuredTableCommandServices
} from "../core/workbook/structuredTables";
import { formatCellAddress } from "./addressing";
import { createBlankWorkbook, setCellComment, setCellContent } from "./workbook";
import { exportWorkbookToXlsx, importWorkbookFromXlsx } from "./xlsx";
import { validateXlsxArchive } from "./xlsxSecurity";

const fixtureDirectory = resolve("src/test/fixtures/xlsx");
const MAX_XLSX_WORKSHEET_ID = 100_000;
const PATHOLOGICAL_EXCELJS_WORKSHEET_ID = 0xffff_fffe;
const tableCommandServices: StructuredTableCommandServices = {
  createId: () => "unused-id",
  getCellEvaluation: () => null
};

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

function uppercaseFirstWorksheetPart(bytes: Uint8Array): Uint8Array {
  const entries = unzipSync(bytes);
  entries["xl/worksheets/Sheet1.xml"] = entries["xl/worksheets/sheet1.xml"];
  delete entries["xl/worksheets/sheet1.xml"];
  entries["xl/worksheets/_rels/Sheet1.xml.rels"] =
    entries["xl/worksheets/_rels/sheet1.xml.rels"];
  delete entries["xl/worksheets/_rels/sheet1.xml.rels"];
  entries["xl/_rels/workbook.xml.rels"] = strToU8(
    strFromU8(entries["xl/_rels/workbook.xml.rels"])
      .replace("worksheets/sheet1.xml", "worksheets/Sheet1.xml")
  );
  entries["[Content_Types].xml"] = strToU8(
    strFromU8(entries["[Content_Types].xml"])
      .replace("/xl/worksheets/sheet1.xml", "/xl/worksheets/Sheet1.xml")
  );
  return zipSync(entries);
}

function prefixFirstWorksheetTargetWithWhitespace(bytes: Uint8Array): Uint8Array {
  const entries = unzipSync(bytes);
  entries["xl/_rels/workbook.xml.rels"] = strToU8(
    strFromU8(entries["xl/_rels/workbook.xml.rels"]).replace(
      'Target="/xl/worksheets/sheet1.xml"',
      'Target=" /xl/worksheets/sheet1.xml"'
    )
  );
  return zipSync(entries);
}

function suffixFirstWorksheetPartWithEncodedSpace(bytes: Uint8Array): Uint8Array {
  const entries = unzipSync(bytes);
  entries["xl/worksheets/sheet1.xml%20"] = strToU8(
    strFromU8(entries["xl/worksheets/sheet1.xml"])
      .replace(/<tableParts\b[\s\S]*?<\/tableParts>/, "")
  );
  delete entries["xl/worksheets/sheet1.xml"];
  entries["xl/worksheets/_rels/sheet1.xml%20.rels"] =
    entries["xl/worksheets/_rels/sheet1.xml.rels"];
  delete entries["xl/worksheets/_rels/sheet1.xml.rels"];
  for (const name of ["xl/_rels/workbook.xml.rels", "[Content_Types].xml"]) {
    entries[name] = strToU8(strFromU8(entries[name]).replaceAll("sheet1.xml", "sheet1.xml%20"));
  }
  return zipSync(entries);
}

function rebindFirstWorksheetRelationshipPrefix(bytes: Uint8Array): Uint8Array {
  const entries = unzipSync(bytes);
  entries["xl/workbook.xml"] = strToU8(
    strFromU8(entries["xl/workbook.xml"])
      .replace("xmlns:r=", "xmlns:x=")
      .replace('r:id="rId1"', 'x:id="rId1"')
  );
  return zipSync(entries);
}

function turnFirstWorksheetPartIntoDirectoryEntry(bytes: Uint8Array): Uint8Array {
  const entries = unzipSync(bytes);
  entries["xl/worksheets/sheet1.xml/"] = strToU8(
    strFromU8(entries["xl/worksheets/sheet1.xml"])
      .replace(/<tableParts\b[\s\S]*?<\/tableParts>/, "")
  );
  delete entries["xl/worksheets/sheet1.xml"];
  delete entries["xl/worksheets/_rels/sheet1.xml.rels"];
  entries["xl/worksheets/_rels/sheet2.xml.rels"] = strToU8(
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" ' +
    'Target="/xl/comments/comment1.xml" Id="comments"/></Relationships>'
  );
  entries["xl/_rels/workbook.xml.rels"] = strToU8(
    strFromU8(entries["xl/_rels/workbook.xml.rels"])
      .replace("sheet1.xml", "sheet1.xml/")
  );
  entries["[Content_Types].xml"] = strToU8(
    strFromU8(entries["[Content_Types].xml"])
      .replace(/<Override PartName="\/xl\/worksheets\/sheet1\.xml"[^>]*\/>/, "")
  );
  return zipSync(entries);
}

function aliasLongCommentWorksheets(
  bytes: Uint8Array,
  aliases: readonly {
    name?: string;
    sourceIndex: number;
    sheetId?: number | string | null;
  }[]
): Uint8Array {
  const entries = unzipSync(bytes);
  const workbookXml = strFromU8(entries["xl/workbook.xml"]);
  const sourceSheets = workbookXml.match(/<sheet\b[^>]*\/>/g);
  expect(sourceSheets).toHaveLength(2);
  const aliasedSheets = aliases.map(({ name, sourceIndex, sheetId }, index) => {
    const sourceSheet = sheetId === null
      ? sourceSheets![sourceIndex].replace(/\ssheetId="[^"]*"/, "")
      : sourceSheets![sourceIndex]
          .replace(/sheetId="[^"]*"/, `sheetId="${sheetId ?? index + 1}"`);
    return name === undefined
      ? sourceSheet.replace(/\sname="[^"]*"/, "")
      : sourceSheet.replace(/name="[^"]*"/, `name="${name}"`);
  }).join("");
  entries["xl/workbook.xml"] = strToU8(
    workbookXml.replace(/<sheets>[\s\S]*?<\/sheets>/, `<sheets>${aliasedSheets}</sheets>`)
  );
  return zipSync(entries);
}

function nestFirstWorkbookSheetInUnknownWrapper(bytes: Uint8Array): Uint8Array {
  const entries = unzipSync(bytes);
  const workbookXml = strFromU8(entries["xl/workbook.xml"]);
  const nestedWorkbookXml = workbookXml.replace(
    /(<sheets>)(<sheet\b[^>]*\/>)/,
    "$1<wrapper>$2</wrapper>"
  );
  expect(nestedWorkbookXml).not.toBe(workbookXml);
  entries["xl/workbook.xml"] = strToU8(nestedWorkbookXml);
  return zipSync(entries);
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

  it("matches native comments by sheet order when ExcelJS double-decodes the sheet name", async () => {
    const workbook = await importWorkbookFromXlsx(
      await fixture("variant-comment-double-escaped-sheet.xlsx")
    );

    expect(workbook.sheets[0].name).toBe("Q&A");
    expect(workbook.sheets[0].comments.A1).toBe("hello");
  });

  it("keeps long-name comments, formulas, defined names, and table formulas coherent", async () => {
    const workbook = await importWorkbookFromXlsx(
      await fixture("variant-comment-long-sheet.xlsx")
    );

    expect(workbook.sheets.map((sheet) => sheet.name)).toEqual([
      "Commented worksheet with a very",
      "Commented worksheet with a ve 1"
    ]);
    expect(workbook.sheets[0].comments.A1).toBe("hello");
    expect(workbook.sheets[0].cells.C2).toBe(
      "='Commented worksheet with a ve 1'!$A$1"
    );
    expect(workbook.namedRanges).toEqual([{
      name: "LongSheetCell",
      sheetId: "sheet-2",
      range: { start: { row: 0, column: 0 }, end: { row: 0, column: 0 } }
    }]);
    expect(
      workbook.tables[0].columns.find((column) => column.name === "Linked")?.calculatedFormula
    ).toBe("='Commented worksheet with a ve 1'!$A$1");
  });

  it("does not assign comments from worksheet parts that ExcelJS skips", async () => {
    const source = uppercaseFirstWorksheetPart(
      await fixture("variant-comment-long-sheet.xlsx")
    );

    expect(validateXlsxArchive(source).ok).toBe(true);
    const workbook = await importWorkbookFromXlsx(source);

    expect(workbook.sheets.map((sheet) => sheet.name)).toEqual([
      "Commented worksheet with a ve 1"
    ]);
    expect(workbook.sheets[0].comments).toEqual({});
  });

  it("keeps comment indexes aligned for worksheet targets normalized by ExcelJS", async () => {
    const source = prefixFirstWorksheetTargetWithWhitespace(
      await fixture("variant-comment-long-sheet.xlsx")
    );

    expect(validateXlsxArchive(source).ok).toBe(true);
    const workbook = await importWorkbookFromXlsx(source);

    expect(workbook.sheets).toHaveLength(2);
    expect(workbook.sheets[0].comments).toEqual({ A1: "hello" });
    expect(workbook.sheets[1].comments).toEqual({});
  });

  it("keeps native comments aligned for encoded worksheet targets ExcelJS treats literally", async () => {
    const source = suffixFirstWorksheetPartWithEncodedSpace(
      await fixture("variant-comment-long-sheet.xlsx")
    );

    expect(validateXlsxArchive(source)).toEqual({ ok: true });
    const workbook = await importWorkbookFromXlsx(source);

    expect(workbook.sheets.map((sheet) => ({
      name: sheet.name,
      value: sheet.cells.A1,
      comments: sheet.comments
    }))).toEqual([
      { name: "Commented worksheet with a very", value: "Q", comments: { A1: "hello" } },
      { name: "Commented worksheet with a ve 1", value: 7, comments: {} }
    ]);
  });

  it("requires ExcelJS's literal r:id worksheet relationship prefix for comment indexes", async () => {
    const source = rebindFirstWorksheetRelationshipPrefix(
      await fixture("variant-comment-long-sheet.xlsx")
    );

    expect(validateXlsxArchive(source)).toEqual({ ok: true });
    const workbook = await importWorkbookFromXlsx(source);

    expect(workbook.sheets.map((sheet) => ({
      name: sheet.name,
      value: sheet.cells.A1,
      comments: sheet.comments
    }))).toEqual([
      { name: "Commented worksheet with a ve 1", value: 7, comments: {} }
    ]);
  });

  it("does not index worksheet-shaped ZIP directory entries that ExcelJS skips", async () => {
    const source = turnFirstWorksheetPartIntoDirectoryEntry(
      await fixture("variant-comment-long-sheet.xlsx")
    );

    expect(validateXlsxArchive(source)).toEqual({ ok: true });
    const workbook = await importWorkbookFromXlsx(source);

    expect(workbook.sheets.map((sheet) => ({
      name: sheet.name,
      value: sheet.cells.A1,
      comments: sheet.comments
    }))).toEqual([
      { name: "Commented worksheet with a ve 1", value: 7, comments: { A1: "hello" } }
    ]);
  });

  it.each([
    {
      description: "plain aliases sandwich the commented worksheet",
      aliases: [
        { name: "A", sourceIndex: 1 },
        { name: "S", sourceIndex: 0 },
        { name: "B", sourceIndex: 1 }
      ],
      expected: [
        { name: "S", value: "Q", comments: { A1: "hello" } },
        { name: "B", value: 7, comments: {} }
      ]
    },
    {
      description: "plain aliases wholly before the commented worksheet",
      aliases: [
        { name: "Plain first", sourceIndex: 1 },
        { name: "Plain last", sourceIndex: 1 },
        { name: "Commented", sourceIndex: 0 }
      ],
      expected: [
        { name: "Plain last", value: 7, comments: {} },
        { name: "Commented", value: "Q", comments: { A1: "hello" } }
      ]
    },
    {
      description: "plain aliases wholly after the commented worksheet",
      aliases: [
        { name: "Commented", sourceIndex: 0 },
        { name: "Plain first", sourceIndex: 1 },
        { name: "Plain last", sourceIndex: 1 }
      ],
      expected: [
        { name: "Commented", value: "Q", comments: { A1: "hello" } },
        { name: "Plain last", value: 7, comments: {} }
      ]
    },
    {
      description: "a nameless plain worksheet precedes the commented worksheet",
      aliases: [
        { sourceIndex: 1 },
        { name: "Commented", sourceIndex: 0 }
      ],
      expected: [
        { name: "sheet1", value: 7, comments: {} },
        { name: "Commented", value: "Q", comments: { A1: "hello" } }
      ]
    },
    {
      description: "a worksheet with a missing sheet ID is absent from ExcelJS's registry",
      aliases: [
        { name: "S", sourceIndex: 0, sheetId: 1 },
        { name: "N", sourceIndex: 1, sheetId: null }
      ],
      expected: [
        { name: "S", value: "Q", comments: { A1: "hello" } }
      ]
    },
    {
      description: "distinct worksheet parts with one ID collapse to ExcelJS's last part",
      aliases: [
        { name: "X", sourceIndex: 0, sheetId: 7 },
        { name: "Y", sourceIndex: 1, sheetId: 7 }
      ],
      expected: [
        { name: "Y", value: 7, comments: {} }
      ]
    },
    {
      description: "a radix-prefixed sheet ID occupies ExcelJS's hidden slot zero",
      aliases: [
        { name: "Dropped", sourceIndex: 0, sheetId: "0x2" },
        { name: "Survivor", sourceIndex: 1, sheetId: 2 }
      ],
      expected: [
        { name: "Survivor", value: 7, comments: {} }
      ]
    },
    {
      description: "a synthesized worksheet name avoids a real name collision",
      aliases: [
        { name: "sheet2", sourceIndex: 1, sheetId: 1 },
        { sourceIndex: 0, sheetId: 2 }
      ],
      expected: [
        { name: "sheet2", value: 7, comments: {} },
        { name: "sheet2 1", value: "Q", comments: { A1: "hello" } }
      ]
    },
    {
      description: "synthesized worksheet names detect real collisions case-insensitively",
      aliases: [
        { name: "ShEeT2", sourceIndex: 1, sheetId: 1 },
        { sourceIndex: 0, sheetId: 2 }
      ],
      expected: [
        { name: "ShEeT2", value: 7, comments: {} },
        { name: "sheet2 1", value: "Q", comments: { A1: "hello" } }
      ]
    },
    {
      description: "nameless distinct parts with one ID collapse after name synthesis",
      aliases: [
        { sourceIndex: 0, sheetId: 2 },
        { sourceIndex: 1, sheetId: 2 }
      ],
      expected: [
        { name: "sheet2 1", value: 7, comments: {} }
      ]
    },
    {
      description: "nameless aliases of one part reuse one synthesized name",
      aliases: [
        { sourceIndex: 0, sheetId: 2 },
        { sourceIndex: 0, sheetId: 2 }
      ],
      expected: [
        { name: "sheet2", value: "Q", comments: { A1: "hello" } }
      ]
    },
    {
      description: "nameless aliases use the part's final sheet ID for name synthesis",
      aliases: [
        { sourceIndex: 0, sheetId: 1 },
        { sourceIndex: 0, sheetId: 2 },
        { sourceIndex: 1, sheetId: 1 }
      ],
      expected: [
        { name: "sheet1", value: 7, comments: {} },
        { name: "sheet2", value: "Q", comments: { A1: "hello" } }
      ]
    },
    {
      description: "duplicated IDs place a plain alias group before the commented worksheet",
      aliases: [
        { name: "A", sourceIndex: 1, sheetId: 2 },
        { name: "S", sourceIndex: 0, sheetId: 1 },
        { name: "B", sourceIndex: 1, sheetId: 2 }
      ],
      expected: [
        { name: "B", value: 7, comments: {} },
        { name: "S", value: "Q", comments: { A1: "hello" } }
      ]
    },
    {
      description: "duplicated IDs keep a symmetric commented alias group ahead of the solo plain worksheet",
      aliases: [
        { name: "Commented first", sourceIndex: 0, sheetId: 1 },
        { name: "Plain", sourceIndex: 1, sheetId: 2 },
        { name: "Commented last", sourceIndex: 0, sheetId: 1 }
      ],
      expected: [
        { name: "Commented last", value: "Q", comments: { A1: "hello" } },
        { name: "Plain", value: 7, comments: {} }
      ]
    }
  ])("keeps native comments aligned when $description", async ({ aliases, expected }) => {
    const source = aliasLongCommentWorksheets(
      await fixture("variant-comment-long-sheet.xlsx"),
      aliases
    );

    expect(validateXlsxArchive(source).ok).toBe(true);
    const workbook = await importWorkbookFromXlsx(source);

    expect(workbook.sheets.map((sheet) => ({
      name: sheet.name,
      value: sheet.cells.A1,
      comments: sheet.comments
    }))).toEqual(expected);
  });

  it("imports the maximum worksheet ID with native comments aligned", async () => {
    const source = aliasLongCommentWorksheets(
      await fixture("variant-comment-long-sheet.xlsx"),
      [
        {
          name: "Commented at the worksheet ID limit",
          sourceIndex: 0,
          sheetId: MAX_XLSX_WORKSHEET_ID
        },
        { name: "Plain", sourceIndex: 1, sheetId: 1 }
      ]
    );

    const workbook = await importWorkbookFromXlsx(source);

    expect(workbook.sheets.map((sheet) => ({
      name: sheet.name,
      value: sheet.cells.A1,
      comments: sheet.comments
    }))).toEqual([
      {
        name: "Commented at the worksheet ID l",
        value: "Q",
        comments: { A1: "hello" }
      },
      { name: "Plain", value: 7, comments: {} }
    ]);
  });

  it("rejects a pathological ExcelJS worksheet array index within five seconds", async () => {
    const source = aliasLongCommentWorksheets(
      await fixture("variant-comment-long-sheet.xlsx"),
      [
        {
          name: "Pathological",
          sourceIndex: 0,
          sheetId: PATHOLOGICAL_EXCELJS_WORKSHEET_ID
        },
        { name: "Plain", sourceIndex: 1, sheetId: 1 }
      ]
    );
    const startedAt = performance.now();

    await expect(importWorkbookFromXlsx(source)).rejects.toMatchObject({
      code: "XLSX_ARCHIVE_LIMIT",
      issue: {
        code: "XLSX_ARCHIVE_LIMIT",
        message:
          "Workbook worksheet ID 4294967294 exceeds the 100000 consistency limit."
      }
    });
    expect(performance.now() - startedAt).toBeLessThan(5_000);
  });

  it("rejects a pathological worksheet ID nested in the ExcelJS sheets parser", async () => {
    const source = nestFirstWorkbookSheetInUnknownWrapper(
      aliasLongCommentWorksheets(
        await fixture("variant-comment-long-sheet.xlsx"),
        [
          {
            name: "Pathological",
            sourceIndex: 0,
            sheetId: PATHOLOGICAL_EXCELJS_WORKSHEET_ID
          },
          { name: "Plain", sourceIndex: 1, sheetId: 1 }
        ]
      )
    );
    const startedAt = performance.now();

    await expect(importWorkbookFromXlsx(source)).rejects.toMatchObject({
      code: "XLSX_ARCHIVE_LIMIT",
      issue: {
        code: "XLSX_ARCHIVE_LIMIT",
        message:
          "Workbook worksheet ID 4294967294 exceeds the 100000 consistency limit."
      }
    });
    expect(performance.now() - startedAt).toBeLessThan(5_000);
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
    const source = zipSync(entries);

    expect(validateXlsxArchive(source)).toEqual({ ok: true });
    const workbook = await importWorkbookFromXlsx(source);

    expect(workbook.sheets[0]).toMatchObject({
      rowCount: 100,
      columnCount: 26
    });
    expect(workbook.sheets[0].comments).toEqual({});
  });

  it("grows the import grid for a supported AZ200 comment-only cell", async () => {
    let source = createBlankWorkbook();
    source = setCellComment(source, source.activeSheetId, "AZ200", "orphaned note");

    const workbook = await importWorkbookFromXlsx(await exportWorkbookToXlsx(source));

    expect(workbook.sheets[0]).toMatchObject({
      rowCount: 200,
      columnCount: 52,
      comments: { AZ200: "orphaned note" }
    });
    expect(workbook.sheets[0].cells).not.toHaveProperty("AZ200");
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

  it("treats renaming an imported interop table to its current name as unchanged", async () => {
    const workbook = await importWorkbookFromXlsx(await fixture("variant-table.xlsx"));
    const table = workbook.tables[0];

    const result = reduceStructuredTableCommand(workbook, {
      type: "table.rename",
      tableId: table.id,
      name: "T1"
    }, tableCommandServices);

    expect(result.status).toBe("unchanged");
    expect(result.workbook).toBe(workbook);
  });

  it("still rejects a genuinely new strict-invalid name for an imported interop table", async () => {
    const workbook = await importWorkbookFromXlsx(await fixture("variant-table.xlsx"));
    const table = workbook.tables[0];

    const result = reduceStructuredTableCommand(workbook, {
      type: "table.rename",
      tableId: table.id,
      name: "A1"
    }, tableCommandServices);

    expect(result).toMatchObject({
      status: "rejected",
      workbook,
      issues: [{ code: "TABLE_NAME_INVALID" }]
    });
    expect(result.workbook).toBe(workbook);
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
