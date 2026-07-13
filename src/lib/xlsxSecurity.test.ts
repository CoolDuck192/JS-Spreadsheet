import { DOMParser } from "@xmldom/xmldom";
import { zipSync } from "fflate";
import { describe, expect, it, vi } from "vitest";

import {
  makeCachedFormulaWorksheetPackage,
  makeDenseWorksheetPackage,
  makeLargeDenseWorksheetPackage,
  makeMultipleDenseWorksheetsPackage
} from "../test/xlsxSecurityFixtures";

import {
  validateXlsxArchive,
  type XlsxSecurityLimits
} from "./xlsxSecurity";

const TABLE_RELATIONSHIP_TYPE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table";
const TABLE_CONTENT_TYPE =
  "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml";
const WORKSHEET_CONTENT_TYPE =
  "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";
const SHARED_STRINGS_CONTENT_TYPE =
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml";

const encoder = new TextEncoder();

type PackageOptions = {
  contentTypesXml?: string;
  relationshipXml?: string | null;
  tableXml?: string | null;
  worksheetXml?: string;
  extraEntries?: Readonly<Record<string, Uint8Array | string>>;
};

function contentTypes(
  contentType = TABLE_CONTENT_TYPE,
  extraOverrides: ReadonlyArray<readonly [partName: string, contentType: string]> = []
): string {
  return `<?xml version="1.0" encoding="UTF-8"?>
    <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
      <Override PartName="/xl/tables/table1.xml" ContentType="${contentType}"/>
      ${extraOverrides.map(([partName, type]) =>
        `<Override PartName="${partName}" ContentType="${type}"/>`
      ).join("")}
    </Types>`;
}

function worksheetContentTypes(): string {
  return contentTypes(TABLE_CONTENT_TYPE, [[
    "/xl/worksheets/sheet1.xml",
    WORKSHEET_CONTENT_TYPE
  ]]);
}

function sharedStringsContentTypes(): string {
  return contentTypes(TABLE_CONTENT_TYPE, [[
    "/xl/sharedStrings.xml",
    SHARED_STRINGS_CONTENT_TYPE
  ]]);
}

function relationships(
  target = "../tables/table1.xml",
  targetMode = ""
): string {
  return `<?xml version="1.0" encoding="UTF-8"?>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1" Type="${TABLE_RELATIONSHIP_TYPE}" Target="${target}"${targetMode}/>
    </Relationships>`;
}

function makePackage(options: PackageOptions = {}): Uint8Array {
  const entries: Record<string, Uint8Array> = {
    "[Content_Types].xml": encoder.encode(
      options.contentTypesXml ?? contentTypes()
    ),
    "xl/worksheets/sheet1.xml": encoder.encode(
      options.worksheetXml ??
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>'
    )
  };

  const relationshipXml =
    options.relationshipXml === undefined
      ? relationships()
      : options.relationshipXml;
  if (relationshipXml !== null) {
    entries["xl/worksheets/_rels/sheet1.xml.rels"] =
      encoder.encode(relationshipXml);
  }

  const tableXml =
    options.tableXml === undefined
      ? '<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Table1" displayName="Table1" ref="A1:B2"/>'
      : options.tableXml;
  if (tableXml !== null) {
    entries["xl/tables/table1.xml"] = encoder.encode(tableXml);
  }

  for (const [name, value] of Object.entries(options.extraEntries ?? {})) {
    entries[name] = typeof value === "string" ? encoder.encode(value) : value;
  }

  return zipSync(entries, { level: 0 });
}

function makePackageWithGenericXmlBeforeGrantedWorksheets(): Uint8Array {
  const entries: Record<string, Uint8Array> = {
    "[Content_Types].xml": encoder.encode(contentTypes(TABLE_CONTENT_TYPE, [
      ["/xl/worksheets/sheet1.xml", WORKSHEET_CONTENT_TYPE],
      ["/xl/worksheets/sheet2.xml", WORKSHEET_CONTENT_TYPE]
    ]))
  };
  const genericXml = encoder.encode(`<root>${"<value/>".repeat(900_000)}</root>`);
  for (let index = 1; index <= 9; index += 1) {
    entries[`custom/before-${index}.xml`] = genericXml;
  }
  const worksheet = encoder.encode(
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      '<dimension ref="A1:T100000"/><sheetData/></worksheet>'
  );
  entries["xl/worksheets/sheet1.xml"] = worksheet;
  entries["xl/worksheets/sheet2.xml"] = worksheet;
  return zipSync(entries, { level: 0 });
}

let overFloorRows: string | undefined;
let overFloorSharedStringItems: string | undefined;
let richSharedStringItems: string | undefined;

function worksheetRowsAboveFixedElementLimit(): string {
  overFloorRows ??= `<row>${"<c><v/></c>".repeat(10)}</row>`.repeat(47_620);
  return overFloorRows;
}

function sharedStringItemsAboveFixedElementLimit(): string {
  overFloorSharedStringItems ??= "<si><r><t>x</t></r></si>".repeat(333_334);
  return overFloorSharedStringItems;
}

function twoRunRichSharedStringItems(): string {
  richSharedStringItems ??=
    "<si><r><t>x</t></r><r><t>y</t></r></si>".repeat(300_000);
  return richSharedStringItems;
}

function worksheetWithDimension(
  beforeSheetData: string,
  afterSheetData = "",
  rootAttributes = ""
): string {
  return `<worksheet${rootAttributes}>${beforeSheetData}<sheetData>${worksheetRowsAboveFixedElementLimit()}</sheetData>${afterSheetData}</worksheet>`;
}

function readU16(bytes: Uint8Array, offset: number): number {
  return bytes[offset] | (bytes[offset + 1] << 8);
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

function findEocd(bytes: Uint8Array): number {
  for (let offset = bytes.length - 22; offset >= 0; offset -= 1) {
    if (readU32(bytes, offset) === 0x06054b50) return offset;
  }
  throw new Error("EOCD not found in test archive");
}

type CentralRecord = {
  centralOffset: number;
  localOffset: number;
  name: string;
};

function centralRecords(bytes: Uint8Array): CentralRecord[] {
  const eocd = findEocd(bytes);
  const count = readU16(bytes, eocd + 10);
  let offset = readU32(bytes, eocd + 16);
  const records: CentralRecord[] = [];

  for (let index = 0; index < count; index += 1) {
    if (readU32(bytes, offset) !== 0x02014b50) {
      throw new Error("bad central record in test archive");
    }
    const nameLength = readU16(bytes, offset + 28);
    const extraLength = readU16(bytes, offset + 30);
    const commentLength = readU16(bytes, offset + 32);
    const nameBytes = bytes.subarray(offset + 46, offset + 46 + nameLength);
    records.push({
      centralOffset: offset,
      localOffset: readU32(bytes, offset + 42),
      name: new TextDecoder().decode(nameBytes)
    });
    offset += 46 + nameLength + extraLength + commentLength;
  }

  return records;
}

function recordNamed(bytes: Uint8Array, name: string): CentralRecord {
  const record = centralRecords(bytes).find((candidate) => candidate.name === name);
  if (!record) throw new Error(`missing test ZIP record: ${name}`);
  return record;
}

function forgeDeclaredUncompressedSize(
  bytes: Uint8Array,
  record: CentralRecord,
  size: number,
  updateLocal = true
): void {
  writeU32(bytes, record.centralOffset + 24, size);
  if (updateLocal) writeU32(bytes, record.localOffset + 22, size);
}

type SecurityIssueCode =
  | "XLSX_ARCHIVE_LIMIT"
  | "XLSX_SHEET_TOO_LARGE"
  | "XLSX_XML_UNSAFE"
  | "XLSX_RELATIONSHIP_INVALID";

function expectRejectedBeforeLoad(
  bytes: Uint8Array,
  code: SecurityIssueCode,
  limits?: Partial<XlsxSecurityLimits>
): void {
  const excelJsLoader = vi.fn((_bytes: Uint8Array) => ({ workbook: "partial" }));
  const validation = validateXlsxArchive(bytes, limits);
  const workbook = validation.ok ? excelJsLoader(bytes) : undefined;

  expect(validation).toMatchObject({ ok: false, issue: { code } });
  expect(workbook).toBeUndefined();
  expect(excelJsLoader).not.toHaveBeenCalled();
}

describe("validateXlsxArchive ZIP preflight", () => {
  it("accepts a minimal native-table package", () => {
    expect(validateXlsxArchive(makePackage())).toEqual({ ok: true });
  });

  it("rejects input over the default 64 MiB limit before parsing", () => {
    expectRejectedBeforeLoad(
      new Uint8Array(64 * 1024 * 1024 + 1),
      "XLSX_ARCHIVE_LIMIT"
    );
  });

  it("rejects more than 4,096 central-directory entries", () => {
    const entries: Record<string, Uint8Array> = {};
    for (let index = 0; index < 4_097; index += 1) {
      entries[`entry-${String(index).padStart(4, "0")}.bin`] = new Uint8Array();
    }
    expectRejectedBeforeLoad(zipSync(entries), "XLSX_ARCHIVE_LIMIT");
  });

  it("rejects a forged entry declared over 32 MiB without allocating its output", () => {
    const bytes = makePackage();
    forgeDeclaredUncompressedSize(
      bytes,
      recordNamed(bytes, "xl/tables/table1.xml"),
      32 * 1024 * 1024 + 1
    );
    expectRejectedBeforeLoad(bytes, "XLSX_ARCHIVE_LIMIT");
  });

  it("rejects a declared total over 256 MiB", () => {
    const entries: Record<string, Uint8Array> = {};
    for (let index = 0; index < 9; index += 1) {
      entries[`large-${index}.bin`] = new Uint8Array();
    }
    const bytes = zipSync(entries, { level: 0 });
    for (const record of centralRecords(bytes)) {
      forgeDeclaredUncompressedSize(bytes, record, 30 * 1024 * 1024);
    }
    expectRejectedBeforeLoad(bytes, "XLSX_ARCHIVE_LIMIT");
  });

  it("rejects a highly compressed repeated-data ZIP bomb over 100:1", () => {
    const bytes = zipSync(
      { "repeated.bin": new Uint8Array(1024 * 1024) },
      { level: 9 }
    );
    expectRejectedBeforeLoad(bytes, "XLSX_ARCHIVE_LIMIT");
  });

  it("rejects encrypted entries", () => {
    const bytes = makePackage();
    const record = recordNamed(bytes, "xl/tables/table1.xml");
    writeU16(
      bytes,
      record.centralOffset + 8,
      readU16(bytes, record.centralOffset + 8) | 0x0001
    );
    writeU16(
      bytes,
      record.localOffset + 6,
      readU16(bytes, record.localOffset + 6) | 0x0001
    );
    expectRejectedBeforeLoad(bytes, "XLSX_ARCHIVE_LIMIT");
  });

  it("rejects duplicate entry names", () => {
    const bytes = zipSync({
      "a.xml": encoder.encode("a"),
      "b.xml": encoder.encode("b")
    });
    const [first, second] = centralRecords(bytes);
    const firstName = bytes.subarray(first.centralOffset + 46, first.centralOffset + 51);
    bytes.set(firstName, second.centralOffset + 46);
    bytes.set(firstName, second.localOffset + 30);
    expectRejectedBeforeLoad(bytes, "XLSX_ARCHIVE_LIMIT");
  });

  it("rejects traversal and absolute entry names", () => {
    expectRejectedBeforeLoad(
      zipSync({ "../escape.xml": encoder.encode("x") }),
      "XLSX_ARCHIVE_LIMIT"
    );
    expectRejectedBeforeLoad(
      zipSync({ "/absolute.xml": encoder.encode("x") }),
      "XLSX_ARCHIVE_LIMIT"
    );
  });

  it("rejects ZIP64 sentinels", () => {
    const bytes = makePackage();
    const eocd = findEocd(bytes);
    writeU16(bytes, eocd + 10, 0xffff);
    expectRejectedBeforeLoad(bytes, "XLSX_ARCHIVE_LIMIT");
  });

  it("rejects overlapping local entry offsets", () => {
    const bytes = makePackage();
    const records = centralRecords(bytes);
    writeU32(bytes, records[1].centralOffset + 42, records[0].localOffset);
    expectRejectedBeforeLoad(bytes, "XLSX_ARCHIVE_LIMIT");
  });

  it("rejects inconsistent forged central-directory sizes before inflation", () => {
    const bytes = makePackage();
    forgeDeclaredUncompressedSize(
      bytes,
      recordNamed(bytes, "xl/tables/table1.xml"),
      1024,
      false
    );
    expectRejectedBeforeLoad(bytes, "XLSX_ARCHIVE_LIMIT");
  });

  it("rejects a table XML entry over the configured 4 MiB boundary", () => {
    const bytes = makePackage();
    forgeDeclaredUncompressedSize(
      bytes,
      recordNamed(bytes, "xl/tables/table1.xml"),
      4 * 1024 * 1024 + 1
    );
    expectRejectedBeforeLoad(bytes, "XLSX_ARCHIVE_LIMIT");
  });
});

describe("validateXlsxArchive bounded XML validation", () => {
  it("accepts a dense A1:J100000 worksheet at the import security seam", () => {
    expect(validateXlsxArchive(makeDenseWorksheetPackage())).toEqual({ ok: true });
  });

  it(
    "accepts cached formula values across a dense A1:P100000 worksheet",
    { timeout: 30_000 },
    () => {
      expect(validateXlsxArchive(makeCachedFormulaWorksheetPackage())).toEqual({ ok: true });
    }
  );

  it(
    "scales package totals across three valid dense A1:P100000 worksheets",
    { timeout: 30_000 },
    () => {
      expect(validateXlsxArchive(makeMultipleDenseWorksheetsPackage())).toEqual({ ok: true });
    }
  );

  it(
    "applies validated worksheet grants independently of ZIP entry order",
    { timeout: 30_000 },
    () => {
      expect(validateXlsxArchive(makePackageWithGenericXmlBeforeGrantedWorksheets()))
        .toEqual({ ok: true });
    }
  );

  it("allows a dense worksheet entry above 32 MiB by default but honors an explicit low byte limit", () => {
    const data = makeLargeDenseWorksheetPackage();

    expect(validateXlsxArchive(data)).toEqual({ ok: true });
    expectRejectedBeforeLoad(data, "XLSX_ARCHIVE_LIMIT", {
      maxEntryBytes: 32 * 1024 * 1024
    });
  });

  it.each([
    ["missing", null],
    ["wrong", "application/xml"]
  ])("does not grant worksheet node scaling for a %s content type", (
    _label,
    worksheetContentType
  ) => {
    expectRejectedBeforeLoad(
      makeDenseWorksheetPackage(worksheetContentType),
      "XLSX_XML_UNSAFE"
    );
  });

  it.each([
    ["missing", null],
    ["wrong", "application/xml"]
  ])("does not grant the worksheet 64 MiB limit for a %s content type", (
    _label,
    worksheetContentType
  ) => {
    expectRejectedBeforeLoad(
      makeLargeDenseWorksheetPackage(worksheetContentType),
      "XLSX_ARCHIVE_LIMIT"
    );
  });

  it("keeps explicit low XML limits authoritative over a valid worksheet dimension", () => {
    expectRejectedBeforeLoad(
      makeDenseWorksheetPackage(),
      "XLSX_XML_UNSAFE",
      { maxXmlElements: 1_000_000 }
    );
    expectRejectedBeforeLoad(
      makePackage({
        contentTypesXml: worksheetContentTypes(),
        worksheetXml:
          '<worksheet><dimension ref="A1:J100000"/><sheetData>' +
          '<row r="1"><c r="A1" t="n"/></row></sheetData></worksheet>'
      }),
      "XLSX_XML_UNSAFE",
      { maxXmlAttributes: 3 }
    );
  });

  it.each([
    ["missing", "", "", ""],
    ["invalid", '<dimension ref="A0:J47620"/>', "", ""],
    [
      "duplicate",
      '<dimension ref="A1:J47620"/><dimension ref="A1:J47620"/>',
      "",
      ""
    ],
    [
      "namespace-prefixed",
      '<evil:dimension ref="A1:J47620"/>',
      "",
      ' xmlns:evil="urn:example:evil"'
    ],
    ["out-of-bounds", '<dimension ref="A1:XFE47620"/>', "", ""],
    [
      "unsafe-integer",
      '<dimension ref="A1:J9007199254740992"/>',
      "",
      ""
    ]
  ])("does not scale the parser for a %s worksheet dimension", (
    _label,
    beforeSheetData,
    afterSheetData,
    rootAttributes
  ) => {
    expectRejectedBeforeLoad(
      makePackage({
        contentTypesXml: worksheetContentTypes(),
        worksheetXml: worksheetWithDimension(
          beforeSheetData,
          afterSheetData,
          rootAttributes
        )
      }),
      "XLSX_XML_UNSAFE"
    );
  });

  it("does not scale the parser for a dimension after sheetData", () => {
    expectRejectedBeforeLoad(
      makePackage({
        contentTypesXml: worksheetContentTypes(),
        worksheetXml: worksheetWithDimension(
          "",
          '<dimension ref="A1:J47620"/>'
        )
      }),
      "XLSX_XML_UNSAFE"
    );
  });

  it("rejects a worksheet dimension beyond the supported dense envelope clearly", () => {
    const validation = validateXlsxArchive(makePackage({
      contentTypesXml: worksheetContentTypes(),
      worksheetXml:
        '<worksheet><dimension ref="A1:XFD1048576"/><sheetData/></worksheet>'
    }));

    expect(validation).toMatchObject({
      ok: false,
      issue: {
        code: "XLSX_SHEET_TOO_LARGE",
        message: expect.stringMatching(/worksheet.*too large/i)
      }
    });
  });

  it("scales sharedStrings from safe root counts above the fixed element limit", () => {
    expect(validateXlsxArchive(makePackage({
      contentTypesXml: sharedStringsContentTypes(),
      extraEntries: {
        "xl/sharedStrings.xml":
          '<sst count="333334" uniqueCount="333334">' +
          sharedStringItemsAboveFixedElementLimit() +
          "</sst>"
      }
    }))).toEqual({ ok: true });
  });

  it("allows two rich-text runs for every declared shared string", () => {
    expect(validateXlsxArchive(makePackage({
      contentTypesXml: sharedStringsContentTypes(),
      extraEntries: {
        "xl/sharedStrings.xml":
          '<sst count="300000" uniqueCount="300000">' +
          twoRunRichSharedStringItems() +
          "</sst>"
      }
    }))).toEqual({ ok: true });
  });

  it.each([
    ["negative", 'count="-1" uniqueCount="-1"'],
    ["unsafe integer", 'count="9007199254740992" uniqueCount="333334"'],
    ["unique count above count", 'count="333333" uniqueCount="333334"'],
    [
      "namespace-prefixed",
      'xmlns:evil="urn:example:evil" evil:count="333334" evil:uniqueCount="333334"'
    ]
  ])("does not scale sharedStrings for %s root counts", (_label, attributes) => {
    expectRejectedBeforeLoad(
      makePackage({
        contentTypesXml: sharedStringsContentTypes(),
        extraEntries: {
          "xl/sharedStrings.xml":
            `<sst ${attributes}>` +
            sharedStringItemsAboveFixedElementLimit() +
            "</sst>"
        }
      }),
      "XLSX_XML_UNSAFE"
    );
  });

  it.each([
    ["missing", null],
    ["wrong", "application/xml"]
  ])("does not grant sharedStrings node scaling for a %s content type", (
    _label,
    sharedStringsContentType
  ) => {
    expectRejectedBeforeLoad(
      makePackage({
        contentTypesXml: contentTypes(
          TABLE_CONTENT_TYPE,
          sharedStringsContentType === null
            ? []
            : [["/xl/sharedStrings.xml", sharedStringsContentType]]
        ),
        extraEntries: {
          "xl/sharedStrings.xml":
            '<sst count="333334" uniqueCount="333334">' +
            sharedStringItemsAboveFixedElementLimit() +
            "</sst>"
        }
      }),
      "XLSX_XML_UNSAFE"
    );
  });

  it("allows sharedStrings above 32 MiB under its 64 MiB default", () => {
    expect(validateXlsxArchive(makePackage({
      contentTypesXml: sharedStringsContentTypes(),
      extraEntries: {
        "xl/sharedStrings.xml":
          '<sst count="1" uniqueCount="1"><si><t>' +
          "x".repeat(32 * 1024 * 1024) +
          "</t></si></sst>"
      }
    }))).toEqual({ ok: true });
  });

  it.each([
    ["missing", null],
    ["wrong", "application/xml"]
  ])("does not grant the sharedStrings 64 MiB limit for a %s content type", (
    _label,
    sharedStringsContentType
  ) => {
    expectRejectedBeforeLoad(
      makePackage({
        contentTypesXml: contentTypes(
          TABLE_CONTENT_TYPE,
          sharedStringsContentType === null
            ? []
            : [["/xl/sharedStrings.xml", sharedStringsContentType]]
        ),
        extraEntries: {
          "xl/sharedStrings.xml":
            '<sst count="1" uniqueCount="1"><si><t>' +
            "x".repeat(32 * 1024 * 1024) +
            "</t></si></sst>"
        }
      }),
      "XLSX_ARCHIVE_LIMIT"
    );
  });

  it("does not scale table XML from its body ref", () => {
    expectRejectedBeforeLoad(
      makePackage({
        tableXml:
          '<table name="Table1" ref="A1:J100000">' +
          "<a/>".repeat(1_000_000) +
          "</table>"
      }),
      "XLSX_XML_UNSAFE"
    );
  });

  it("accepts a legitimate worksheet above the former global element budget", () => {
    const cells = "<c/>".repeat(120_000);

    expect(validateXlsxArchive(makePackage({
      worksheetXml: `<worksheet><sheetData>${cells}</sheetData></worksheet>`
    }))).toEqual({ ok: true });
  });

  it("applies element budgets per XML part", () => {
    expect(validateXlsxArchive(makePackage({
      extraEntries: {
        "custom/one.xml": "<one><value/></one>",
        "custom/two.xml": "<two><value/></two>"
      }
    }), {
      maxXmlElements: 2,
      maxXmlAttributes: 10
    })).toEqual({ ok: true });
  });

  it("retains an aggregate XML element budget across parts", () => {
    expectRejectedBeforeLoad(
      makePackage({
        extraEntries: {
          "custom/one.xml": "<one><value/></one>",
          "custom/two.xml": "<two><value/></two>"
        }
      }),
      "XLSX_XML_UNSAFE",
      { maxTotalXmlElements: 8 }
    );
  });

  it("retains an aggregate XML attribute budget across parts", () => {
    expectRejectedBeforeLoad(
      makePackage(),
      "XLSX_XML_UNSAFE",
      { maxTotalXmlAttributes: 12 }
    );
  });

  it.each(["DOCTYPE", "doctype"])(
    "rejects a case-insensitive %s declaration",
    (keyword) => {
      expectRejectedBeforeLoad(
        makePackage({
          tableXml: `<!${keyword} table SYSTEM "file:///etc/passwd"><table/>`
        }),
        "XLSX_XML_UNSAFE"
      );
    }
  );

  it("rejects ENTITY declarations", () => {
    expectRejectedBeforeLoad(
      makePackage({
        tableXml: '<!DOCTYPE table [<!ENTITY xxe SYSTEM "file:///etc/passwd">]><table>&xxe;</table>'
      }),
      "XLSX_XML_UNSAFE"
    );
  });

  it("rejects malformed XML", () => {
    expectRejectedBeforeLoad(
      makePackage({ tableXml: "<table><broken></table>" }),
      "XLSX_XML_UNSAFE"
    );
  });

  it("rejects excessive XML depth", () => {
    expectRejectedBeforeLoad(
      makePackage({ tableXml: "<table><a><b/></a></table>" }),
      "XLSX_XML_UNSAFE",
      { maxXmlDepth: 2 }
    );
  });

  it("rejects excessive XML element count", () => {
    expectRejectedBeforeLoad(
      makePackage({ tableXml: "<table><a/><b/></table>" }),
      "XLSX_XML_UNSAFE",
      { maxXmlElements: 2 }
    );
  });

  it("rejects excessive XML attribute count", () => {
    expectRejectedBeforeLoad(
      makePackage({ tableXml: '<table first="1" second="2"/>' }),
      "XLSX_XML_UNSAFE",
      { maxXmlAttributes: 1 }
    );
  });

  it("rejects an over-budget part before materializing it as a DOM", () => {
    const worksheetXml = "<worksheet><sheetData><c/><c/></sheetData></worksheet>";
    const parseSpy = vi.spyOn(DOMParser.prototype, "parseFromString");

    try {
      expectRejectedBeforeLoad(
        makePackage({ worksheetXml }),
        "XLSX_XML_UNSAFE",
        { maxXmlElements: 3 }
      );
      expect(parseSpy.mock.calls.some(([xml]) => xml === worksheetXml)).toBe(false);
    } finally {
      parseSpy.mockRestore();
    }
  });

  it("does not materialize ordinary XML parts as retained DOMs", () => {
    const worksheetXml = '<worksheet xmlns="urn:test"><sheetData/></worksheet>';
    const tableXml = '<table xmlns="urn:test" id="1"/>';
    const parseSpy = vi.spyOn(DOMParser.prototype, "parseFromString");

    try {
      expect(validateXlsxArchive(makePackage({ worksheetXml, tableXml }))).toEqual({ ok: true });
      const parsedInputs = parseSpy.mock.calls.map(([xml]) => xml);
      expect(parsedInputs).not.toContain(worksheetXml);
      expect(parsedInputs).not.toContain(tableXml);
    } finally {
      parseSpy.mockRestore();
    }
  });

  it("applies the table XML byte limit before DOM parsing", () => {
    expectRejectedBeforeLoad(
      makePackage({ tableXml: "<table/>" }),
      "XLSX_ARCHIVE_LIMIT",
      { maxTableXmlBytes: 7 }
    );
  });
});

describe("validateXlsxArchive OOXML table relationships", () => {
  it("rejects namespace-prefixed lookalikes for required relationship attributes", () => {
    const prefixedOnly = `
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"
        xmlns:evil="urn:example:evil">
        <Relationship evil:Id="rId1" evil:Type="${TABLE_RELATIONSHIP_TYPE}"
          evil:Target="../tables/table1.xml"/>
      </Relationships>`;
    expectRejectedBeforeLoad(
      makePackage({ relationshipXml: prefixedOnly }),
      "XLSX_RELATIONSHIP_INVALID"
    );
  });

  it("uses unqualified relationship attributes when prefixed lookalikes are also present", () => {
    const mixedAttributes = `
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"
        xmlns:evil="urn:example:evil">
        <Relationship evil:Id="spoof" Id="rId1" Type="${TABLE_RELATIONSHIP_TYPE}"
          evil:Target="../../../escape.xml" Target="../tables/table1.xml"/>
      </Relationships>`;
    expect(validateXlsxArchive(makePackage({ relationshipXml: mixedAttributes }))).toEqual({ ok: true });
  });

  it("rejects duplicate relationship IDs", () => {
    const duplicateIds = `
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="${TABLE_RELATIONSHIP_TYPE}" Target="../tables/table1.xml"/>
        <Relationship Id="rId1" Type="urn:example:other" Target="other.xml"/>
      </Relationships>`;
    expectRejectedBeforeLoad(
      makePackage({ relationshipXml: duplicateIds }),
      "XLSX_RELATIONSHIP_INVALID"
    );
  });

  it("rejects duplicate IDs in non-worksheet relationship parts too", () => {
    const duplicateIds = `
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="urn:example:first" Target="first.xml"/>
        <Relationship Id="rId1" Type="urn:example:second" Target="second.xml"/>
      </Relationships>`;
    expectRejectedBeforeLoad(
      makePackage({
        extraEntries: { "xl/_rels/workbook.xml.rels": duplicateIds }
      }),
      "XLSX_RELATIONSHIP_INVALID"
    );
  });

  it("rejects external table relationships", () => {
    expectRejectedBeforeLoad(
      makePackage({
        relationshipXml: relationships(
          "https://example.invalid/table.xml",
          ' TargetMode="External"'
        )
      }),
      "XLSX_RELATIONSHIP_INVALID"
    );
  });

  it.each([
    "../../../escape.xml",
    "..%2F..%2Fescape.xml",
    "../embeddings/table1.xml"
  ])("rejects a table target outside xl/tables/: %s", (target) => {
    expectRejectedBeforeLoad(
      makePackage({ relationshipXml: relationships(target) }),
      "XLSX_RELATIONSHIP_INVALID"
    );
  });

  it("rejects a missing table target", () => {
    expectRejectedBeforeLoad(
      makePackage({ tableXml: null }),
      "XLSX_RELATIONSHIP_INVALID"
    );
  });

  it("rejects a table target with a missing content-type declaration", () => {
    expectRejectedBeforeLoad(
      makePackage({
        contentTypesXml:
          '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>'
      }),
      "XLSX_RELATIONSHIP_INVALID"
    );
  });

  it("rejects a table target with the wrong content type", () => {
    expectRejectedBeforeLoad(
      makePackage({ contentTypesXml: contentTypes("application/xml") }),
      "XLSX_RELATIONSHIP_INVALID"
    );
  });

  it("rejects a worksheet relationship part whose source worksheet is missing", () => {
    const archive = zipSync(
      {
        "[Content_Types].xml": encoder.encode(contentTypes()),
        "xl/worksheets/_rels/sheet1.xml.rels": encoder.encode(relationships()),
        "xl/tables/table1.xml": encoder.encode("<table/>")
      },
      { level: 0 }
    );
    expectRejectedBeforeLoad(archive, "XLSX_RELATIONSHIP_INVALID");
  });
});
