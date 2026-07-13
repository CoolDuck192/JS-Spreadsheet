import { inflateSync } from "fflate";
import { SaxesParser, type SaxesTagNS } from "saxes";

import type { TableIssue } from "../core/commands/types";

export type XlsxSecurityLimits = {
  maxArchiveBytes: number;
  maxEntries: number;
  maxWorksheetId: number;
  maxEntryBytes: number;
  maxTotalUncompressedBytes: number;
  maxCompressionRatio: number;
  maxTableXmlBytes: number;
  maxXmlDepth: number;
  maxXmlElements: number;
  maxXmlAttributes: number;
  maxTotalXmlElements: number;
  maxTotalXmlAttributes: number;
};

export const MAX_XLSX_WORKSHEET_ID = 100_000;

export const DEFAULT_XLSX_SECURITY_LIMITS: Readonly<XlsxSecurityLimits> = {
  maxArchiveBytes: 64 * 1024 * 1024,
  maxEntries: 4_096,
  maxWorksheetId: MAX_XLSX_WORKSHEET_ID,
  maxEntryBytes: 32 * 1024 * 1024,
  maxTotalUncompressedBytes: 256 * 1024 * 1024,
  maxCompressionRatio: 100,
  maxTableXmlBytes: 4 * 1024 * 1024,
  maxXmlDepth: 128,
  maxXmlElements: 1_000_000,
  maxXmlAttributes: 2_000_000,
  maxTotalXmlElements: 8_000_000,
  maxTotalXmlAttributes: 16_000_000
};

type SecurityIssueCode =
  | "XLSX_ARCHIVE_LIMIT"
  | "XLSX_SHEET_TOO_LARGE"
  | "XLSX_XML_UNSAFE"
  | "XLSX_RELATIONSHIP_INVALID";

type ValidationResult =
  | { ok: true }
  | { ok: false; issue: TableIssue };

export type XlsxEntryExtractionResult =
  | { ok: true; entries: ReadonlyMap<string, Uint8Array> }
  | { ok: false; issue: TableIssue };

type CentralDirectoryEntry = {
  name: string;
  flags: number;
  compressionMethod: number;
  crc32: number;
  compressedSize: number;
  uncompressedSize: number;
  localHeaderOffset: number;
  centralNameOffset: number;
  centralNameLength: number;
  dataOffset: number;
  localEndOffset: number;
};

type ParsedArchive = {
  bytes: Uint8Array;
  entries: readonly CentralDirectoryEntry[];
  entriesByName: ReadonlyMap<string, CentralDirectoryEntry>;
};

class XlsxSecurityError extends Error {
  readonly code: SecurityIssueCode;

  constructor(code: SecurityIssueCode, message: string) {
    super(message);
    this.name = "XlsxSecurityError";
    this.code = code;
  }
}

const LOCAL_FILE_HEADER_SIGNATURE = 0x04034b50;
const CENTRAL_DIRECTORY_SIGNATURE = 0x02014b50;
const END_OF_CENTRAL_DIRECTORY_SIGNATURE = 0x06054b50;
const DATA_DESCRIPTOR_SIGNATURE = 0x08074b50;
const ZIP64_EXTRA_FIELD_ID = 0x0001;
const ZIP64_U16_SENTINEL = 0xffff;
const ZIP64_U32_SENTINEL = 0xffffffff;
const DATA_DESCRIPTOR_FLAG = 0x0008;
const ENCRYPTION_FLAGS = 0x0041;
const UTF8_NAME_FLAG = 0x0800;
const LARGE_XML_ENTRY_BYTES = 64 * 1024 * 1024;
const MAX_WORKSHEET_ROWS = 1_048_576;
const MAX_WORKSHEET_COLUMNS = 16_384;
const XML_BUDGET_BASE = 1_024;
// Dimension-derived scaling is deliberately limited to two million populated
// cells per worksheet. This covers the supported dense import envelope without
// allowing a sparse full-grid dimension to mint an effectively unbounded grant.
const MAX_DENSE_WORKSHEET_CELLS = 2_000_000;

export function isSupportedXlsxWorksheetSize(
  rowCount: number,
  columnCount: number
): boolean {
  if (
    !Number.isSafeInteger(rowCount) ||
    !Number.isSafeInteger(columnCount) ||
    rowCount < 1 ||
    columnCount < 1 ||
    rowCount > MAX_WORKSHEET_ROWS ||
    columnCount > MAX_WORKSHEET_COLUMNS
  ) {
    return false;
  }

  const cells = checkedMultiply(rowCount, columnCount);
  return cells !== null && cells <= MAX_DENSE_WORKSHEET_CELLS;
}

const MAX_PART_XML_ELEMENTS =
  XML_BUDGET_BASE + MAX_WORKSHEET_ROWS + MAX_DENSE_WORKSHEET_CELLS * 3;
const MAX_PART_XML_ATTRIBUTES =
  XML_BUDGET_BASE + MAX_WORKSHEET_ROWS * 4 + MAX_DENSE_WORKSHEET_CELLS * 4;
// The 256 MiB uncompressed package ceiling can contain at most four 64 MiB
// scaled XML parts, so aggregate scaling never exceeds four dense envelopes.
const MAX_PACKAGE_XML_ELEMENTS = MAX_PART_XML_ELEMENTS * 4;
const MAX_PACKAGE_XML_ATTRIBUTES = MAX_PART_XML_ATTRIBUTES * 4;
// The validated count bounds unique stored items because uniqueCount <= count.
// Eight elements cover two rich-text runs plus one phonetic run and properties.
const SHARED_STRING_XML_ELEMENTS_PER_COUNT = 8;
const SHARED_STRING_XML_ATTRIBUTES_PER_COUNT = 4;
const SPREADSHEETML_NAMESPACES = new Set([
  "",
  "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
  "http://purl.oclc.org/ooxml/spreadsheetml/main"
]);
const WORKSHEET_CONTENT_TYPES = new Set([
  "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml",
  "application/vnd.ms-excel.worksheet+xml"
]);
const SHARED_STRINGS_CONTENT_TYPES = new Set([
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml",
  "application/vnd.ms-excel.sharedStrings+xml"
]);
const TABLE_CONTENT_TYPE =
  "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml";
const TABLE_RELATIONSHIP_TYPES = new Set([
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table",
  "http://purl.oclc.org/ooxml/officeDocument/relationships/table"
]);

const utf8Decoder = new TextDecoder("utf-8", { fatal: true });
const utf16LeDecoder = new TextDecoder("utf-16le", { fatal: true });
const utf16BeDecoder = new TextDecoder("utf-16be", { fatal: true });

function reject(code: SecurityIssueCode, message: string): never {
  throw new XlsxSecurityError(code, message);
}

function ensureRange(
  bytes: Uint8Array,
  offset: number,
  length: number,
  context: string
): void {
  if (
    !Number.isSafeInteger(offset) ||
    !Number.isSafeInteger(length) ||
    offset < 0 ||
    length < 0 ||
    offset > bytes.length ||
    length > bytes.length - offset
  ) {
    reject("XLSX_ARCHIVE_LIMIT", `Invalid ZIP offset for ${context}.`);
  }
}

function readU16(bytes: Uint8Array, offset: number, context: string): number {
  ensureRange(bytes, offset, 2, context);
  return bytes[offset] | (bytes[offset + 1] << 8);
}

function readU32(bytes: Uint8Array, offset: number, context: string): number {
  ensureRange(bytes, offset, 4, context);
  return (
    (bytes[offset] |
      (bytes[offset + 1] << 8) |
      (bytes[offset + 2] << 16) |
      (bytes[offset + 3] << 24)) >>>
    0
  );
}

function checkedAdd(left: number, right: number, context: string): number {
  const result = left + right;
  if (!Number.isSafeInteger(result) || result < left) {
    reject("XLSX_ARCHIVE_LIMIT", `ZIP size arithmetic overflow for ${context}.`);
  }
  return result;
}

function findEndOfCentralDirectory(bytes: Uint8Array): number {
  if (bytes.length < 22) {
    reject("XLSX_ARCHIVE_LIMIT", "The XLSX ZIP is missing its central directory.");
  }

  const firstCandidate = Math.max(0, bytes.length - 22 - ZIP64_U16_SENTINEL);
  for (let offset = bytes.length - 22; offset >= firstCandidate; offset -= 1) {
    if (
      readU32(bytes, offset, "end-of-central-directory signature") !==
      END_OF_CENTRAL_DIRECTORY_SIGNATURE
    ) {
      continue;
    }
    const commentLength = readU16(bytes, offset + 20, "ZIP comment length");
    if (offset + 22 + commentLength === bytes.length) return offset;
  }

  reject("XLSX_ARCHIVE_LIMIT", "The XLSX ZIP central directory is malformed.");
}

function decodeZipName(bytes: Uint8Array, flags: number): string {
  try {
    if ((flags & UTF8_NAME_FLAG) !== 0 || bytes.every((value) => value < 0x80)) {
      return utf8Decoder.decode(bytes);
    }
  } catch {
    reject("XLSX_ARCHIVE_LIMIT", "A ZIP entry name is not valid UTF-8.");
  }
  reject("XLSX_ARCHIVE_LIMIT", "Non-UTF-8 ZIP entry names are not supported.");
}

function validateEntryName(name: string): void {
  const directoryName = name.endsWith("/") ? name.slice(0, -1) : name;
  const segments = directoryName.split("/");
  if (
    name.length === 0 ||
    directoryName.length === 0 ||
    name.startsWith("/") ||
    name.startsWith("\\") ||
    name.includes("\\") ||
    name.includes("\0") ||
    segments.some(
      (segment) =>
        segment.length === 0 ||
        segment === "." ||
        segment === ".." ||
        segment.includes(":")
    )
  ) {
    reject("XLSX_ARCHIVE_LIMIT", `Unsafe ZIP entry name: ${name || "<empty>"}.`);
  }
}

function validateExtraFields(bytes: Uint8Array, context: string): void {
  let offset = 0;
  while (offset < bytes.length) {
    if (bytes.length - offset < 4) {
      reject("XLSX_ARCHIVE_LIMIT", `Malformed ZIP extra field for ${context}.`);
    }
    const fieldId = readU16(bytes, offset, `${context} extra-field ID`);
    const fieldLength = readU16(bytes, offset + 2, `${context} extra-field length`);
    offset += 4;
    if (fieldLength > bytes.length - offset) {
      reject("XLSX_ARCHIVE_LIMIT", `Malformed ZIP extra field for ${context}.`);
    }
    if (fieldId === ZIP64_EXTRA_FIELD_ID) {
      reject("XLSX_ARCHIVE_LIMIT", "ZIP64 archives are not accepted.");
    }
    offset += fieldLength;
  }
}

function byteRangesEqual(
  left: Uint8Array,
  leftOffset: number,
  right: Uint8Array,
  rightOffset: number,
  length: number
): boolean {
  for (let index = 0; index < length; index += 1) {
    if (left[leftOffset + index] !== right[rightOffset + index]) return false;
  }
  return true;
}

function isTableXmlPath(name: string): boolean {
  return /^xl\/tables\/[^/]+\.xml$/i.test(name);
}

function isWorksheetXmlPath(name: string): boolean {
  return /^xl\/worksheets\/[^/]+\.xml$/i.test(name);
}

function isSharedStringsXmlPath(name: string): boolean {
  return name.toLowerCase() === "xl/sharedstrings.xml";
}

function hasExplicitLimit(
  overrides: Partial<XlsxSecurityLimits> | undefined,
  name: keyof XlsxSecurityLimits
): boolean {
  return Object.prototype.hasOwnProperty.call(overrides ?? {}, name);
}

function parseArchive(
  bytes: Uint8Array,
  limits: XlsxSecurityLimits,
  limitOverrides: Partial<XlsxSecurityLimits> | undefined
): ParsedArchive {
  const eocdOffset = findEndOfCentralDirectory(bytes);
  const diskNumber = readU16(bytes, eocdOffset + 4, "ZIP disk number");
  const centralDisk = readU16(bytes, eocdOffset + 6, "central-directory disk");
  const entriesOnDisk = readU16(bytes, eocdOffset + 8, "entries on disk");
  const entryCount = readU16(bytes, eocdOffset + 10, "entry count");
  const centralSize = readU32(bytes, eocdOffset + 12, "central-directory size");
  const centralOffset = readU32(bytes, eocdOffset + 16, "central-directory offset");

  if (
    diskNumber !== 0 ||
    centralDisk !== 0 ||
    entriesOnDisk !== entryCount
  ) {
    reject("XLSX_ARCHIVE_LIMIT", "Multi-disk ZIP archives are not accepted.");
  }
  if (
    entriesOnDisk === ZIP64_U16_SENTINEL ||
    entryCount === ZIP64_U16_SENTINEL ||
    centralSize === ZIP64_U32_SENTINEL ||
    centralOffset === ZIP64_U32_SENTINEL
  ) {
    reject("XLSX_ARCHIVE_LIMIT", "ZIP64 archives are not accepted.");
  }
  if (entryCount > limits.maxEntries) {
    reject(
      "XLSX_ARCHIVE_LIMIT",
      `The XLSX archive has more than ${limits.maxEntries} entries.`
    );
  }

  const centralEnd = checkedAdd(
    centralOffset,
    centralSize,
    "central directory"
  );
  if (centralEnd !== eocdOffset) {
    reject("XLSX_ARCHIVE_LIMIT", "The ZIP central-directory offsets are invalid.");
  }

  const entries: CentralDirectoryEntry[] = [];
  const names = new Set<string>();
  let totalUncompressedSize = 0;
  let cursor = centralOffset;

  for (let index = 0; index < entryCount; index += 1) {
    ensureRange(bytes, cursor, 46, `central-directory entry ${index + 1}`);
    if (
      readU32(bytes, cursor, `central-directory entry ${index + 1}`) !==
      CENTRAL_DIRECTORY_SIGNATURE
    ) {
      reject("XLSX_ARCHIVE_LIMIT", "A ZIP central-directory entry is malformed.");
    }

    const flags = readU16(bytes, cursor + 8, "central ZIP flags");
    const compressionMethod = readU16(
      bytes,
      cursor + 10,
      "central compression method"
    );
    const crc = readU32(bytes, cursor + 16, "central CRC-32");
    const compressedSize = readU32(bytes, cursor + 20, "central compressed size");
    const uncompressedSize = readU32(
      bytes,
      cursor + 24,
      "central uncompressed size"
    );
    const nameLength = readU16(bytes, cursor + 28, "central name length");
    const extraLength = readU16(bytes, cursor + 30, "central extra length");
    const commentLength = readU16(bytes, cursor + 32, "central comment length");
    const startDisk = readU16(bytes, cursor + 34, "central starting disk");
    const localHeaderOffset = readU32(
      bytes,
      cursor + 42,
      "local-header offset"
    );

    if ((flags & ENCRYPTION_FLAGS) !== 0) {
      reject("XLSX_ARCHIVE_LIMIT", "Encrypted ZIP entries are not accepted.");
    }
    if (compressionMethod !== 0 && compressionMethod !== 8) {
      reject("XLSX_ARCHIVE_LIMIT", "The ZIP uses an unsupported compression method.");
    }
    if (
      compressedSize === ZIP64_U32_SENTINEL ||
      uncompressedSize === ZIP64_U32_SENTINEL ||
      localHeaderOffset === ZIP64_U32_SENTINEL ||
      startDisk === ZIP64_U16_SENTINEL
    ) {
      reject("XLSX_ARCHIVE_LIMIT", "ZIP64 archives are not accepted.");
    }
    if (startDisk !== 0) {
      reject("XLSX_ARCHIVE_LIMIT", "Multi-disk ZIP archives are not accepted.");
    }

    let recordEnd = checkedAdd(cursor, 46, "central-directory record");
    recordEnd = checkedAdd(recordEnd, nameLength, "central-directory name");
    recordEnd = checkedAdd(recordEnd, extraLength, "central-directory extra field");
    recordEnd = checkedAdd(recordEnd, commentLength, "central-directory comment");
    if (recordEnd > centralEnd) {
      reject("XLSX_ARCHIVE_LIMIT", "A ZIP central-directory record is truncated.");
    }

    const nameBytes = bytes.subarray(cursor + 46, cursor + 46 + nameLength);
    const name = decodeZipName(nameBytes, flags);
    validateEntryName(name);
    const foldedName = name.toLowerCase();
    if (names.has(foldedName)) {
      reject("XLSX_ARCHIVE_LIMIT", `Duplicate ZIP entry name: ${name}.`);
    }
    names.add(foldedName);

    validateExtraFields(
      bytes.subarray(
        cursor + 46 + nameLength,
        cursor + 46 + nameLength + extraLength
      ),
      name
    );

    const maxEntryBytes = hasExplicitLimit(limitOverrides, "maxEntryBytes")
      ? limits.maxEntryBytes
      : isWorksheetXmlPath(name) || isSharedStringsXmlPath(name)
        ? LARGE_XML_ENTRY_BYTES
        : limits.maxEntryBytes;
    if (uncompressedSize > maxEntryBytes) {
      reject(
        "XLSX_ARCHIVE_LIMIT",
        `ZIP entry ${name} exceeds the per-entry size limit.`
      );
    }
    if (isTableXmlPath(name) && uncompressedSize > limits.maxTableXmlBytes) {
      reject(
        "XLSX_ARCHIVE_LIMIT",
        `Native table XML ${name} exceeds the table XML size limit.`
      );
    }
    totalUncompressedSize = checkedAdd(
      totalUncompressedSize,
      uncompressedSize,
      "total uncompressed size"
    );
    if (totalUncompressedSize > limits.maxTotalUncompressedBytes) {
      reject(
        "XLSX_ARCHIVE_LIMIT",
        "The XLSX archive exceeds the total uncompressed size limit."
      );
    }
    if (
      uncompressedSize > 0 &&
      (compressedSize === 0 ||
        uncompressedSize / compressedSize > limits.maxCompressionRatio)
    ) {
      reject(
        "XLSX_ARCHIVE_LIMIT",
        `ZIP entry ${name} exceeds the compression-ratio limit.`
      );
    }

    entries.push({
      name,
      flags,
      compressionMethod,
      crc32: crc,
      compressedSize,
      uncompressedSize,
      localHeaderOffset,
      centralNameOffset: cursor + 46,
      centralNameLength: nameLength,
      dataOffset: 0,
      localEndOffset: 0
    });
    cursor = recordEnd;
  }

  if (cursor !== centralEnd) {
    reject("XLSX_ARCHIVE_LIMIT", "The ZIP central-directory size is inconsistent.");
  }

  for (const entry of entries) {
    const localOffset = entry.localHeaderOffset;
    ensureRange(bytes, localOffset, 30, `local header for ${entry.name}`);
    if (
      localOffset >= centralOffset ||
      readU32(bytes, localOffset, `local header for ${entry.name}`) !==
        LOCAL_FILE_HEADER_SIGNATURE
    ) {
      reject("XLSX_ARCHIVE_LIMIT", `Invalid local ZIP header for ${entry.name}.`);
    }

    const localFlags = readU16(bytes, localOffset + 6, "local ZIP flags");
    const localMethod = readU16(bytes, localOffset + 8, "local compression method");
    const localCrc = readU32(bytes, localOffset + 14, "local CRC-32");
    const localCompressedSize = readU32(
      bytes,
      localOffset + 18,
      "local compressed size"
    );
    const localUncompressedSize = readU32(
      bytes,
      localOffset + 22,
      "local uncompressed size"
    );
    const localNameLength = readU16(bytes, localOffset + 26, "local name length");
    const localExtraLength = readU16(bytes, localOffset + 28, "local extra length");

    if (
      localFlags !== entry.flags ||
      localMethod !== entry.compressionMethod ||
      localNameLength === 0
    ) {
      reject("XLSX_ARCHIVE_LIMIT", `Inconsistent local ZIP header for ${entry.name}.`);
    }

    let dataOffset = checkedAdd(localOffset, 30, `local header for ${entry.name}`);
    dataOffset = checkedAdd(dataOffset, localNameLength, `local name for ${entry.name}`);
    dataOffset = checkedAdd(dataOffset, localExtraLength, `local extra for ${entry.name}`);
    if (dataOffset > centralOffset) {
      reject("XLSX_ARCHIVE_LIMIT", `Truncated local ZIP header for ${entry.name}.`);
    }

    if (
      localNameLength !== entry.centralNameLength ||
      !byteRangesEqual(
        bytes,
        localOffset + 30,
        bytes,
        entry.centralNameOffset,
        localNameLength
      )
    ) {
      reject("XLSX_ARCHIVE_LIMIT", `Inconsistent ZIP entry name for ${entry.name}.`);
    }

    validateExtraFields(
      bytes.subarray(
        localOffset + 30 + localNameLength,
        localOffset + 30 + localNameLength + localExtraLength
      ),
      entry.name
    );

    const usesDescriptor = (entry.flags & DATA_DESCRIPTOR_FLAG) !== 0;
    if (usesDescriptor) {
      if (
        (localCrc !== 0 && localCrc !== entry.crc32) ||
        (localCompressedSize !== 0 &&
          localCompressedSize !== entry.compressedSize) ||
        (localUncompressedSize !== 0 &&
          localUncompressedSize !== entry.uncompressedSize)
      ) {
        reject("XLSX_ARCHIVE_LIMIT", `Inconsistent ZIP sizes for ${entry.name}.`);
      }
    } else if (
      localCrc !== entry.crc32 ||
      localCompressedSize !== entry.compressedSize ||
      localUncompressedSize !== entry.uncompressedSize
    ) {
      reject("XLSX_ARCHIVE_LIMIT", `Inconsistent ZIP sizes for ${entry.name}.`);
    }

    if (
      entry.compressionMethod === 0 &&
      entry.compressedSize !== entry.uncompressedSize
    ) {
      reject("XLSX_ARCHIVE_LIMIT", `Invalid stored ZIP sizes for ${entry.name}.`);
    }

    let localEnd = checkedAdd(dataOffset, entry.compressedSize, entry.name);
    if (localEnd > centralOffset) {
      reject("XLSX_ARCHIVE_LIMIT", `Compressed data for ${entry.name} is truncated.`);
    }

    if (usesDescriptor) {
      let descriptorOffset = localEnd;
      ensureRange(bytes, descriptorOffset, 12, `data descriptor for ${entry.name}`);
      if (
        readU32(bytes, descriptorOffset, `data descriptor for ${entry.name}`) ===
        DATA_DESCRIPTOR_SIGNATURE
      ) {
        descriptorOffset += 4;
        ensureRange(bytes, descriptorOffset, 12, `data descriptor for ${entry.name}`);
      }
      const descriptorCrc = readU32(
        bytes,
        descriptorOffset,
        `data descriptor CRC for ${entry.name}`
      );
      const descriptorCompressed = readU32(
        bytes,
        descriptorOffset + 4,
        `data descriptor compressed size for ${entry.name}`
      );
      const descriptorUncompressed = readU32(
        bytes,
        descriptorOffset + 8,
        `data descriptor uncompressed size for ${entry.name}`
      );
      if (
        descriptorCrc !== entry.crc32 ||
        descriptorCompressed !== entry.compressedSize ||
        descriptorUncompressed !== entry.uncompressedSize
      ) {
        reject("XLSX_ARCHIVE_LIMIT", `Invalid data descriptor for ${entry.name}.`);
      }
      localEnd = descriptorOffset + 12;
      if (localEnd > centralOffset) {
        reject("XLSX_ARCHIVE_LIMIT", `Data descriptor for ${entry.name} is truncated.`);
      }
    }

    entry.dataOffset = dataOffset;
    entry.localEndOffset = localEnd;
  }

  const orderedEntries = [...entries].sort(
    (left, right) => left.localHeaderOffset - right.localHeaderOffset
  );
  for (let index = 1; index < orderedEntries.length; index += 1) {
    if (
      orderedEntries[index].localHeaderOffset <
      orderedEntries[index - 1].localEndOffset
    ) {
      reject("XLSX_ARCHIVE_LIMIT", "ZIP local entry ranges overlap.");
    }
  }

  return {
    bytes,
    entries,
    entriesByName: new Map(entries.map((entry) => [entry.name, entry]))
  };
}

const crcTable = (() => {
  const table = new Uint32Array(256);
  for (let index = 0; index < table.length; index += 1) {
    let value = index;
    for (let bit = 0; bit < 8; bit += 1) {
      value = (value & 1) !== 0 ? 0xedb88320 ^ (value >>> 1) : value >>> 1;
    }
    table[index] = value >>> 0;
  }
  return table;
})();

function calculateCrc32(bytes: Uint8Array): number {
  let crc = 0xffffffff;
  for (const value of bytes) {
    crc = crcTable[(crc ^ value) & 0xff] ^ (crc >>> 8);
  }
  return (crc ^ 0xffffffff) >>> 0;
}

function inflateEntry(archive: ParsedArchive, entry: CentralDirectoryEntry): Uint8Array {
  const compressed = archive.bytes.subarray(
    entry.dataOffset,
    entry.dataOffset + entry.compressedSize
  );
  let inflated: Uint8Array;

  try {
    if (entry.compressionMethod === 0) {
      inflated = compressed.slice();
    } else {
      const boundedOutput = new Uint8Array(entry.uncompressedSize + 1);
      inflated = inflateSync(compressed, { out: boundedOutput });
    }
  } catch {
    reject("XLSX_ARCHIVE_LIMIT", `ZIP entry ${entry.name} cannot be safely inflated.`);
  }

  if (inflated.byteLength !== entry.uncompressedSize) {
    reject(
      "XLSX_ARCHIVE_LIMIT",
      `ZIP entry ${entry.name} expands to an inconsistent size.`
    );
  }
  if (calculateCrc32(inflated) !== entry.crc32) {
    reject("XLSX_ARCHIVE_LIMIT", `ZIP entry ${entry.name} has an invalid CRC-32.`);
  }
  return inflated.slice();
}

function decodeXml(bytes: Uint8Array, name: string): string {
  try {
    if (bytes.length >= 2 && bytes[0] === 0xff && bytes[1] === 0xfe) {
      return utf16LeDecoder.decode(bytes.subarray(2));
    }
    if (bytes.length >= 2 && bytes[0] === 0xfe && bytes[1] === 0xff) {
      return utf16BeDecoder.decode(bytes.subarray(2));
    }
    return utf8Decoder.decode(bytes);
  } catch {
    reject("XLSX_XML_UNSAFE", `XML part ${name} has an invalid character encoding.`);
  }
}

type XmlCounts = { elements: number; attributes: number };
type XmlPackageBudget = {
  counts: XmlCounts;
  scaledGrants: XmlCounts;
  scaledGrantsByPart: Map<string, XmlCounts>;
  pendingScaledParts: number;
};
type XmlPartKind = "generic" | "worksheet" | "shared-strings";

function checkedMultiply(left: number, right: number): number | null {
  const result = left * right;
  return Number.isSafeInteger(result) ? result : null;
}

function checkedBudgetAdd(left: number, right: number): number | null {
  const result = left + right;
  return Number.isSafeInteger(result) ? result : null;
}

function effectiveTotalXmlLimit(
  configuredLimit: number,
  absoluteLimit: number,
  scaledGrant: number,
  explicit: boolean,
  pendingScaledParts: number
): number {
  const fixedLimit = Math.min(configuredLimit, absoluteLimit);
  if (explicit) return fixedLimit;
  if (pendingScaledParts > 0) return absoluteLimit;
  return Math.min(absoluteLimit, Math.max(fixedLimit, scaledGrant));
}

function enforcePackageCounts(
  packageBudget: XmlPackageBudget,
  limits: XlsxSecurityLimits,
  limitOverrides: Partial<XlsxSecurityLimits> | undefined
): void {
  const elementLimit = effectiveTotalXmlLimit(
    limits.maxTotalXmlElements,
    MAX_PACKAGE_XML_ELEMENTS,
    packageBudget.scaledGrants.elements,
    hasExplicitLimit(limitOverrides, "maxTotalXmlElements"),
    packageBudget.pendingScaledParts
  );
  const attributeLimit = effectiveTotalXmlLimit(
    limits.maxTotalXmlAttributes,
    MAX_PACKAGE_XML_ATTRIBUTES,
    packageBudget.scaledGrants.attributes,
    hasExplicitLimit(limitOverrides, "maxTotalXmlAttributes"),
    packageBudget.pendingScaledParts
  );
  if (packageBudget.counts.elements > elementLimit) {
    reject("XLSX_XML_UNSAFE", "The XLSX package exceeds the total XML element limit.");
  }
  if (packageBudget.counts.attributes > attributeLimit) {
    reject("XLSX_XML_UNSAFE", "The XLSX package exceeds the total XML attribute limit.");
  }
}

function worksheetColumnNumber(value: string): number | null {
  let column = 0;
  for (const character of value) {
    column = column * 26 + character.charCodeAt(0) - 64;
    if (!Number.isSafeInteger(column) || column > MAX_WORKSHEET_COLUMNS) return null;
  }
  return column > 0 ? column : null;
}

function worksheetDimensionBudget(value: string): XmlCounts | null {
  const match = /^([A-Z]{1,3})([1-9]\d*)(?::([A-Z]{1,3})([1-9]\d*))?$/.exec(
    value
  );
  if (!match) return null;

  const startColumn = worksheetColumnNumber(match[1]);
  const endColumn = worksheetColumnNumber(match[3] ?? match[1]);
  const startRow = Number(match[2]);
  const endRow = Number(match[4] ?? match[2]);
  if (
    startColumn === null ||
    endColumn === null ||
    !Number.isSafeInteger(startRow) ||
    !Number.isSafeInteger(endRow) ||
    startRow < 1 ||
    endRow < startRow ||
    endRow > MAX_WORKSHEET_ROWS ||
    endColumn < startColumn
  ) {
    return null;
  }

  const rows = endRow - startRow + 1;
  const columns = endColumn - startColumn + 1;
  const cells = checkedMultiply(rows, columns);
  if (cells === null) return null;
  if (!isSupportedXlsxWorksheetSize(rows, columns)) {
    reject(
      "XLSX_SHEET_TOO_LARGE",
      `Worksheet dimension ${value} is too large to import safely.`
    );
  }

  const cellElements = checkedMultiply(cells, 3);
  const rowAttributes = checkedMultiply(rows, 4);
  const cellAttributes = checkedMultiply(cells, 4);
  if (cellElements === null || rowAttributes === null || cellAttributes === null) {
    return null;
  }

  const elementsWithRows = checkedBudgetAdd(XML_BUDGET_BASE, rows);
  const attributesWithRows = checkedBudgetAdd(XML_BUDGET_BASE, rowAttributes);
  if (elementsWithRows === null || attributesWithRows === null) return null;
  const elements = checkedBudgetAdd(elementsWithRows, cellElements);
  const attributes = checkedBudgetAdd(attributesWithRows, cellAttributes);
  return elements === null || attributes === null ? null : { elements, attributes };
}

function safeNonNegativeInteger(value: string): number | null {
  if (!/^\d+$/.test(value)) return null;
  const parsed = Number(value);
  return Number.isSafeInteger(parsed) ? parsed : null;
}

function sharedStringsBudget(tag: SaxesTagNS): XmlCounts | null {
  const count = safeNonNegativeInteger(saxAttribute(tag, "count"));
  const uniqueCount = safeNonNegativeInteger(saxAttribute(tag, "uniqueCount"));
  if (count === null || uniqueCount === null || uniqueCount > count) return null;

  const stringElements = checkedMultiply(
    count,
    SHARED_STRING_XML_ELEMENTS_PER_COUNT
  );
  const stringAttributes = checkedMultiply(
    count,
    SHARED_STRING_XML_ATTRIBUTES_PER_COUNT
  );
  if (stringElements === null || stringAttributes === null) return null;
  const elements = checkedBudgetAdd(XML_BUDGET_BASE, stringElements);
  const attributes = checkedBudgetAdd(XML_BUDGET_BASE, stringAttributes);
  return elements === null || attributes === null ? null : { elements, attributes };
}

type TableRelationshipSummary = {
  target: string;
  targetMode: string;
};

type XmlPartSummary =
  | {
      kind: "content-types";
      rootName: string;
      overrides: ReadonlyMap<string, string>;
      defaults: ReadonlyMap<string, string>;
    }
  | {
      kind: "relationships";
      rootName: string;
      tableRelationships: readonly TableRelationshipSummary[];
    }
  | { kind: "other"; rootName: string };

function saxAttribute(tag: SaxesTagNS, wantedName: string): string {
  for (const attribute of Object.values(tag.attributes)) {
    if (attribute.prefix === "" && attribute.local === wantedName) return attribute.value;
  }
  return "";
}

function scanSafeXml(
  xml: string,
  name: string,
  partKind: XmlPartKind,
  limits: XlsxSecurityLimits,
  packageBudget: XmlPackageBudget,
  limitOverrides: Partial<XlsxSecurityLimits> | undefined
): XmlPartSummary {
  if (/<!\s*(?:DOCTYPE|ENTITY)\b/i.test(xml)) {
    reject("XLSX_XML_UNSAFE", `XML declarations are unsafe in ${name}.`);
  }

  const isContentTypes = name === "[Content_Types].xml";
  const isRelationships = name.toLowerCase().endsWith(".rels");
  const isWorksheet = partKind === "worksheet";
  const isSharedStrings = partKind === "shared-strings";
  const partCounts: XmlCounts = { elements: 0, attributes: 0 };
  const overrides = new Map<string, string>();
  const defaults = new Map<string, string>();
  const relationshipIds = new Set<string>();
  const tableRelationships: TableRelationshipSummary[] = [];
  const elementNames: string[] = [];
  let depth = 0;
  let rootName = "";
  let rootPrefix = "";
  let rootUri = "";
  const fixedElementLimit = Math.min(limits.maxXmlElements, MAX_PART_XML_ELEMENTS);
  const fixedAttributeLimit = Math.min(limits.maxXmlAttributes, MAX_PART_XML_ATTRIBUTES);
  let partElementLimit = fixedElementLimit;
  let partAttributeLimit = fixedAttributeLimit;
  let worksheetDimensionCount = 0;
  let worksheetSheetDataSeen = false;
  const replaceScaledGrant = (grant: XmlCounts): void => {
    const previous = packageBudget.scaledGrantsByPart.get(name) ?? {
      elements: 0,
      attributes: 0
    };
    const elements = checkedBudgetAdd(
      packageBudget.scaledGrants.elements - previous.elements,
      grant.elements
    );
    const attributes = checkedBudgetAdd(
      packageBudget.scaledGrants.attributes - previous.attributes,
      grant.attributes
    );
    if (elements === null || attributes === null) {
      reject("XLSX_ARCHIVE_LIMIT", "The XLSX package XML budget is too large.");
    }
    packageBudget.scaledGrants = { elements, attributes };
    if (grant.elements === 0 && grant.attributes === 0) {
      packageBudget.scaledGrantsByPart.delete(name);
    } else {
      packageBudget.scaledGrantsByPart.set(name, grant);
    }
    enforcePackageCounts(packageBudget, limits, limitOverrides);
  };

  const enforcePartCounts = (): void => {
    if (partCounts.elements > partElementLimit) {
      reject("XLSX_XML_UNSAFE", `XML part ${name} exceeds the XML element limit.`);
    }
    if (partCounts.attributes > partAttributeLimit) {
      reject("XLSX_XML_UNSAFE", `XML part ${name} exceeds the XML attribute limit.`);
    }
  };
  const applyScaledBudget = (budget: XmlCounts | null): void => {
    if (budget === null) {
      partElementLimit = fixedElementLimit;
      partAttributeLimit = fixedAttributeLimit;
    } else {
      if (!hasExplicitLimit(limitOverrides, "maxXmlElements")) {
        partElementLimit = Math.min(
          MAX_PART_XML_ELEMENTS,
          Math.max(fixedElementLimit, budget.elements)
        );
      }
      if (!hasExplicitLimit(limitOverrides, "maxXmlAttributes")) {
        partAttributeLimit = Math.min(
          MAX_PART_XML_ATTRIBUTES,
          Math.max(fixedAttributeLimit, budget.attributes)
        );
      }
    }
    replaceScaledGrant({
      elements:
        budget !== null && !hasExplicitLimit(limitOverrides, "maxXmlElements")
          ? partElementLimit
          : 0,
      attributes:
        budget !== null && !hasExplicitLimit(limitOverrides, "maxXmlAttributes")
          ? partAttributeLimit
          : 0
    });
    enforcePartCounts();
  };

  try {
    const parser = new SaxesParser({
      xmlns: true as const,
      fileName: name,
      position: false
    });
    parser.on("doctype", () => {
      reject("XLSX_XML_UNSAFE", `XML declarations are unsafe in ${name}.`);
    });
    parser.on("opentagstart", () => {
      depth += 1;
      partCounts.elements += 1;
      packageBudget.counts.elements += 1;
      if (depth > limits.maxXmlDepth) {
        reject("XLSX_XML_UNSAFE", `XML part ${name} exceeds the depth limit.`);
      }
      enforcePartCounts();
      enforcePackageCounts(packageBudget, limits, limitOverrides);
    });
    parser.on("attribute", () => {
      partCounts.attributes += 1;
      packageBudget.counts.attributes += 1;
      enforcePartCounts();
      enforcePackageCounts(packageBudget, limits, limitOverrides);
    });
    parser.on("opentag", (tag) => {
      if (depth === 1) {
        rootName = tag.local;
        rootPrefix = tag.prefix;
        rootUri = tag.uri;
      }

      if (
        name === "xl/workbook.xml" &&
        rootName === "workbook" &&
        rootPrefix === "" &&
        elementNames.includes("sheets") &&
        tag.prefix === "" &&
        tag.local === "sheet"
      ) {
        const worksheetId = Number.parseInt(saxAttribute(tag, "sheetId"), 10);
        if (worksheetId > limits.maxWorksheetId) {
          reject(
            "XLSX_ARCHIVE_LIMIT",
            `Workbook worksheet ID ${worksheetId} exceeds the ${limits.maxWorksheetId} consistency limit.`
          );
        }
      }
      elementNames.push(tag.prefix === "" ? tag.local : "");

      if (isWorksheet && depth === 2 && tag.local === "dimension") {
        worksheetDimensionCount += 1;
        const budget =
          worksheetDimensionCount === 1 &&
          !worksheetSheetDataSeen &&
          rootName === "worksheet" &&
          rootPrefix === "" &&
          SPREADSHEETML_NAMESPACES.has(rootUri) &&
          tag.prefix === "" &&
          tag.uri === rootUri
            ? worksheetDimensionBudget(saxAttribute(tag, "ref"))
            : null;
        applyScaledBudget(budget);
      } else if (
        isWorksheet &&
        depth === 2 &&
        rootName === "worksheet" &&
        rootPrefix === "" &&
        SPREADSHEETML_NAMESPACES.has(rootUri) &&
        tag.prefix === "" &&
        tag.uri === rootUri &&
        tag.local === "sheetData"
      ) {
        worksheetSheetDataSeen = true;
      } else if (
        isSharedStrings &&
        depth === 1 &&
        tag.prefix === "" &&
        tag.local === "sst" &&
        SPREADSHEETML_NAMESPACES.has(tag.uri)
      ) {
        applyScaledBudget(sharedStringsBudget(tag));
      }

      if (isContentTypes && tag.local === "Override") {
        const partName = normalizeAbsolutePartName(saxAttribute(tag, "PartName"));
        const contentType = saxAttribute(tag, "ContentType");
        if (!partName || !contentType || overrides.has(partName)) {
          reject(
            "XLSX_RELATIONSHIP_INVALID",
            "The XLSX package has an invalid or duplicate content-type override."
          );
        }
        if (overrides.size >= limits.maxEntries) {
          reject("XLSX_RELATIONSHIP_INVALID", "The XLSX package has too many content-type overrides.");
        }
        overrides.set(partName, contentType);
      } else if (isContentTypes && tag.local === "Default") {
        const extension = saxAttribute(tag, "Extension").replace(/^\./, "").toLowerCase();
        const contentType = saxAttribute(tag, "ContentType");
        if (!extension || !contentType || defaults.has(extension)) {
          reject(
            "XLSX_RELATIONSHIP_INVALID",
            "The XLSX package has an invalid or duplicate default content type."
          );
        }
        if (defaults.size >= limits.maxEntries) {
          reject("XLSX_RELATIONSHIP_INVALID", "The XLSX package has too many default content types.");
        }
        defaults.set(extension, contentType);
      } else if (isRelationships && tag.local === "Relationship") {
        const id = saxAttribute(tag, "Id");
        if (!id || relationshipIds.has(id)) {
          reject(
            "XLSX_RELATIONSHIP_INVALID",
            `Relationship part ${name} contains duplicate or blank IDs.`
          );
        }
        relationshipIds.add(id);
        const type = saxAttribute(tag, "Type");
        if (TABLE_RELATIONSHIP_TYPES.has(type)) {
          if (tableRelationships.length >= limits.maxEntries) {
            reject("XLSX_RELATIONSHIP_INVALID", `Relationship part ${name} has too many table relationships.`);
          }
          tableRelationships.push({
            target: saxAttribute(tag, "Target"),
            targetMode: saxAttribute(tag, "TargetMode")
          });
        }
      }
    });
    parser.on("closetag", () => {
      elementNames.pop();
      depth -= 1;
    });
    parser.write(xml).close();
  } catch (error) {
    if (error instanceof XlsxSecurityError) throw error;
    reject("XLSX_XML_UNSAFE", `XML part ${name} is malformed.`);
  }

  if (!rootName) {
    reject("XLSX_XML_UNSAFE", `XML part ${name} has no document element.`);
  }
  if (isContentTypes) {
    return { kind: "content-types", rootName, overrides, defaults };
  }
  if (isRelationships) {
    return { kind: "relationships", rootName, tableRelationships };
  }
  return { kind: "other", rootName };
}

function isXmlPart(name: string): boolean {
  const lowerName = name.toLowerCase();
  return lowerName.endsWith(".xml") || lowerName.endsWith(".rels");
}

function decodePackageUri(value: string): string | null {
  if (value.includes("?") || value.includes("#")) return null;
  try {
    const decoded = decodeURIComponent(value);
    if (decoded.includes("\\") || decoded.includes("\0")) return null;
    return decoded;
  } catch {
    return null;
  }
}

function normalizeAbsolutePartName(value: string): string | null {
  const decoded = decodePackageUri(value);
  if (!decoded?.startsWith("/") || decoded.startsWith("//")) return null;
  const segments = decoded.slice(1).split("/");
  if (
    segments.length === 0 ||
    segments.some(
      (segment) =>
        segment.length === 0 || segment === "." || segment === ".." || segment.includes(":")
    )
  ) {
    return null;
  }
  return segments.join("/");
}

function resolveRelationshipTarget(sourcePart: string, value: string): string | null {
  const decoded = decodePackageUri(value);
  if (!decoded || /^[a-z][a-z\d+.-]*:/i.test(decoded) || decoded.startsWith("//")) {
    return null;
  }

  const segments = decoded.startsWith("/")
    ? []
    : sourcePart.split("/").slice(0, -1);
  for (const segment of decoded.replace(/^\//, "").split("/")) {
    if (!segment || segment === ".") continue;
    if (segment === "..") {
      if (segments.length === 0) return null;
      segments.pop();
      continue;
    }
    if (segment.includes(":")) return null;
    segments.push(segment);
  }
  return segments.length > 0 ? segments.join("/") : null;
}

function sourcePartForRelationships(name: string): string | null {
  const marker = "/_rels/";
  const markerOffset = name.indexOf(marker);
  if (markerOffset < 0 || !name.endsWith(".rels")) return null;
  const prefix = name.slice(0, markerOffset);
  const relatedName = name.slice(markerOffset + marker.length, -".rels".length);
  if (!relatedName || relatedName.includes("/")) return null;
  return `${prefix}/${relatedName}`;
}

type RelationshipPartSummary = Extract<XmlPartSummary, { kind: "relationships" }>;
type ContentTypeResolver = (partName: string) => string | undefined;

function validatedXmlPartKind(
  name: string,
  contentTypeFor: ContentTypeResolver
): XmlPartKind {
  const contentType = contentTypeFor(name);
  if (isWorksheetXmlPath(name) && WORKSHEET_CONTENT_TYPES.has(contentType ?? "")) {
    return "worksheet";
  }
  if (
    isSharedStringsXmlPath(name) &&
    SHARED_STRINGS_CONTENT_TYPES.has(contentType ?? "")
  ) {
    return "shared-strings";
  }
  return "generic";
}

function enforceValidatedEntryByteLimit(
  entry: CentralDirectoryEntry,
  partKind: XmlPartKind,
  limits: XlsxSecurityLimits,
  limitOverrides: Partial<XlsxSecurityLimits> | undefined
): void {
  const maxEntryBytes = hasExplicitLimit(limitOverrides, "maxEntryBytes")
    ? limits.maxEntryBytes
    : partKind === "worksheet" || partKind === "shared-strings"
      ? LARGE_XML_ENTRY_BYTES
      : limits.maxEntryBytes;
  if (entry.uncompressedSize > maxEntryBytes) {
    reject(
      "XLSX_ARCHIVE_LIMIT",
      `ZIP entry ${entry.name} exceeds the per-entry size limit.`
    );
  }
}

function validateContentTypes(
  archive: ParsedArchive,
  contentTypes: XmlPartSummary | undefined
): ContentTypeResolver {
  if (
    !contentTypes ||
    contentTypes.kind !== "content-types" ||
    contentTypes.rootName !== "Types"
  ) {
    reject(
      "XLSX_RELATIONSHIP_INVALID",
      "The XLSX package is missing a valid [Content_Types].xml part."
    );
  }

  const contentTypeFor: ContentTypeResolver = (partName) => {
    const exact = contentTypes.overrides.get(partName);
    if (exact) return exact;
    const fileName = partName.split("/").pop() ?? "";
    const dot = fileName.lastIndexOf(".");
    return dot >= 0
      ? contentTypes.defaults.get(fileName.slice(dot + 1).toLowerCase())
      : undefined;
  };

  for (const entry of archive.entries) {
    if (
      isTableXmlPath(entry.name) &&
      contentTypeFor(entry.name) !== TABLE_CONTENT_TYPE
    ) {
      reject(
        "XLSX_RELATIONSHIP_INVALID",
        `Native table part ${entry.name} has a missing or invalid content type.`
      );
    }
  }
  return contentTypeFor;
}

function validateRelationships(
  archive: ParsedArchive,
  relationshipPartName: string,
  summary: RelationshipPartSummary,
  contentTypeFor: ContentTypeResolver
): void {
  const isWorksheetRelationships = relationshipPartName.startsWith(
    "xl/worksheets/_rels/"
  );
  const sourcePart = isWorksheetRelationships
    ? sourcePartForRelationships(relationshipPartName)
    : null;
  if (
    isWorksheetRelationships &&
    (!sourcePart || !archive.entriesByName.has(sourcePart))
  ) {
    reject(
      "XLSX_RELATIONSHIP_INVALID",
      `Relationship source for ${relationshipPartName} is missing.`
    );
  }
  if (summary.rootName !== "Relationships") {
    reject(
      "XLSX_RELATIONSHIP_INVALID",
      `Relationship part ${relationshipPartName} has an invalid root element.`
    );
  }

  for (const relationship of summary.tableRelationships) {
    if (!isWorksheetRelationships || !sourcePart) {
      reject(
        "XLSX_RELATIONSHIP_INVALID",
        "Native-table relationships must belong to a worksheet."
      );
    }
    if (relationship.targetMode && relationship.targetMode.toLowerCase() !== "internal") {
      reject(
        "XLSX_RELATIONSHIP_INVALID",
        "External native-table relationships are not accepted."
      );
    }
    const target = resolveRelationshipTarget(sourcePart, relationship.target);
    if (!target || !target.startsWith("xl/tables/") || !isTableXmlPath(target)) {
      reject(
        "XLSX_RELATIONSHIP_INVALID",
        "A native-table relationship target escapes xl/tables/."
      );
    }
    if (!archive.entriesByName.has(target)) {
      reject(
        "XLSX_RELATIONSHIP_INVALID",
        `Native-table relationship target ${target} is missing.`
      );
    }
    if (contentTypeFor(target) !== TABLE_CONTENT_TYPE) {
      reject(
        "XLSX_RELATIONSHIP_INVALID",
        `Native-table relationship target ${target} has the wrong content type.`
      );
    }
  }
}

function mergeLimits(overrides?: Partial<XlsxSecurityLimits>): XlsxSecurityLimits {
  const limits = { ...DEFAULT_XLSX_SECURITY_LIMITS, ...overrides };
  for (const [name, value] of Object.entries(limits)) {
    const integerRequired = name !== "maxCompressionRatio";
    if (
      typeof value !== "number" ||
      !Number.isFinite(value) ||
      value <= 0 ||
      (integerRequired && !Number.isSafeInteger(value))
    ) {
      reject("XLSX_ARCHIVE_LIMIT", `Invalid XLSX security limit: ${name}.`);
    }
  }
  return limits;
}

export function validateXlsxArchive(
  data: Uint8Array,
  limitOverrides?: Partial<XlsxSecurityLimits>
): ValidationResult {
  const result = inspectXlsxArchive(data, limitOverrides, false);
  return result.ok ? { ok: true } : result;
}

export function extractValidatedXlsxEntries(
  data: Uint8Array,
  limitOverrides?: Partial<XlsxSecurityLimits>
): XlsxEntryExtractionResult {
  return inspectXlsxArchive(data, limitOverrides, true);
}

function inspectXlsxArchive(
  data: Uint8Array,
  limitOverrides: Partial<XlsxSecurityLimits> | undefined,
  extractAllEntries: boolean
): XlsxEntryExtractionResult {
  try {
    const limits = mergeLimits(limitOverrides);
    if (data.byteLength > limits.maxArchiveBytes) {
      reject(
        "XLSX_ARCHIVE_LIMIT",
        `The XLSX archive exceeds the ${limits.maxArchiveBytes}-byte input limit.`
      );
    }

    // Validation owns an immutable snapshot. Callers cannot mutate the supplied view
    // while central-directory and XML checks are in progress.
    const bytes = data.slice();
    const archive = parseArchive(bytes, limits, limitOverrides);
    const extractedEntries = new Map<string, Uint8Array>();
    const packageXmlBudget: XmlPackageBudget = {
      counts: { elements: 0, attributes: 0 },
      scaledGrants: { elements: 0, attributes: 0 },
      scaledGrantsByPart: new Map(),
      pendingScaledParts: 0
    };
    const contentTypesEntry = archive.entriesByName.get("[Content_Types].xml");
    let contentTypes: XmlPartSummary | undefined;
    if (contentTypesEntry) {
      const xmlBytes = inflateEntry(archive, contentTypesEntry);
      if (extractAllEntries) extractedEntries.set(contentTypesEntry.name, xmlBytes);
      contentTypes = scanSafeXml(
        decodeXml(xmlBytes, contentTypesEntry.name),
        contentTypesEntry.name,
        "generic",
        limits,
        packageXmlBudget,
        limitOverrides
      );
    }
    const contentTypeFor = validateContentTypes(archive, contentTypes);
    const partKinds = new Map<string, XmlPartKind>();
    for (const entry of archive.entries) {
      const partKind = validatedXmlPartKind(entry.name, contentTypeFor);
      enforceValidatedEntryByteLimit(entry, partKind, limits, limitOverrides);
      partKinds.set(entry.name, partKind);
    }
    packageXmlBudget.pendingScaledParts = [...partKinds.values()].filter(
      (partKind) => partKind === "worksheet" || partKind === "shared-strings"
    ).length;

    for (const entry of archive.entries) {
      if (entry.name === contentTypesEntry?.name) continue;
      if (!extractAllEntries && !isXmlPart(entry.name)) continue;
      const xmlBytes = inflateEntry(archive, entry);
      if (extractAllEntries) extractedEntries.set(entry.name, xmlBytes);
      if (!isXmlPart(entry.name)) continue;
      const partKind = partKinds.get(entry.name) ?? "generic";
      const summary = scanSafeXml(
        decodeXml(xmlBytes, entry.name),
        entry.name,
        partKind,
        limits,
        packageXmlBudget,
        limitOverrides
      );
      if (summary.kind === "relationships") {
        validateRelationships(archive, entry.name, summary, contentTypeFor);
      }
      if (partKind === "worksheet" || partKind === "shared-strings") {
        packageXmlBudget.pendingScaledParts -= 1;
        enforcePackageCounts(packageXmlBudget, limits, limitOverrides);
      }
    }

    return { ok: true, entries: extractedEntries };
  } catch (error) {
    if (error instanceof XlsxSecurityError) {
      return {
        ok: false,
        issue: { code: error.code, message: error.message }
      };
    }
    return {
      ok: false,
      issue: {
        code: "XLSX_ARCHIVE_LIMIT",
        message: "The XLSX archive could not be safely validated."
      }
    };
  }
}
