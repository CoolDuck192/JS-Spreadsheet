import {
  DOMParser,
  type Document as XmlDocument,
  type Element as XmlElement
} from "@xmldom/xmldom";
import { inflateSync } from "fflate";

import type { TableIssue } from "../core/commands/types";

export type XlsxSecurityLimits = {
  maxArchiveBytes: number;
  maxEntries: number;
  maxEntryBytes: number;
  maxTotalUncompressedBytes: number;
  maxCompressionRatio: number;
  maxTableXmlBytes: number;
  maxXmlDepth: number;
  maxXmlElements: number;
  maxXmlAttributes: number;
};

export const DEFAULT_XLSX_SECURITY_LIMITS: Readonly<XlsxSecurityLimits> = {
  maxArchiveBytes: 64 * 1024 * 1024,
  maxEntries: 4_096,
  maxEntryBytes: 32 * 1024 * 1024,
  maxTotalUncompressedBytes: 256 * 1024 * 1024,
  maxCompressionRatio: 100,
  maxTableXmlBytes: 4 * 1024 * 1024,
  maxXmlDepth: 128,
  maxXmlElements: 100_000,
  maxXmlAttributes: 200_000
};

type SecurityIssueCode =
  | "XLSX_ARCHIVE_LIMIT"
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

function parseArchive(
  bytes: Uint8Array,
  limits: XlsxSecurityLimits
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

    if (uncompressedSize > limits.maxEntryBytes) {
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

function parseSafeXml(
  xml: string,
  name: string,
  limits: XlsxSecurityLimits,
  counts: { elements: number; attributes: number }
): XmlDocument {
  if (/<!\s*(?:DOCTYPE|ENTITY)\b/i.test(xml)) {
    reject("XLSX_XML_UNSAFE", `XML declarations are unsafe in ${name}.`);
  }

  let document: XmlDocument;
  try {
    document = new DOMParser({
      onError(_level, message) {
        throw new Error(message);
      }
    }).parseFromString(xml, "application/xml");
  } catch {
    reject("XLSX_XML_UNSAFE", `XML part ${name} is malformed.`);
  }

  if (!document.documentElement) {
    reject("XLSX_XML_UNSAFE", `XML part ${name} has no document element.`);
  }

  const stack: Array<{ node: XmlElement; depth: number }> = [
    { node: document.documentElement, depth: 1 }
  ];
  while (stack.length > 0) {
    const current = stack.pop();
    if (!current) break;
    if (current.depth > limits.maxXmlDepth) {
      reject("XLSX_XML_UNSAFE", `XML part ${name} exceeds the depth limit.`);
    }
    counts.elements += 1;
    counts.attributes += current.node.attributes.length;
    if (counts.elements > limits.maxXmlElements) {
      reject("XLSX_XML_UNSAFE", "The XLSX package exceeds the XML element limit.");
    }
    if (counts.attributes > limits.maxXmlAttributes) {
      reject("XLSX_XML_UNSAFE", "The XLSX package exceeds the XML attribute limit.");
    }

    for (let index = 0; index < current.node.childNodes.length; index += 1) {
      const child = current.node.childNodes.item(index);
      if (child?.nodeType === 1) {
        stack.push({ node: child as XmlElement, depth: current.depth + 1 });
      }
    }
  }

  return document;
}

function isXmlPart(name: string): boolean {
  const lowerName = name.toLowerCase();
  return lowerName.endsWith(".xml") || lowerName.endsWith(".rels");
}

function localName(element: XmlElement): string {
  return element.localName || element.nodeName.split(":").pop() || element.nodeName;
}

function elementsNamed(document: XmlDocument, wantedName: string): XmlElement[] {
  const result: XmlElement[] = [];
  const all = document.getElementsByTagName("*");
  for (let index = 0; index < all.length; index += 1) {
    const element = all.item(index);
    if (element && localName(element) === wantedName) result.push(element);
  }
  return result;
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

function validateContentTypesAndRelationships(
  archive: ParsedArchive,
  documents: ReadonlyMap<string, XmlDocument>
): void {
  const contentTypesDocument = documents.get("[Content_Types].xml");
  if (
    !contentTypesDocument ||
    !contentTypesDocument.documentElement ||
    localName(contentTypesDocument.documentElement) !== "Types"
  ) {
    reject(
      "XLSX_RELATIONSHIP_INVALID",
      "The XLSX package is missing a valid [Content_Types].xml part."
    );
  }

  const overrides = new Map<string, string>();
  const defaults = new Map<string, string>();
  for (const element of elementsNamed(contentTypesDocument, "Override")) {
    const partName = normalizeAbsolutePartName(element.getAttribute("PartName") ?? "");
    const contentType = element.getAttribute("ContentType") ?? "";
    if (!partName || !contentType || overrides.has(partName)) {
      reject(
        "XLSX_RELATIONSHIP_INVALID",
        "The XLSX package has an invalid or duplicate content-type override."
      );
    }
    overrides.set(partName, contentType);
  }
  for (const element of elementsNamed(contentTypesDocument, "Default")) {
    const extension = (element.getAttribute("Extension") ?? "")
      .replace(/^\./, "")
      .toLowerCase();
    const contentType = element.getAttribute("ContentType") ?? "";
    if (!extension || !contentType || defaults.has(extension)) {
      reject(
        "XLSX_RELATIONSHIP_INVALID",
        "The XLSX package has an invalid or duplicate default content type."
      );
    }
    defaults.set(extension, contentType);
  }

  const contentTypeFor = (partName: string): string | undefined => {
    const exact = overrides.get(partName);
    if (exact) return exact;
    const fileName = partName.split("/").pop() ?? "";
    const dot = fileName.lastIndexOf(".");
    return dot >= 0 ? defaults.get(fileName.slice(dot + 1).toLowerCase()) : undefined;
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

  for (const [relationshipPartName, document] of documents) {
    if (!relationshipPartName.endsWith(".rels")) continue;

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
    if (
      !document.documentElement ||
      localName(document.documentElement) !== "Relationships"
    ) {
      reject(
        "XLSX_RELATIONSHIP_INVALID",
        `Relationship part ${relationshipPartName} has an invalid root element.`
      );
    }

    const ids = new Set<string>();
    for (const relationship of elementsNamed(document, "Relationship")) {
      const id = relationship.getAttribute("Id") ?? "";
      if (!id || ids.has(id)) {
        reject(
          "XLSX_RELATIONSHIP_INVALID",
          `Relationship part ${relationshipPartName} contains duplicate or blank IDs.`
        );
      }
      ids.add(id);

      const type = relationship.getAttribute("Type") ?? "";
      if (!TABLE_RELATIONSHIP_TYPES.has(type)) continue;

      if (!isWorksheetRelationships || !sourcePart) {
        reject(
          "XLSX_RELATIONSHIP_INVALID",
          "Native-table relationships must belong to a worksheet."
        );
      }

      const targetMode = relationship.getAttribute("TargetMode") ?? "";
      if (targetMode && targetMode.toLowerCase() !== "internal") {
        reject(
          "XLSX_RELATIONSHIP_INVALID",
          "External native-table relationships are not accepted."
        );
      }
      const target = resolveRelationshipTarget(
        sourcePart,
        relationship.getAttribute("Target") ?? ""
      );
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
    const archive = parseArchive(bytes, limits);
    const documents = new Map<string, XmlDocument>();
    const extractedEntries = new Map<string, Uint8Array>();
    const xmlCounts = { elements: 0, attributes: 0 };

    for (const entry of archive.entries) {
      if (!extractAllEntries && !isXmlPart(entry.name)) continue;
      const xmlBytes = inflateEntry(archive, entry);
      if (extractAllEntries) extractedEntries.set(entry.name, xmlBytes);
      if (!isXmlPart(entry.name)) continue;
      const document = parseSafeXml(
        decodeXml(xmlBytes, entry.name),
        entry.name,
        limits,
        xmlCounts
      );
      documents.set(entry.name, document);
    }

    validateContentTypesAndRelationships(archive, documents);
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
