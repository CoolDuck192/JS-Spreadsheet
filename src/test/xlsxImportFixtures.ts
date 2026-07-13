import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";

const VBA_RELATIONSHIP_TYPE =
  "http://schemas.microsoft.com/office/2006/relationships/vbaProject";
const XLSX_WORKBOOK_CONTENT_TYPE =
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";
const XLSM_WORKBOOK_CONTENT_TYPE =
  "application/vnd.ms-excel.sheet.macroEnabled.main+xml";

export function addSyntheticVbaProject(data: ArrayBuffer | Uint8Array): Uint8Array<ArrayBuffer> {
  const entries = unzipSync(data instanceof Uint8Array ? data : new Uint8Array(data));
  const workbookRelationships = "xl/_rels/workbook.xml.rels";
  entries[workbookRelationships] = strToU8(
    strFromU8(entries[workbookRelationships]).replace(
      "</Relationships>",
      `<Relationship Id="rIdSyntheticVba" Type="${VBA_RELATIONSHIP_TYPE}" Target="vbaProject.bin"/></Relationships>`
    )
  );
  entries["[Content_Types].xml"] = strToU8(
    strFromU8(entries["[Content_Types].xml"])
      .replace(XLSX_WORKBOOK_CONTENT_TYPE, XLSM_WORKBOOK_CONTENT_TYPE)
      .replace(
        "</Types>",
        '<Override PartName="/xl/vbaProject.bin" ContentType="application/vnd.ms-office.vbaProject"/></Types>'
      )
  );
  entries["xl/vbaProject.bin"] = new Uint8Array([0x56, 0x42, 0x41, 0x00]);
  const archive = zipSync(entries);
  const copy = new Uint8Array(archive.byteLength);
  copy.set(archive);
  return copy;
}

export function addUnsafeWorkbookDoctype(data: ArrayBuffer | Uint8Array): Uint8Array<ArrayBuffer> {
  const entries = unzipSync(data instanceof Uint8Array ? data : new Uint8Array(data));
  entries["xl/workbook.xml"] = strToU8(
    '<!DOCTYPE workbook [<!ENTITY xxe "unsafe">]>' + strFromU8(entries["xl/workbook.xml"])
  );
  const archive = zipSync(entries);
  const copy = new Uint8Array(archive.byteLength);
  copy.set(archive);
  return copy;
}

export function addMalformedWorkbookXml(data: ArrayBuffer | Uint8Array): Uint8Array<ArrayBuffer> {
  const entries = unzipSync(data instanceof Uint8Array ? data : new Uint8Array(data));
  entries["xl/workbook.xml"] = strToU8(
    strFromU8(entries["xl/workbook.xml"]).replace("</workbook>", "<broken></workbook>")
  );
  return copyArchive(zipSync(entries));
}

export function removeXlsxContentTypes(data: ArrayBuffer | Uint8Array): Uint8Array<ArrayBuffer> {
  const entries = unzipSync(data instanceof Uint8Array ? data : new Uint8Array(data));
  delete entries["[Content_Types].xml"];
  return copyArchive(zipSync(entries));
}

export function replaceNativeTableDocumentRoot(
  data: ArrayBuffer | Uint8Array
): Uint8Array<ArrayBuffer> {
  const entries = unzipSync(data instanceof Uint8Array ? data : new Uint8Array(data));
  const tableEntry = Object.keys(entries).find((name) => /^xl\/tables\/[^/]+\.xml$/i.test(name));
  if (!tableEntry) throw new Error("Expected an XLSX native table fixture");
  entries[tableEntry] = strToU8(
    strFromU8(entries[tableEntry])
      .replace(/<table\b/, "<notTable")
      .replace("</table>", "</notTable>")
  );
  return copyArchive(zipSync(entries));
}

export type XlsxZipConsistencyMismatch =
  | "invalidLocalHeader"
  | "inconsistentLocalHeader"
  | "inconsistentEntryName"
  | "inconsistentEntryNameWithNewline"
  | "inconsistentSizes"
  | "invalidStoredSizes";

export function addXlsxZipConsistencyMismatch(
  data: ArrayBuffer | Uint8Array,
  mismatch: XlsxZipConsistencyMismatch
): Uint8Array<ArrayBuffer> {
  const source = data instanceof Uint8Array ? data : new Uint8Array(data);
  const archive = new Uint8Array(source.byteLength);
  archive.set(source);
  const endOfCentralDirectory = findSignatureBackwards(archive, 0x06054b50);
  const centralOffset = readU32(archive, endOfCentralDirectory + 16);
  if (readU32(archive, centralOffset) !== 0x02014b50) {
    throw new Error("Expected a ZIP central-directory entry");
  }
  const localOffset = readU32(archive, centralOffset + 42);
  if (readU32(archive, localOffset) !== 0x04034b50) {
    throw new Error("Expected a ZIP local-file header");
  }

  switch (mismatch) {
    case "invalidLocalHeader":
      archive[localOffset] ^= 0xff;
      break;
    case "inconsistentLocalHeader": {
      const localMethod = readU16(archive, localOffset + 8);
      writeU16(archive, localOffset + 8, localMethod === 0 ? 8 : 0);
      break;
    }
    case "inconsistentEntryName":
      archive[localOffset + 30] ^= 0x01;
      break;
    case "inconsistentEntryNameWithNewline":
      archive[centralOffset + 46] = 0x0a;
      break;
    case "inconsistentSizes":
      writeU32(archive, localOffset + 22, readU32(archive, localOffset + 22) + 1);
      break;
    case "invalidStoredSizes":
      writeU16(archive, centralOffset + 10, 0);
      writeU16(archive, localOffset + 8, 0);
      break;
  }
  return archive;
}

function findSignatureBackwards(bytes: Uint8Array, signature: number): number {
  for (let offset = bytes.length - 4; offset >= 0; offset -= 1) {
    if (readU32(bytes, offset) === signature) return offset;
  }
  throw new Error("Expected ZIP signature was not found");
}

function readU16(bytes: Uint8Array, offset: number): number {
  return bytes[offset] | (bytes[offset + 1] << 8);
}

function readU32(bytes: Uint8Array, offset: number): number {
  return (
    bytes[offset]
    | (bytes[offset + 1] << 8)
    | (bytes[offset + 2] << 16)
    | (bytes[offset + 3] << 24)
  ) >>> 0;
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

function copyArchive(archive: Uint8Array): Uint8Array<ArrayBuffer> {
  const copy = new Uint8Array(archive.byteLength);
  copy.set(archive);
  return copy;
}
