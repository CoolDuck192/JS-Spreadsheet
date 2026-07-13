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

function copyArchive(archive: Uint8Array): Uint8Array<ArrayBuffer> {
  const copy = new Uint8Array(archive.byteLength);
  copy.set(archive);
  return copy;
}
