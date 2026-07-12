import { strToU8, zipSync } from "fflate";

const FIXED_ZIP_DATE = new Date("1980-01-01T00:00:00.000Z");
const DENSE_ROW_COUNT = 100_000;
const DENSE_COLUMNS = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J"] as const;
const WORKSHEET_CONTENT_TYPE =
  "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";
const TABLE_CONTENT_TYPE =
  "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml";
const TABLE_RELATIONSHIP_TYPE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table";

let denseWorksheetPackage: Uint8Array | undefined;

function makeDenseWorksheetXml(includeNumericType: boolean): Uint8Array {
  const rows = new Array<string>(DENSE_ROW_COUNT);
  for (let row = 1; row <= DENSE_ROW_COUNT; row += 1) {
    let cells = "";
    for (const column of DENSE_COLUMNS) {
      cells += `<c r="${column}${row}"${includeNumericType ? ' t="n"' : ""}><v>1</v></c>`;
    }
    rows[row - 1] = `<row r="${row}">${cells}</row>`;
  }

  return strToU8(
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' +
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      '<dimension ref="A1:J100000"/><sheetData>' +
      rows.join("") +
      '</sheetData><tableParts count="1"><tablePart r:id="rId1"/></tableParts></worksheet>'
  );
}

function packageWorksheet(
  worksheetXml: Uint8Array,
  worksheetContentType: string | null
): Uint8Array {
  const stored = { level: 0 as const, mtime: FIXED_ZIP_DATE };
  return zipSync({
    "[Content_Types].xml": [
      strToU8(
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
          (worksheetContentType === null
            ? ""
            : `<Override PartName="/xl/worksheets/sheet1.xml" ContentType="${worksheetContentType}"/>`) +
          `<Override PartName="/xl/tables/table1.xml" ContentType="${TABLE_CONTENT_TYPE}"/>` +
          "</Types>"
      ),
      stored
    ],
    "xl/worksheets/sheet1.xml": [worksheetXml, stored],
    "xl/worksheets/_rels/sheet1.xml.rels": [
      strToU8(
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
          `<Relationship Id="rId1" Type="${TABLE_RELATIONSHIP_TYPE}" Target="../tables/table1.xml"/>` +
          "</Relationships>"
      ),
      stored
    ],
    "xl/tables/table1.xml": [
      strToU8(
        '<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' +
          'id="1" name="DenseTable" displayName="DenseTable" ref="A1:J100000">' +
          '<autoFilter ref="A1:J100000"/><tableColumns count="10">' +
          DENSE_COLUMNS.map((_, index) =>
            `<tableColumn id="${index + 1}" name="Column${index + 1}"/>`
          ).join("") +
          '</tableColumns><tableStyleInfo name="TableStyleMedium2" ' +
          'showFirstColumn="0" showLastColumn="0" showRowStripes="1" ' +
          'showColumnStripes="0"/></table>'
      ),
      stored
    ]
  });
}

export function makeDenseWorksheetPackage(
  worksheetContentType: string | null = WORKSHEET_CONTENT_TYPE
): Uint8Array {
  if (worksheetContentType === WORKSHEET_CONTENT_TYPE && denseWorksheetPackage) {
    return denseWorksheetPackage;
  }

  const result = packageWorksheet(makeDenseWorksheetXml(false), worksheetContentType);
  if (worksheetContentType === WORKSHEET_CONTENT_TYPE) denseWorksheetPackage = result;
  return result;
}

export function makeLargeDenseWorksheetPackage(
  worksheetContentType: string | null = WORKSHEET_CONTENT_TYPE
): Uint8Array {
  return packageWorksheet(makeDenseWorksheetXml(true), worksheetContentType);
}
