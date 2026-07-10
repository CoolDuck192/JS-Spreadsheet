import { createHash } from "node:crypto";
import { mkdir, writeFile } from "node:fs/promises";
import { dirname, resolve } from "node:path";
import { DOMParser, XMLSerializer } from "@xmldom/xmldom";
import ExcelJS from "exceljs";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";

const OUTPUT = resolve("src/test/fixtures/xlsx/generated-sales-structured-table.xlsx");
const FIXED_DATE = new Date("2026-01-01T00:00:00.000Z");
const TABLE_XML = "xl/tables/table1.xml";

const workbook = new ExcelJS.Workbook();
workbook.creator = "JS Spreadsheet fixture generator";
workbook.created = FIXED_DATE;
workbook.modified = FIXED_DATE;
workbook.lastPrinted = FIXED_DATE;

const worksheet = workbook.addWorksheet("Sales");
worksheet.addTable({
  name: "SalesTable",
  ref: "A1",
  headerRow: true,
  totalsRow: true,
  style: {
    theme: "TableStyleLight9",
    showRowStripes: true,
    showFirstColumn: false,
    showLastColumn: false,
    showColumnStripes: false
  },
  columns: [
    { name: "Order ID", filterButton: true, totalsRowLabel: "Total" },
    { name: "Order Date", filterButton: true, totalsRowFunction: "none" },
    { name: "Region", filterButton: true, totalsRowFunction: "none" },
    { name: "Units", filterButton: true, totalsRowFunction: "none" },
    { name: "Unit Price", filterButton: true, totalsRowFunction: "none" },
    { name: "Amount", filterButton: true, totalsRowFunction: "sum" }
  ],
  rows: [
    ["SO-1001", new Date("2026-01-15T00:00:00.000Z"), "East", 2, 12.5, { formula: "D2*E2", result: 25 }],
    ["SO-1002", new Date("2026-01-16T00:00:00.000Z"), "West", 10, 3, { formula: "D3*E3", result: 30 }],
    ["SO-1003", new Date("2026-02-01T00:00:00.000Z"), "North", 1, 100, { formula: "D4*E4", result: 100 }],
    ["SO-1004", new Date("2026-02-05T00:00:00.000Z"), "East", 4, 7.5, { formula: "D5*E5", result: 30 }]
  ]
});

for (let row = 2; row <= 5; row += 1) {
  worksheet.getCell(`B${row}`).numFmt = "yyyy-mm-dd";
}

const excelBytes = new Uint8Array(await workbook.xlsx.writeBuffer());
const entries = unzipSync(excelBytes);
const tableXml = entries[TABLE_XML];
if (!tableXml) {
  throw new Error(`ExcelJS did not create ${TABLE_XML}`);
}

entries[TABLE_XML] = patchTableXml(tableXml);
const deterministicEntries = Object.fromEntries(
  Object.entries(entries)
    .sort(([left], [right]) => left.localeCompare(right))
    .map(([name, data]) => [name, [data, { mtime: FIXED_DATE }]])
);
const outputBytes = zipSync(deterministicEntries, { level: 6, mtime: FIXED_DATE });

await mkdir(dirname(OUTPUT), { recursive: true });
await writeFile(OUTPUT, outputBytes);
process.stdout.write(`${OUTPUT}\n${createHash("sha256").update(outputBytes).digest("hex")}\n`);

function patchTableXml(bytes) {
  const parser = new DOMParser({
    onError(level, message) {
      if (level !== "warning") {
        throw new Error(message);
      }
    }
  });
  const document = parser.parseFromString(strFromU8(bytes), "application/xml");
  const table = document.documentElement;
  const namespace = table.namespaceURI;
  const tableColumns = elements(table, "tableColumn");
  const amountColumn = tableColumns.find((node) => node.getAttribute("name") === "Amount");
  if (!amountColumn) {
    throw new Error("Amount table column is missing");
  }
  const calculated = document.createElementNS(namespace, "calculatedColumnFormula");
  calculated.appendChild(document.createTextNode("[@Units]*[@[Unit Price]]"));
  amountColumn.appendChild(calculated);

  const autoFilter = elements(table, "autoFilter")[0];
  if (!autoFilter) {
    throw new Error("Table autoFilter is missing");
  }
  const regionColumnIndex = tableColumns.findIndex((node) => node.getAttribute("name") === "Region");
  if (regionColumnIndex < 0) {
    throw new Error("Region table column is missing");
  }
  const regionColumnId = String(regionColumnIndex);
  for (const existing of elements(autoFilter, "filterColumn")) {
    if (existing.getAttribute("colId") === regionColumnId) {
      autoFilter.removeChild(existing);
    }
  }
  const filterColumn = document.createElementNS(namespace, "filterColumn");
  filterColumn.setAttribute("colId", regionColumnId);
  const filters = document.createElementNS(namespace, "filters");
  for (const value of ["East", "West"]) {
    const filter = document.createElementNS(namespace, "filter");
    filter.setAttribute("val", value);
    filters.appendChild(filter);
  }
  filterColumn.appendChild(filters);
  autoFilter.appendChild(filterColumn);

  return strToU8(new XMLSerializer().serializeToString(document));
}

function elements(root, localName) {
  return Array.from(root.getElementsByTagNameNS("*", localName));
}
