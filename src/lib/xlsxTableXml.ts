import {
  DOMParser,
  XMLSerializer,
  type Document as XmlDocument,
  type Element as XmlElement
} from "@xmldom/xmldom";
import { strFromU8, strToU8, zipSync } from "fflate";
import type { FilterExpression, QueryScalar } from "../table/core/query";
import type {
  StructuredTable,
  StructuredTableColumn,
  TableAggregate,
  WorkbookModel
} from "../types";
import { formatCellAddress } from "./addressing";
import { a1FormulaToStructured } from "./structuredFormula";
import {
  extractValidatedXlsxEntries,
  validateXlsxArchive
} from "./xlsxSecurity";

export type NativeTableXmlMetadata = {
  name: string;
  calculatedColumns: Readonly<Record<string, string>>;
  /** Column IDs are native column names until the XLSX adapter remaps them. */
  filter?: FilterExpression;
  totals: Readonly<Record<string, {
    function?: string;
    formula?: string;
    label?: string;
  }>>;
};

const TABLE_ENTRY_PATTERN = /^xl\/tables\/[^/]+\.xml$/i;
const FIXED_ZIP_DATE = new Date("1980-01-01T00:00:00.000Z");
const XML_NAMESPACE = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

const NATIVE_TO_AGGREGATE: Readonly<Record<string, TableAggregate>> = {
  sum: "sum",
  average: "average",
  count: "count",
  countNums: "countNumbers",
  min: "min",
  max: "max",
  stdDev: "standardDeviation",
  var: "variance"
};

const AGGREGATE_TO_NATIVE: Readonly<Record<TableAggregate, string>> = {
  none: "none",
  sum: "sum",
  average: "average",
  count: "count",
  countNumbers: "countNums",
  min: "min",
  max: "max",
  standardDeviation: "stdDev",
  variance: "var"
};

export function readNativeTableXml(data: Uint8Array): readonly NativeTableXmlMetadata[] {
  const entries = validatedEntries(data);
  return [...entries.keys()]
    .filter((name) => TABLE_ENTRY_PATTERN.test(name))
    .sort(naturalEntryOrder)
    .map((name) => readTableDocument(strFromU8(entries.get(name)!)));
}

export function patchNativeTableXml(
  data: Uint8Array,
  tables: readonly StructuredTable[],
  workbook?: WorkbookModel
): Uint8Array {
  const entries = new Map(validatedEntries(data));
  const byName = new Map(tables.map((table) => [normalizeName(table.name), table]));

  for (const entryName of [...entries.keys()].filter((name) => TABLE_ENTRY_PATTERN.test(name))) {
    const document = parseXml(strFromU8(entries.get(entryName)!));
    const root = xmlRoot(document);
    const table = byName.get(normalizeName(root.getAttribute("name") ?? ""));
    if (!table) continue;
    patchTableDocument(document, table, workbook);
    entries.set(entryName, strToU8(new XMLSerializer().serializeToString(document)));
  }

  const output = zipEntries(entries);
  assertSafeArchive(output);
  return output;
}

/**
 * ExcelJS 4.4 cannot parse table-column formula children before the final
 * column. Import uses this validated derivative only for ExcelJS's cell/table
 * model; the original XML is read first and remains the metadata authority.
 */
export function prepareNativeTableXmlForExcelJs(data: Uint8Array): Uint8Array {
  const entries = new Map(validatedEntries(data));
  for (const entryName of [...entries.keys()].filter((name) => TABLE_ENTRY_PATTERN.test(name))) {
    const document = parseXml(strFromU8(entries.get(entryName)!));
    for (const column of allElementsByLocalName(document, "tableColumn")) {
      removeDirectChildren(column, "calculatedColumnFormula");
      removeDirectChildren(column, "totalsRowFormula");
    }
    entries.set(entryName, strToU8(new XMLSerializer().serializeToString(document)));
  }
  const output = zipEntries(entries);
  assertSafeArchive(output);
  return output;
}

function readTableDocument(xml: string): NativeTableXmlMetadata {
  const document = parseXml(xml);
  const root = xmlRoot(document);
  if (root.localName !== "table") throw tableXmlError("Expected a native table document");
  const name = root.getAttribute("name") || root.getAttribute("displayName");
  if (!name) throw tableXmlError("Native table is missing its name");

  const columns = directChildren(firstDirectChild(root, "tableColumns"), "tableColumn");
  const columnNames = columns.map((column) => column.getAttribute("name") ?? "");
  const calculatedColumns: Record<string, string> = {};
  const totals: Record<string, { function?: string; formula?: string; label?: string }> = {};
  for (const column of columns) {
    const columnName = column.getAttribute("name");
    if (!columnName) continue;
    const calculated = firstDirectChild(column, "calculatedColumnFormula")?.textContent;
    if (calculated) calculatedColumns[columnName] = ensureLeadingEquals(calculated);

    const label = column.getAttribute("totalsRowLabel");
    const nativeFunction = column.getAttribute("totalsRowFunction");
    const formula = firstDirectChild(column, "totalsRowFormula")?.textContent;
    if (label) totals[columnName] = { label };
    else if (formula) totals[columnName] = { function: "custom", formula: ensureLeadingEquals(formula) };
    else if (nativeFunction && nativeFunction !== "none") totals[columnName] = { function: nativeFunction };
  }

  const filter = readAutoFilter(firstDirectChild(root, "autoFilter"), columnNames);
  return {
    name,
    calculatedColumns,
    ...(filter ? { filter } : {}),
    totals
  };
}

function readAutoFilter(autoFilter: XmlElement | undefined, columnNames: readonly string[]): FilterExpression | undefined {
  if (!autoFilter) return undefined;
  const expressions: FilterExpression[] = [];
  for (const filterColumn of directChildren(autoFilter, "filterColumn")) {
    const columnIndex = parseNonNegativeInteger(filterColumn.getAttribute("colId") ?? "");
    const columnId = columnIndex === undefined ? undefined : columnNames[columnIndex];
    if (!columnId) continue;
    const filters = firstDirectChild(filterColumn, "filters");
    if (filters) {
      const values = directChildren(filters, "filter").map((filter) => ({
        type: "string" as const,
        value: filter.getAttribute("val") ?? ""
      }));
      if (values.length > 0) {
        expressions.push({ kind: "set", columnId, operator: "in", values });
      }
      continue;
    }
    const customFilters = firstDirectChild(filterColumn, "customFilters");
    if (!customFilters) continue;
    const customExpressions: Array<Extract<FilterExpression, { kind: "comparison" }>> = [];
    for (const custom of directChildren(customFilters, "customFilter")) {
      const operator = nativeComparisonOperator(custom.getAttribute("operator") ?? "equal");
      if (!operator) continue;
      customExpressions.push({
        kind: "comparison" as const,
        columnId,
        operator,
        value: { type: "string" as const, value: custom.getAttribute("val") ?? "" }
      });
    }
    if (customExpressions.length === 1) expressions.push(customExpressions[0]);
    else if (customExpressions.length > 1) {
      expressions.push({
        kind: "logical",
        operator: customFilters.getAttribute("and") === "1" ? "and" : "or",
        operands: customExpressions
      });
    }
  }
  if (expressions.length === 0) return undefined;
  return expressions.length === 1
    ? expressions[0]
    : { kind: "logical", operator: "and", operands: expressions };
}

function patchTableDocument(document: XmlDocument, table: StructuredTable, workbook?: WorkbookModel): void {
  const root = xmlRoot(document);
  patchStyle(document, root, table);
  const tableColumns = firstDirectChild(root, "tableColumns");
  if (!tableColumns) throw tableXmlError(`Native table ${table.name} has no tableColumns element`);
  const nativeColumns = directChildren(tableColumns, "tableColumn");
  for (const nativeColumn of nativeColumns) {
    const column = findColumnByName(table, nativeColumn.getAttribute("name") ?? "");
    if (!column) continue;
    patchCalculatedFormula(document, nativeColumn, table, column);
    patchTotals(document, nativeColumn, table, column, workbook);
  }
  if (table.filter) patchFilter(document, root, table, table.filter);
}

function patchStyle(document: XmlDocument, root: XmlElement, table: StructuredTable): void {
  if (!table.style) return;
  let style = firstDirectChild(root, "tableStyleInfo");
  if (!style) {
    style = document.createElementNS(root.namespaceURI || XML_NAMESPACE, "tableStyleInfo");
    root.appendChild(style);
  }
  style.setAttribute("name", table.style.theme);
  style.setAttribute("showFirstColumn", booleanAttribute(table.style.showFirstColumn));
  style.setAttribute("showLastColumn", booleanAttribute(table.style.showLastColumn));
  style.setAttribute("showRowStripes", booleanAttribute(table.style.showRowStripes));
  style.setAttribute("showColumnStripes", booleanAttribute(table.style.showColumnStripes));
}

function patchCalculatedFormula(
  document: XmlDocument,
  nativeColumn: XmlElement,
  table: StructuredTable,
  column: StructuredTableColumn
): void {
  removeDirectChildren(nativeColumn, "calculatedColumnFormula");
  if (!column.calculatedFormula) return;
  const translated = a1FormulaToStructured(column.calculatedFormula, table, firstBodyRow(table));
  if (!translated.ok) throw issueError(translated.issue);
  const formula = document.createElementNS(nativeColumn.namespaceURI || XML_NAMESPACE, "calculatedColumnFormula");
  formula.appendChild(document.createTextNode(withoutLeadingEquals(translated.formula)));
  nativeColumn.appendChild(formula);
}

function patchTotals(
  document: XmlDocument,
  nativeColumn: XmlElement,
  table: StructuredTable,
  column: StructuredTableColumn,
  workbook?: WorkbookModel
): void {
  if (column.totalsLabel !== undefined) {
    nativeColumn.removeAttribute("totalsRowFunction");
    nativeColumn.setAttribute("totalsRowLabel", column.totalsLabel);
    removeDirectChildren(nativeColumn, "totalsRowFormula");
    return;
  }
  if (column.totalsFunction && column.totalsFunction !== "none") {
    nativeColumn.removeAttribute("totalsRowLabel");
    nativeColumn.setAttribute("totalsRowFunction", AGGREGATE_TO_NATIVE[column.totalsFunction]);
    removeDirectChildren(nativeColumn, "totalsRowFormula");
    return;
  }
  if (!workbook || !table.totalsRow) return;
  const sheet = workbook.sheets.find((candidate) => candidate.id === table.sheetId);
  if (!sheet) return;
  const address = formatCellAddress({ row: table.range.end.row, column: column.sheetColumn });
  const value = sheet.cells[address];
  if (value === undefined || value === null || value === "") {
    nativeColumn.removeAttribute("totalsRowFunction");
    nativeColumn.removeAttribute("totalsRowLabel");
    removeDirectChildren(nativeColumn, "totalsRowFormula");
    return;
  }

  nativeColumn.removeAttribute("totalsRowLabel");
  nativeColumn.setAttribute("totalsRowFunction", "custom");
  removeDirectChildren(nativeColumn, "totalsRowFormula");
  if (typeof value === "string" && value.startsWith("=")) {
    const translated = a1FormulaToStructured(value, table, firstBodyRow(table));
    if (!translated.ok) throw issueError(translated.issue);
    const formula = document.createElementNS(nativeColumn.namespaceURI || XML_NAMESPACE, "totalsRowFormula");
    formula.appendChild(document.createTextNode(withoutLeadingEquals(translated.formula)));
    nativeColumn.appendChild(formula);
  }
}

function patchFilter(document: XmlDocument, root: XmlElement, table: StructuredTable, filter: FilterExpression): void {
  let autoFilter = firstDirectChild(root, "autoFilter");
  if (!autoFilter) {
    autoFilter = document.createElementNS(root.namespaceURI || XML_NAMESPACE, "autoFilter");
    const columns = firstDirectChild(root, "tableColumns");
    root.insertBefore(autoFilter, columns ?? root.firstChild);
  }
  const existingColumns = directChildren(autoFilter, "filterColumn");
  const existingByIndex = new Map<number, XmlElement[]>();
  for (const node of existingColumns) {
    const index = parseNonNegativeInteger(node.getAttribute("colId") ?? "");
    if (index !== undefined) existingByIndex.set(index, [...(existingByIndex.get(index) ?? []), node]);
    autoFilter.removeChild(node);
  }

  const leaves = flattenNativeFilter(filter);
  const leavesByColumn = new Map<number, (typeof leaves)[number]>();
  for (const leaf of leaves) {
    const columnIndex = table.columns.findIndex((column) => column.id === leaf.columnId);
    if (columnIndex < 0) throw tableXmlError("Native table filter refers to an unknown column");
    if (leavesByColumn.has(columnIndex)) {
      throw tableXmlError("Unsupported native table filter has multiple expressions for one column");
    }
    leavesByColumn.set(columnIndex, leaf);
  }

  for (let columnIndex = 0; columnIndex < table.columns.length; columnIndex += 1) {
    const leaf = leavesByColumn.get(columnIndex);
    const candidates = existingByIndex.get(columnIndex) ?? [];
    const filterColumn = candidates.find(hasUnsupportedFilterChild)
      ?? candidates[0]
      ?? document.createElementNS(root.namespaceURI || XML_NAMESPACE, "filterColumn");
    filterColumn.setAttribute("colId", String(columnIndex));
    if (!leaf) {
      if (!hasUnsupportedFilterChild(filterColumn)) {
        removeDirectChildren(filterColumn, "filters");
        removeDirectChildren(filterColumn, "customFilters");
      }
      if (!filterColumn.hasAttribute("hiddenButton")) filterColumn.setAttribute("hiddenButton", "0");
      autoFilter.appendChild(filterColumn);
      continue;
    }
    removeDirectChildren(filterColumn, "filters");
    removeDirectChildren(filterColumn, "customFilters");
    if (leaf.kind === "set") {
      if (leaf.operator !== "in") throw tableXmlError("Unsupported native table filter set operator");
      const filters = document.createElementNS(root.namespaceURI || XML_NAMESPACE, "filters");
      for (const value of leaf.values) {
        const filterNode = document.createElementNS(root.namespaceURI || XML_NAMESPACE, "filter");
        filterNode.setAttribute("val", scalarText(value));
        filters.appendChild(filterNode);
      }
      filterColumn.appendChild(filters);
    } else {
      const operator = appComparisonOperator(leaf.operator);
      if (!operator) throw tableXmlError("Unsupported native table filter comparison operator");
      const customFilters = document.createElementNS(root.namespaceURI || XML_NAMESPACE, "customFilters");
      const custom = document.createElementNS(root.namespaceURI || XML_NAMESPACE, "customFilter");
      if (operator !== "equal") custom.setAttribute("operator", operator);
      custom.setAttribute("val", scalarText(leaf.value));
      customFilters.appendChild(custom);
      filterColumn.appendChild(customFilters);
    }
    autoFilter.appendChild(filterColumn);
  }
}

function flattenNativeFilter(
  filter: FilterExpression
): Array<Extract<FilterExpression, { kind: "set" | "comparison" }>> {
  if (filter.kind === "set" || filter.kind === "comparison") return [filter];
  if (filter.kind === "logical" && filter.operator === "and") {
    return filter.operands.flatMap(flattenNativeFilter);
  }
  throw tableXmlError("Unsupported native table filter expression");
}

function nativeComparisonOperator(
  operator: string
): Extract<FilterExpression, { kind: "comparison" }>["operator"] | undefined {
  return ({
    equal: "eq",
    notEqual: "neq",
    greaterThan: "gt",
    greaterThanOrEqual: "gte",
    lessThan: "lt",
    lessThanOrEqual: "lte"
  } as const)[operator as "equal"];
}

function appComparisonOperator(
  operator: Extract<FilterExpression, { kind: "comparison" }>["operator"]
): string | undefined {
  return ({
    eq: "equal",
    neq: "notEqual",
    gt: "greaterThan",
    gte: "greaterThanOrEqual",
    lt: "lessThan",
    lte: "lessThanOrEqual"
  } as Partial<Record<typeof operator, string>>)[operator];
}

function scalarText(value: QueryScalar): string {
  switch (value.type) {
    case "null": return "";
    case "boolean": return value.value ? "1" : "0";
    default: return String(value.value);
  }
}

function parseXml(xml: string): XmlDocument {
  const errors: string[] = [];
  const document = new DOMParser({
    onError: (level, message) => {
      if (level === "error" || level === "fatalError") errors.push(message);
    }
  }).parseFromString(xml, "application/xml");
  if (errors.length > 0 || !document.documentElement || document.documentElement.localName === "parsererror") {
    throw tableXmlError("Malformed native table XML");
  }
  return document;
}

function directChildren(parent: XmlElement | null | undefined, localName: string): XmlElement[] {
  if (!parent) return [];
  const children: XmlElement[] = [];
  for (let node = parent.firstChild; node; node = node.nextSibling) {
    if (node.nodeType === 1 && (node as XmlElement).localName === localName) children.push(node as XmlElement);
  }
  return children;
}

function firstDirectChild(parent: XmlElement | null | undefined, localName: string): XmlElement | undefined {
  return directChildren(parent, localName)[0];
}

function removeDirectChildren(parent: XmlElement, localName: string): void {
  for (const child of directChildren(parent, localName)) parent.removeChild(child);
}

function allElementsByLocalName(document: XmlDocument, localName: string): XmlElement[] {
  return Array.from(document.getElementsByTagName("*")).filter((node) => node.localName === localName);
}

function xmlRoot(document: XmlDocument): XmlElement {
  const root = document.documentElement;
  if (!root) throw tableXmlError("Native table XML has no document element");
  return root;
}

function hasUnsupportedFilterChild(filterColumn: XmlElement): boolean {
  for (let node = filterColumn.firstChild; node; node = node.nextSibling) {
    if (node.nodeType !== 1) continue;
    const name = (node as XmlElement).localName;
    if (name !== "filters" && name !== "customFilters") return true;
  }
  return false;
}

function findColumnByName(table: StructuredTable, name: string): StructuredTableColumn | undefined {
  return table.columns.find((column) => normalizeName(column.name) === normalizeName(name));
}

function firstBodyRow(table: StructuredTable): number {
  return table.range.start.row + (table.headerRow ? 1 : 0);
}

function booleanAttribute(value: boolean | undefined): string {
  return value ? "1" : "0";
}

function normalizeName(value: string): string {
  return value.normalize("NFKC").toLocaleLowerCase("en-US");
}

function ensureLeadingEquals(value: string): string {
  return value.startsWith("=") ? value : `=${value}`;
}

function withoutLeadingEquals(value: string): string {
  return value.startsWith("=") ? value.slice(1) : value;
}

function parseNonNegativeInteger(value: string): number | undefined {
  if (!/^\d+$/.test(value)) return undefined;
  const parsed = Number(value);
  return Number.isSafeInteger(parsed) ? parsed : undefined;
}

function naturalEntryOrder(left: string, right: string): number {
  return left.localeCompare(right, "en", { numeric: true });
}

function assertSafeArchive(data: Uint8Array): void {
  const result = validateXlsxArchive(data);
  if (!result.ok) throw issueError(result.issue);
}

function validatedEntries(data: Uint8Array): ReadonlyMap<string, Uint8Array> {
  const result = extractValidatedXlsxEntries(data);
  if (!result.ok) throw issueError(result.issue);
  return result.entries;
}

function zipEntries(entries: ReadonlyMap<string, Uint8Array>): Uint8Array {
  const zippedEntries: Record<string, [Uint8Array, { mtime: Date; level: 6 }]> = {};
  for (const name of [...entries.keys()].sort()) {
    zippedEntries[name] = [entries.get(name)!, { mtime: FIXED_ZIP_DATE, level: 6 }];
  }
  return zipSync(zippedEntries);
}

function issueError(issue: { code: string; message: string }): Error {
  return Object.assign(new Error(issue.message), { code: issue.code, issue });
}

function tableXmlError(message: string): Error {
  return Object.assign(new Error(message), { code: "XLSX_XML_UNSAFE" });
}

// Kept exported for the XLSX adapter's explicit standard-function mapping.
export function nativeTotalsFunctionToAggregate(value: string): TableAggregate | undefined {
  return NATIVE_TO_AGGREGATE[value];
}
