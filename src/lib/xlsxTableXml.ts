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
import { formatCellAddress, parseCellAddress } from "./addressing";
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
const COMMENT_ENTRY_PATTERN = /^xl\/comments(?:\/[^/]+|[^/]*)\.xml$/i;
const DRAWING_ENTRY_PATTERN = /^xl\/(?:drawings|charts)\//i;
const WORKSHEET_ENTRY_PATTERN = /^xl\/worksheets\/[^/]+\.xml$/i;
const RELATIONSHIP_ENTRY_PATTERN = /\.rels$/i;
const VBA_ENTRY_PATTERN = /^xl\/(?:vbaProject(?:Signature)?\.bin|_rels\/vbaProject(?:Signature)?\.bin\.rels)$/i;
const FIXED_ZIP_DATE = new Date("1980-01-01T00:00:00.000Z");
const XML_NAMESPACE = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const EXCEL_MAX_ROWS = 1_048_576;
const EXCEL_MAX_COLUMNS = 16_384;
const XLSX_WORKBOOK_CONTENT_TYPE =
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";
const XLSM_WORKBOOK_CONTENT_TYPE =
  "application/vnd.ms-excel.sheet.macroEnabled.main+xml";
const DRAWING_RELATIONSHIP_TYPES = new Set([
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing",
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing",
  "http://purl.oclc.org/ooxml/officeDocument/relationships/drawing",
  "http://purl.oclc.org/ooxml/officeDocument/relationships/vmlDrawing"
]);
const COMMENT_RELATIONSHIP_TYPES = new Set([
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments",
  "http://purl.oclc.org/ooxml/officeDocument/relationships/comments"
]);
const TABLE_RELATIONSHIP_TYPES = new Set([
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table",
  "http://purl.oclc.org/ooxml/officeDocument/relationships/table"
]);
const VBA_RELATIONSHIP_TYPES = new Set([
  "http://schemas.microsoft.com/office/2006/relationships/vbaProject",
  "http://schemas.microsoft.com/office/2006/relationships/vbaProjectSignature"
]);

export type NativeWorksheetComments = {
  sheetName: string;
  comments: Readonly<Record<string, string>>;
};

export type PreparedXlsxImport = {
  tableMetadata: readonly NativeTableXmlMetadata[];
  comments: readonly NativeWorksheetComments[];
  excelJsBytes: Uint8Array;
};

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
  return readNativeTableEntries(validatedEntries(data));
}

function readNativeTableEntries(
  entries: ReadonlyMap<string, Uint8Array>
): readonly NativeTableXmlMetadata[] {
  return [...entries.keys()]
    .filter((name) => TABLE_ENTRY_PATTERN.test(name))
    .sort(naturalEntryOrder)
    .map((name) => readTableDocument(strFromU8(entries.get(name)!)));
}

export function readNativeCommentsXml(data: Uint8Array): readonly NativeWorksheetComments[] {
  return readNativeCommentEntries(validatedEntries(data));
}

function readNativeCommentEntries(
  entries: ReadonlyMap<string, Uint8Array>
): readonly NativeWorksheetComments[] {
  const workbookXml = entries.get("xl/workbook.xml");
  const workbookRelationshipsXml = entries.get("xl/_rels/workbook.xml.rels");
  if (!workbookXml || !workbookRelationshipsXml) return [];

  const workbook = parseXml(strFromU8(workbookXml));
  const workbookRelationships = relationshipsById(parseXml(strFromU8(workbookRelationshipsXml)));
  const relationshipsByPart = new Map<string, ReadonlyMap<string, NativeRelationship>>();
  const commentsByPart = new Map<string, Readonly<Record<string, string>>>();
  const commentsByWorksheetPart = new Map<string, Readonly<Record<string, string>>>();
  const worksheets: NativeWorksheetComments[] = [];
  for (const sheet of allElementsByLocalName(workbook, "sheet")) {
    const sheetName = sheet.getAttribute("name");
    const relationship = workbookRelationships.get(attributeByLocalName(sheet, "id"));
    const worksheetPart = relationship
      ? resolveRelationshipTarget("xl/workbook.xml", relationship.target)
      : undefined;
    if (!sheetName || !worksheetPart) continue;

    let comments = commentsByWorksheetPart.get(worksheetPart);
    if (!comments) {
      const relationshipPart = relationshipPartName(worksheetPart);
      const relationshipXml = entries.get(relationshipPart);
      const relationships = relationshipsByPart.get(relationshipPart)
        ?? (relationshipXml ? relationshipsById(parseXml(strFromU8(relationshipXml))) : new Map());
      relationshipsByPart.set(relationshipPart, relationships);

      const parsedComments: Record<string, string> = {};
      for (const candidate of relationships.values()) {
        if (
          !COMMENT_RELATIONSHIP_TYPES.has(candidate.type)
          || candidate.targetMode.toLowerCase() === "external"
        ) continue;
        const commentPart = resolveRelationshipTarget(worksheetPart, candidate.target);
        if (!commentPart || !COMMENT_ENTRY_PATTERN.test(commentPart)) continue;
        const commentXml = entries.get(commentPart);
        if (!commentXml) continue;
        let commentMap = commentsByPart.get(commentPart);
        if (!commentMap) {
          commentMap = readCommentDocument(strFromU8(commentXml));
          commentsByPart.set(commentPart, commentMap);
        }
        Object.assign(parsedComments, commentMap);
      }
      comments = parsedComments;
      commentsByWorksheetPart.set(worksheetPart, comments);
    }
    worksheets.push({ sheetName, comments });
  }
  return worksheets;
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
  return prepareEntriesForExcelJs(validatedEntries(data));
}

export function prepareXlsxImportForExcelJs(data: Uint8Array): PreparedXlsxImport {
  const entries = validatedEntries(data);
  return {
    tableMetadata: readNativeTableEntries(entries),
    comments: readNativeCommentEntries(entries),
    excelJsBytes: prepareEntriesForExcelJs(entries)
  };
}

function prepareEntriesForExcelJs(entriesInput: ReadonlyMap<string, Uint8Array>): Uint8Array {
  const entries = new Map(entriesInput);
  removeUnsupportedDrawingParts(entries);
  removeCommentParts(entries);
  removeVbaProjectParts(entries);
  canonicalizeTableRelationshipTargets(entries);
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

function removeVbaProjectParts(entries: Map<string, Uint8Array>): void {
  const removedParts = new Set(
    [...entries.keys()].filter((name) => VBA_ENTRY_PATTERN.test(name))
  );

  for (const entryName of [...entries.keys()].filter((name) => RELATIONSHIP_ENTRY_PATTERN.test(name))) {
    const sourcePart = sourcePartForRelationships(entryName);
    const document = parseXml(strFromU8(entries.get(entryName)!));
    for (const relationship of allElementsByLocalName(document, "Relationship")) {
      const type = relationship.getAttribute("Type") ?? "";
      if (!VBA_RELATIONSHIP_TYPES.has(type)) continue;
      const target = sourcePart
        ? resolveRelationshipTarget(sourcePart, relationship.getAttribute("Target") ?? "")
        : undefined;
      if (target) {
        removedParts.add(target);
        removedParts.add(relationshipPartName(target));
      }
      relationship.parentNode?.removeChild(relationship);
    }
    entries.set(entryName, strToU8(new XMLSerializer().serializeToString(document)));
  }
  for (const partName of removedParts) entries.delete(partName);

  const contentTypes = entries.get("[Content_Types].xml");
  if (!contentTypes) return;
  const document = parseXml(strFromU8(contentTypes));
  for (const override of allElementsByLocalName(document, "Override")) {
    const partName = override.getAttribute("PartName")?.replace(/^\//, "") ?? "";
    if (removedParts.has(partName) || VBA_ENTRY_PATTERN.test(partName)) {
      override.parentNode?.removeChild(override);
    } else if (override.getAttribute("ContentType") === XLSM_WORKBOOK_CONTENT_TYPE) {
      override.setAttribute("ContentType", XLSX_WORKBOOK_CONTENT_TYPE);
    }
  }
  entries.set("[Content_Types].xml", strToU8(new XMLSerializer().serializeToString(document)));
}

function canonicalizeTableRelationshipTargets(entries: Map<string, Uint8Array>): void {
  for (const entryName of [...entries.keys()].filter((name) => (
    /^xl\/worksheets\/_rels\/[^/]+\.xml\.rels$/i.test(name)
  ))) {
    const sourcePart = sourcePartForRelationships(entryName);
    if (!sourcePart) continue;
    const document = parseXml(strFromU8(entries.get(entryName)!));
    for (const relationship of allElementsByLocalName(document, "Relationship")) {
      if (!TABLE_RELATIONSHIP_TYPES.has(relationship.getAttribute("Type") ?? "")) continue;
      const target = resolveRelationshipTarget(sourcePart, relationship.getAttribute("Target") ?? "");
      const tableName = target?.match(/^xl\/tables\/([^/]+\.xml)$/i)?.[1];
      if (tableName) relationship.setAttribute("Target", `../tables/${tableName}`);
    }
    entries.set(entryName, strToU8(new XMLSerializer().serializeToString(document)));
  }
}

function removeCommentParts(entries: Map<string, Uint8Array>): void {
  for (const entryName of [...entries.keys()]) {
    if (COMMENT_ENTRY_PATTERN.test(entryName)) entries.delete(entryName);
  }

  for (const entryName of [...entries.keys()].filter((name) => RELATIONSHIP_ENTRY_PATTERN.test(name))) {
    const document = parseXml(strFromU8(entries.get(entryName)!));
    for (const relationship of allElementsByLocalName(document, "Relationship")) {
      if (COMMENT_RELATIONSHIP_TYPES.has(relationship.getAttribute("Type") ?? "")) {
        relationship.parentNode?.removeChild(relationship);
      }
    }
    entries.set(entryName, strToU8(new XMLSerializer().serializeToString(document)));
  }

  const contentTypes = entries.get("[Content_Types].xml");
  if (!contentTypes) return;
  const document = parseXml(strFromU8(contentTypes));
  for (const override of allElementsByLocalName(document, "Override")) {
    const partName = override.getAttribute("PartName")?.replace(/^\//, "") ?? "";
    if (COMMENT_ENTRY_PATTERN.test(partName)) override.parentNode?.removeChild(override);
  }
  entries.set("[Content_Types].xml", strToU8(new XMLSerializer().serializeToString(document)));
}

function readCommentDocument(xml: string): Record<string, string> {
  const document = parseXml(xml);
  const comments: Record<string, string> = {};
  for (const comment of allElementsByLocalName(document, "comment")) {
    const address = normalizeCommentAddress(comment.getAttribute("ref") ?? "");
    if (!address) continue;
    const text = firstDirectChild(comment, "text");
    comments[address] = text
      ? allDescendantsByLocalName(text, "t").map((node) => node.textContent ?? "").join("")
      : "";
  }
  return comments;
}

function normalizeCommentAddress(value: string): string | undefined {
  try {
    const coordinate = parseCellAddress(value.replace(/\$/g, "").toUpperCase());
    if (
      !Number.isSafeInteger(coordinate.row)
      || !Number.isSafeInteger(coordinate.column)
      || coordinate.row < 0
      || coordinate.column < 0
      || coordinate.row >= EXCEL_MAX_ROWS
      || coordinate.column >= EXCEL_MAX_COLUMNS
    ) return undefined;
    return formatCellAddress(coordinate);
  } catch {
    return undefined;
  }
}

type NativeRelationship = {
  target: string;
  targetMode: string;
  type: string;
};

function relationshipsById(document: XmlDocument): ReadonlyMap<string, NativeRelationship> {
  const relationships = new Map<string, NativeRelationship>();
  for (const relationship of allElementsByLocalName(document, "Relationship")) {
    const id = relationship.getAttribute("Id");
    const target = relationship.getAttribute("Target");
    if (!id || !target) continue;
    relationships.set(id, {
      target,
      targetMode: relationship.getAttribute("TargetMode") ?? "",
      type: relationship.getAttribute("Type") ?? ""
    });
  }
  return relationships;
}

function attributeByLocalName(element: XmlElement, localName: string): string {
  for (let index = 0; index < element.attributes.length; index += 1) {
    const attribute = element.attributes.item(index);
    if (attribute?.localName === localName) return attribute.value;
  }
  return "";
}

function relationshipPartName(partName: string): string {
  const segments = partName.split("/");
  const fileName = segments.pop() ?? "";
  return [...segments, "_rels", `${fileName}.rels`].join("/");
}

function sourcePartForRelationships(relationshipPartName: string): string | undefined {
  const marker = "/_rels/";
  const markerOffset = relationshipPartName.indexOf(marker);
  if (markerOffset < 0 || !relationshipPartName.endsWith(".rels")) return undefined;
  const prefix = relationshipPartName.slice(0, markerOffset);
  const relatedName = relationshipPartName.slice(markerOffset + marker.length, -".rels".length);
  return relatedName && !relatedName.includes("/") ? `${prefix}/${relatedName}` : undefined;
}

function resolveRelationshipTarget(sourcePart: string, target: string): string | undefined {
  if (target.includes("?") || target.includes("#") || target.includes("\\") || target.includes("\0")) {
    return undefined;
  }
  let decoded: string;
  try {
    decoded = decodeURIComponent(target);
  } catch {
    return undefined;
  }
  if (
    decoded.includes("?")
    || decoded.includes("#")
    || decoded.includes("\\")
    || decoded.includes("\0")
    || /^[a-z][a-z\d+.-]*:/i.test(decoded)
    || decoded.startsWith("//")
  ) return undefined;
  const segments = decoded.startsWith("/") ? [] : sourcePart.split("/").slice(0, -1);
  for (const segment of decoded.replace(/^\//, "").split("/")) {
    if (!segment || segment === ".") continue;
    if (segment === "..") {
      if (segments.length === 0) return undefined;
      segments.pop();
    } else if (segment.includes(":")) return undefined;
    else segments.push(segment);
  }
  return segments.length > 0 ? segments.join("/") : undefined;
}

function removeUnsupportedDrawingParts(entries: Map<string, Uint8Array>): void {
  for (const entryName of [...entries.keys()]) {
    if (DRAWING_ENTRY_PATTERN.test(entryName)) entries.delete(entryName);
  }

  for (const entryName of [...entries.keys()].filter((name) => WORKSHEET_ENTRY_PATTERN.test(name))) {
    const document = parseXml(strFromU8(entries.get(entryName)!));
    removeElementsByLocalName(document, new Set(["drawing", "legacyDrawing", "legacyDrawingHF"]));
    entries.set(entryName, strToU8(new XMLSerializer().serializeToString(document)));
  }

  for (const entryName of [...entries.keys()].filter((name) => RELATIONSHIP_ENTRY_PATTERN.test(name))) {
    const sourcePart = sourcePartForRelationships(entryName);
    const document = parseXml(strFromU8(entries.get(entryName)!));
    for (const relationship of allElementsByLocalName(document, "Relationship")) {
      const type = relationship.getAttribute("Type") ?? "";
      const target = relationship.getAttribute("Target") ?? "";
      const targetMode = relationship.getAttribute("TargetMode") ?? "";
      const resolvedTarget = sourcePart && targetMode.toLowerCase() !== "external"
        ? resolveRelationshipTarget(sourcePart, target)
        : undefined;
      if (
        DRAWING_RELATIONSHIP_TYPES.has(type)
        || (resolvedTarget !== undefined && DRAWING_ENTRY_PATTERN.test(resolvedTarget))
      ) {
        relationship.parentNode?.removeChild(relationship);
      }
    }
    entries.set(entryName, strToU8(new XMLSerializer().serializeToString(document)));
  }

  const contentTypes = entries.get("[Content_Types].xml");
  if (contentTypes) {
    const document = parseXml(strFromU8(contentTypes));
    for (const override of allElementsByLocalName(document, "Override")) {
      const partName = override.getAttribute("PartName")?.replace(/^\//, "") ?? "";
      if (DRAWING_ENTRY_PATTERN.test(partName)) override.parentNode?.removeChild(override);
    }
    entries.set(
      "[Content_Types].xml",
      strToU8(new XMLSerializer().serializeToString(document))
    );
  }
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
      const nativeOperator = custom.getAttribute("operator") ?? "equal";
      const operator = nativeComparisonOperator(nativeOperator);
      if (!operator) continue;
      const criterion = decodeExcelFilterCriterion(
        operator,
        custom.getAttribute("val") ?? ""
      );
      customExpressions.push({
        kind: "comparison" as const,
        columnId,
        operator: criterion.operator,
        value: { type: "string" as const, value: criterion.value }
      });
    }
    if (
      customExpressions.length > 1
      && customFilters.getAttribute("and") === "1"
      && customExpressions.every((expression) => expression.operator === "neq")
    ) {
      expressions.push({
        kind: "set",
        columnId,
        operator: "notIn",
        values: customExpressions.map((expression) => expression.value)
      });
    } else if (customExpressions.length === 1) expressions.push(customExpressions[0]);
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

  const columnFilters = nativeColumnFilters(filter);
  const filtersByColumn = new Map<number, NativeColumnFilter>();
  for (const columnFilter of columnFilters) {
    const columnIndex = table.columns.findIndex((column) => column.id === columnFilter.columnId);
    if (columnIndex < 0) throw tableXmlError("Native table filter refers to an unknown column");
    if (filtersByColumn.has(columnIndex)) {
      throw tableXmlError("Unsupported native table filter has multiple expressions for one column");
    }
    filtersByColumn.set(columnIndex, columnFilter);
  }

  for (let columnIndex = 0; columnIndex < table.columns.length; columnIndex += 1) {
    const columnFilter = filtersByColumn.get(columnIndex);
    const candidates = existingByIndex.get(columnIndex) ?? [];
    const filterColumn = candidates.find(hasUnsupportedFilterChild)
      ?? candidates[0]
      ?? document.createElementNS(root.namespaceURI || XML_NAMESPACE, "filterColumn");
    filterColumn.setAttribute("colId", String(columnIndex));
    if (!columnFilter) {
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
    if (columnFilter.kind === "set") {
      const filters = document.createElementNS(root.namespaceURI || XML_NAMESPACE, "filters");
      for (const value of columnFilter.values) {
        const filterNode = document.createElementNS(root.namespaceURI || XML_NAMESPACE, "filter");
        filterNode.setAttribute("val", scalarText(value));
        filters.appendChild(filterNode);
      }
      filterColumn.appendChild(filters);
    } else {
      const customFilters = document.createElementNS(root.namespaceURI || XML_NAMESPACE, "customFilters");
      if (columnFilter.comparisons.length > 1) {
        customFilters.setAttribute("and", columnFilter.operator === "and" ? "1" : "0");
      }
      for (const comparison of columnFilter.comparisons) {
        const operator = appComparisonOperator(comparison.operator);
        if (!operator) throw tableXmlError("Unsupported native table filter comparison operator");
        const custom = document.createElementNS(root.namespaceURI || XML_NAMESPACE, "customFilter");
        if (operator !== "equal") custom.setAttribute("operator", operator);
        custom.setAttribute("val", excelFilterCriterion(comparison));
        customFilters.appendChild(custom);
      }
      filterColumn.appendChild(customFilters);
    }
    autoFilter.appendChild(filterColumn);
  }
}

type NativeComparisonFilter = Extract<FilterExpression, { kind: "comparison" }>;
type NativeColumnFilter =
  | Extract<FilterExpression, { kind: "set" }>
  | {
      kind: "custom";
      columnId: string;
      operator: "and" | "or";
      comparisons: readonly NativeComparisonFilter[];
    };

function nativeColumnFilters(filter: FilterExpression): NativeColumnFilter[] {
  if (filter.kind === "set") {
    if (filter.operator === "in") return [filter];
    if (filter.values.length === 0 || filter.values.length > 2) {
      throw tableXmlError("Native table notIn filters require one or two values");
    }
    return [{
      kind: "custom",
      columnId: filter.columnId,
      operator: "and",
      comparisons: filter.values.map((value) => ({
        kind: "comparison",
        columnId: filter.columnId,
        operator: "neq",
        value
      }))
    }];
  }
  if (filter.kind === "comparison") {
    return [{ kind: "custom", columnId: filter.columnId, operator: "and", comparisons: [filter] }];
  }
  if (filter.kind === "range") {
    const between = filter.operator === "between";
    return [{
      kind: "custom",
      columnId: filter.columnId,
      operator: between ? "and" : "or",
      comparisons: [
        {
          kind: "comparison",
          columnId: filter.columnId,
          operator: between ? "gte" : "lt",
          value: filter.lower
        },
        {
          kind: "comparison",
          columnId: filter.columnId,
          operator: between ? "lte" : "gt",
          value: filter.upper
        }
      ]
    }];
  }
  if (filter.kind === "logical" && filter.operator === "or") {
    if (filter.operands.length === 0 || filter.operands.length > 2) {
      throw tableXmlError("Unsupported native table filter OR expression");
    }
    const comparisons = filter.operands.map((operand) => {
      if (operand.kind !== "comparison") {
        throw tableXmlError("Unsupported native table filter OR expression");
      }
      return operand;
    });
    const columnId = comparisons[0].columnId;
    if (comparisons.some((comparison) => comparison.columnId !== columnId)) {
      throw tableXmlError("Unsupported native table filter OR spans multiple columns");
    }
    return [{ kind: "custom", columnId, operator: "or", comparisons }];
  }
  if (filter.kind === "logical" && filter.operator === "and") {
    if (filter.operands.length === 0) throw tableXmlError("Unsupported native table filter expression");
    return mergeAndColumnFilters(filter.operands.flatMap(nativeColumnFilters));
  }
  throw tableXmlError("Unsupported native table filter expression");
}

function mergeAndColumnFilters(filters: readonly NativeColumnFilter[]): NativeColumnFilter[] {
  const merged = new Map<string, NativeColumnFilter>();
  const order: string[] = [];
  for (const filter of filters) {
    const existing = merged.get(filter.columnId);
    if (!existing) {
      merged.set(filter.columnId, filter);
      order.push(filter.columnId);
      continue;
    }
    if (
      existing.kind !== "custom"
      || filter.kind !== "custom"
      || existing.comparisons.length !== 1
      || filter.comparisons.length !== 1
    ) {
      throw tableXmlError("Unsupported native table filter has multiple expressions for one column");
    }
    merged.set(filter.columnId, {
      kind: "custom",
      columnId: filter.columnId,
      operator: "and",
      comparisons: [...existing.comparisons, ...filter.comparisons]
    });
  }
  return order.map((columnId) => merged.get(columnId)!);
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
    contains: "equal",
    startsWith: "equal",
    endsWith: "equal",
    gt: "greaterThan",
    gte: "greaterThanOrEqual",
    lt: "lessThan",
    lte: "lessThanOrEqual"
  } as Partial<Record<typeof operator, string>>)[operator];
}

function excelFilterCriterion(comparison: NativeComparisonFilter): string {
  const value = scalarText(comparison.value);
  if (comparison.value.type !== "string") return value;
  const escaped = escapeExcelFilterWildcards(value);
  if (comparison.operator === "contains") return `*${escaped}*`;
  if (comparison.operator === "startsWith") return `${escaped}*`;
  if (comparison.operator === "endsWith") return `*${escaped}`;
  return escaped;
}

function decodeExcelFilterCriterion(
  operator: NativeComparisonFilter["operator"],
  criterion: string
): { operator: NativeComparisonFilter["operator"]; value: string } {
  if (operator !== "eq") {
    return { operator, value: unescapeExcelFilterWildcards(criterion) };
  }
  const leadingWildcard = criterion.startsWith("*");
  const trailingIndex = criterion.length - 1;
  const trailingWildcard = trailingIndex >= 0
    && criterion[trailingIndex] === "*"
    && !excelFilterCharacterIsEscaped(criterion, trailingIndex);
  let value = criterion;
  if (leadingWildcard) value = value.slice(1);
  if (trailingWildcard && value.length > 0) value = value.slice(0, -1);
  return {
    operator: leadingWildcard && trailingWildcard
      ? "contains"
      : trailingWildcard
        ? "startsWith"
        : leadingWildcard
          ? "endsWith"
          : operator,
    value: unescapeExcelFilterWildcards(value)
  };
}

function escapeExcelFilterWildcards(value: string): string {
  return value.replace(/[~*?]/g, (character) => `~${character}`);
}

function unescapeExcelFilterWildcards(value: string): string {
  return value.replace(/~([~*?])/g, "$1");
}

function excelFilterCharacterIsEscaped(value: string, index: number): boolean {
  let tildes = 0;
  for (let cursor = index - 1; cursor >= 0 && value[cursor] === "~"; cursor -= 1) {
    tildes += 1;
  }
  return tildes % 2 === 1;
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

function allDescendantsByLocalName(element: XmlElement, localName: string): XmlElement[] {
  return Array.from(element.getElementsByTagName("*")).filter((node) => node.localName === localName);
}

function removeElementsByLocalName(document: XmlDocument, localNames: ReadonlySet<string>): void {
  for (const element of Array.from(document.getElementsByTagName("*"))) {
    if (localNames.has(element.localName ?? "")) element.parentNode?.removeChild(element);
  }
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
