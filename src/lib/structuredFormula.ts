import type { TableIssue } from "../core/commands/types";
import type { CellRange, StructuredTable, StructuredTableColumn } from "../types";
import { columnIndexToName } from "./addressing";
import { extractFormulaReferences } from "./formulaReferences";

export type StructuredFormulaResult =
  | { ok: true; formula: string }
  | { ok: false; issue: TableIssue };

type StructuredSelector = "thisRow" | "headers" | "totals" | "data" | "all";

type ParsedStructuredReference = {
  selector: StructuredSelector;
  startColumn: StructuredTableColumn;
  endColumn: StructuredTableColumn;
};

type ParsedA1Reference = {
  range: CellRange;
  endIndex: number;
};

const STRUCTURED_ESCAPE_CHARACTERS = new Set(["[", "]", "#", "'", "@"]);

export function structuredFormulaToA1(
  formula: string,
  table: StructuredTable,
  anchorBodyRow: number
): StructuredFormulaResult {
  if (!formula.startsWith("=")) return { ok: true, formula };
  const geometryIssue = validateTableGeometry(table, anchorBodyRow);
  if (geometryIssue) return unsupported(geometryIssue);

  let result = "";
  let cursor = 0;
  while (cursor < formula.length) {
    if (formula[cursor] === '"') {
      const end = skipDoubleQuotedText(formula, cursor);
      result += formula.slice(cursor, end);
      cursor = end;
      continue;
    }

    const candidate = findStructuredCandidate(formula, cursor, table.name);
    if (!candidate) {
      result += formula[cursor];
      cursor += 1;
      continue;
    }
    if (candidate.error) return unsupported(candidate.error);

    const balanced = readBalancedBracket(formula, candidate.bracketIndex);
    if (!balanced) return unsupported("Malformed structured table reference");
    const parsed = parseStructuredReference(balanced.content, table);
    if (typeof parsed === "string") return unsupported(parsed);
    const address = structuredReferenceAddress(parsed, table, anchorBodyRow);
    if (typeof address === "string") {
      result += formula.slice(cursor, candidate.startIndex) + address;
      cursor = balanced.endIndex;
      continue;
    }
    return unsupported(address.message);
  }

  return { ok: true, formula: result };
}

export function a1FormulaToStructured(
  formula: string,
  table: StructuredTable,
  anchorBodyRow: number
): StructuredFormulaResult {
  if (!formula.startsWith("=")) return { ok: true, formula };
  const geometryIssue = validateTableGeometry(table, anchorBodyRow);
  if (geometryIssue) return unsupported(geometryIssue);

  // The foundation tokenizer is the source of truth for which local A1 tokens
  // are formula references. The positional scanner below only performs the
  // replacement; it deliberately does not use one regex to parse formulas.
  const recognized = new Set(extractFormulaReferences(formula).map(({ range }) => rangeKey(range)));
  let result = "";
  let cursor = 0;
  while (cursor < formula.length) {
    if (formula[cursor] === '"') {
      const end = skipDoubleQuotedText(formula, cursor);
      result += formula.slice(cursor, end);
      cursor = end;
      continue;
    }

    const reference = parseLocalA1ReferenceAt(formula, cursor);
    if (!reference || !recognized.has(rangeKey(reference.range))) {
      result += formula[cursor];
      cursor += 1;
      continue;
    }

    const structured = a1RangeToStructured(reference.range, table, anchorBodyRow);
    result += structured ?? formula.slice(cursor, reference.endIndex);
    cursor = reference.endIndex;
  }

  return { ok: true, formula: result };
}

function validateTableGeometry(table: StructuredTable, anchorBodyRow: number): string | undefined {
  if (table.columns.length === 0) return "A structured table must contain at least one column";
  if (!Number.isInteger(anchorBodyRow)) return "The formula anchor must be an integer body row";
  const { bodyStart, bodyEnd } = tableRows(table);
  if (anchorBodyRow < bodyStart || anchorBodyRow > bodyEnd) {
    return "The formula anchor must be inside the table body";
  }
  const seen = new Set<number>();
  for (const column of table.columns) {
    if (
      column.sheetColumn < table.range.start.column ||
      column.sheetColumn > table.range.end.column ||
      seen.has(column.sheetColumn)
    ) {
      return "Table columns must map uniquely inside the table range";
    }
    seen.add(column.sheetColumn);
  }
  return undefined;
}

function findStructuredCandidate(
  formula: string,
  index: number,
  tableName: string
): { startIndex: number; bracketIndex: number; error?: string } | null {
  if (formula[index] === "[") {
    return { startIndex: index, bracketIndex: index };
  }
  if (!isIdentifierStart(formula[index] ?? "")) return null;
  if (index > 0 && isIdentifierContinue(formula[index - 1] ?? "")) return null;

  let cursor = index + 1;
  while (cursor < formula.length && isIdentifierContinue(formula[cursor])) cursor += 1;
  if (formula[cursor] !== "[") return null;
  const qualifier = formula.slice(index, cursor);
  if (normalizeName(qualifier) !== normalizeName(tableName)) {
    return {
      startIndex: index,
      bracketIndex: cursor,
      error: `Structured reference targets unsupported table ${qualifier}`
    };
  }
  return { startIndex: index, bracketIndex: cursor };
}

function parseStructuredReference(
  content: string,
  table: StructuredTable
): ParsedStructuredReference | string {
  const trimmed = content.trim();
  if (!trimmed) return "Empty structured table reference";

  if (trimmed.startsWith("@")) {
    const columnName = parseColumnToken(trimmed.slice(1));
    if (typeof columnName !== "string") return columnName.message;
    return resolveStructuredColumns("thisRow", [columnName], table);
  }

  if (!trimmed.startsWith("[")) {
    const columnName = decodeStructuredHeader(trimmed);
    return resolveStructuredColumns("data", [columnName], table);
  }

  const sequence = parseNestedSpecifierSequence(trimmed);
  if (typeof sequence === "string") return sequence;
  const selectors: StructuredSelector[] = [];
  const columns: string[] = [];
  for (const specifier of sequence.specifiers) {
    const selector = parseSelector(specifier);
    if (selector) selectors.push(selector);
    else columns.push(decodeStructuredHeader(specifier));
  }
  if (selectors.length > 1) return "Combined structured table selectors are unsupported";
  if (columns.length > 2) return "A structured table reference may span at most two columns";
  if (sequence.hasRange && columns.length !== 2) return "Malformed structured table column range";
  if (!sequence.hasRange && columns.length > 1) return "Ambiguous structured table column list";

  const selector = selectors[0] ?? "data";
  return resolveStructuredColumns(selector, columns, table);
}

function parseColumnToken(source: string): string | { message: string } {
  const trimmed = source.trim();
  if (!trimmed) return { message: "Current-row reference is missing a column" };
  if (!trimmed.startsWith("[")) return decodeStructuredHeader(trimmed);
  const bracket = readBalancedBracket(trimmed, 0);
  if (!bracket || bracket.endIndex !== trimmed.length || bracket.content.includes("[")) {
    return { message: "Malformed structured table column" };
  }
  return decodeStructuredHeader(bracket.content);
}

function parseNestedSpecifierSequence(
  source: string
): { specifiers: string[]; hasRange: boolean } | string {
  const specifiers: string[] = [];
  let hasRange = false;
  let cursor = 0;
  let expectSpecifier = true;
  while (cursor < source.length) {
    while (cursor < source.length && isWhitespace(source[cursor])) cursor += 1;
    if (cursor >= source.length) break;
    if (expectSpecifier) {
      if (source[cursor] !== "[") return "Malformed structured table specifier";
      const bracket = readBalancedBracket(source, cursor);
      if (!bracket || bracket.content.includes("[")) return "Malformed structured table specifier";
      specifiers.push(bracket.content.trim());
      cursor = bracket.endIndex;
      expectSpecifier = false;
      continue;
    }

    const delimiter = source[cursor];
    if (delimiter !== "," && delimiter !== ":") return "Unsupported structured table operator";
    if (delimiter === ":") {
      if (hasRange) return "A structured table reference may contain only one column range";
      hasRange = true;
    }
    cursor += 1;
    expectSpecifier = true;
  }
  if (expectSpecifier || specifiers.length === 0) return "Malformed structured table reference";
  return { specifiers, hasRange };
}

function resolveStructuredColumns(
  selector: StructuredSelector,
  columnNames: readonly string[],
  table: StructuredTable
): ParsedStructuredReference | string {
  const columns = columnNames.length === 0
    ? [table.columns[0], table.columns[table.columns.length - 1]]
    : columnNames.map((name) => findTableColumn(table, name));
  if (columns.some((column) => !column)) {
    return `Unknown structured table column ${columnNames.find((_, index) => !columns[index]) ?? ""}`.trim();
  }
  const [startColumn, endColumn = startColumn] = columns as StructuredTableColumn[];
  if (startColumn.sheetColumn > endColumn.sheetColumn) {
    return "Structured table column ranges must be in worksheet order";
  }
  return { selector, startColumn, endColumn };
}

function structuredReferenceAddress(
  reference: ParsedStructuredReference,
  table: StructuredTable,
  anchorBodyRow: number
): string | { message: string } {
  const { bodyStart, bodyEnd } = tableRows(table);
  let startRow: number;
  let endRow: number;
  switch (reference.selector) {
    case "thisRow":
      startRow = endRow = anchorBodyRow;
      break;
    case "headers":
      if (!table.headerRow) return { message: "The table has no header row" };
      startRow = endRow = table.range.start.row;
      break;
    case "totals":
      if (!table.totalsRow) return { message: "The table has no totals row" };
      startRow = endRow = table.range.end.row;
      break;
    case "data":
      startRow = bodyStart;
      endRow = bodyEnd;
      break;
    case "all":
      startRow = table.range.start.row;
      endRow = table.range.end.row;
      break;
  }
  if (startRow > endRow) return { message: "The table has no body rows" };

  const start = formatA1(reference.startColumn.sheetColumn, startRow);
  const end = formatA1(reference.endColumn.sheetColumn, endRow);
  return start === end ? start : `${start}:${end}`;
}

function a1RangeToStructured(
  range: CellRange,
  table: StructuredTable,
  anchorBodyRow: number
): string | undefined {
  const startRow = Math.min(range.start.row, range.end.row);
  const endRow = Math.max(range.start.row, range.end.row);
  const startColumnIndex = Math.min(range.start.column, range.end.column);
  const endColumnIndex = Math.max(range.start.column, range.end.column);
  const startColumn = table.columns.find((column) => column.sheetColumn === startColumnIndex);
  const endColumn = table.columns.find((column) => column.sheetColumn === endColumnIndex);
  if (!startColumn || !endColumn) return undefined;

  const { bodyStart, bodyEnd } = tableRows(table);
  let selector: StructuredSelector | undefined;
  if (startRow === anchorBodyRow && endRow === anchorBodyRow) selector = "thisRow";
  else if (table.headerRow && startRow === table.range.start.row && endRow === startRow) selector = "headers";
  else if (table.totalsRow && startRow === table.range.end.row && endRow === startRow) selector = "totals";
  else if (startRow === bodyStart && endRow === bodyEnd) selector = "data";
  else if (startRow === table.range.start.row && endRow === table.range.end.row) selector = "all";
  if (!selector) return undefined;

  return formatStructuredReference(table.name, selector, startColumn, endColumn);
}

function formatStructuredReference(
  tableName: string,
  selector: StructuredSelector,
  startColumn: StructuredTableColumn,
  endColumn: StructuredTableColumn
): string {
  const start = formatStructuredColumn(startColumn.name);
  const end = formatStructuredColumn(endColumn.name);
  if (selector === "thisRow" && startColumn.id === endColumn.id) {
    return `${tableName}[@${requiresNestedColumnBrackets(startColumn.name) ? `[${start}]` : start}]`;
  }
  const selectorText = {
    headers: "#Headers",
    totals: "#Totals",
    data: "#Data",
    all: "#All",
    thisRow: "#This Row"
  }[selector];
  const columns = startColumn.id === endColumn.id ? `[${start}]` : `[${start}]:[${end}]`;
  return `${tableName}[[${selectorText}],${columns}]`;
}

function formatStructuredColumn(name: string): string {
  return [...name].map((character) =>
    STRUCTURED_ESCAPE_CHARACTERS.has(character) ? `'${character}` : character
  ).join("");
}

function requiresNestedColumnBrackets(name: string): boolean {
  return /[\t\n\r,:.\[\]#'"{}$^&*+=\-><\/@\\!()%?`;~_ ]/.test(name);
}

function parseSelector(source: string): StructuredSelector | undefined {
  const normalized = source.trim().replace(/\s+/g, " ").toLowerCase();
  if (normalized === "#this row") return "thisRow";
  if (normalized === "#headers") return "headers";
  if (normalized === "#totals") return "totals";
  if (normalized === "#data") return "data";
  if (normalized === "#all") return "all";
  return undefined;
}

function decodeStructuredHeader(source: string): string {
  let result = "";
  for (let index = 0; index < source.length; index += 1) {
    if (source[index] === "'" && STRUCTURED_ESCAPE_CHARACTERS.has(source[index + 1] ?? "")) {
      result += source[index + 1];
      index += 1;
    } else {
      result += source[index];
    }
  }
  return result;
}

function findTableColumn(table: StructuredTable, name: string): StructuredTableColumn | undefined {
  const exact = table.columns.find((column) => column.name === name);
  return exact ?? table.columns.find((column) => normalizeName(column.name) === normalizeName(name));
}

function parseLocalA1ReferenceAt(formula: string, index: number): ParsedA1Reference | null {
  const previous = formula[index - 1] ?? "";
  if (isIdentifierContinue(previous) || previous === "!" || previous === ":") return null;
  const start = parseA1CellAt(formula, index);
  if (!start) return null;
  let end = start;
  let cursor = start.endIndex;
  if (formula[cursor] === ":") {
    const candidate = parseA1CellAt(formula, cursor + 1);
    if (candidate) {
      end = candidate;
      cursor = candidate.endIndex;
    }
  }
  const next = formula[cursor] ?? "";
  if (isIdentifierContinue(next) || next === "(" || next === "[" || next === "!") return null;
  return {
    range: {
      start: { row: start.row, column: start.column },
      end: { row: end.row, column: end.column }
    },
    endIndex: cursor
  };
}

function parseA1CellAt(
  formula: string,
  index: number
): { row: number; column: number; endIndex: number } | null {
  let cursor = index;
  if (formula[cursor] === "$") cursor += 1;
  const columnStart = cursor;
  while (isAsciiLetter(formula[cursor] ?? "")) cursor += 1;
  if (cursor === columnStart) return null;
  const columnName = formula.slice(columnStart, cursor).toUpperCase();
  if (formula[cursor] === "$") cursor += 1;
  const rowStart = cursor;
  while (isAsciiDigit(formula[cursor] ?? "")) cursor += 1;
  if (cursor === rowStart || formula[rowStart] === "0") return null;
  const column = columnNameToIndexSafely(columnName);
  if (column === undefined || column > 16_383) return null;
  return { row: Number(formula.slice(rowStart, cursor)) - 1, column, endIndex: cursor };
}

function columnNameToIndexSafely(name: string): number | undefined {
  let value = 0;
  for (const character of name) {
    const code = character.charCodeAt(0);
    if (code < 65 || code > 90) return undefined;
    value = value * 26 + code - 64;
  }
  return value - 1;
}

function readBalancedBracket(
  source: string,
  startIndex: number
): { content: string; endIndex: number } | null {
  if (source[startIndex] !== "[") return null;
  let depth = 1;
  let cursor = startIndex + 1;
  while (cursor < source.length) {
    const character = source[cursor];
    if (character === "'" && STRUCTURED_ESCAPE_CHARACTERS.has(source[cursor + 1] ?? "")) {
      cursor += 2;
      continue;
    }
    if (character === "[") depth += 1;
    else if (character === "]") {
      depth -= 1;
      if (depth === 0) {
        return { content: source.slice(startIndex + 1, cursor), endIndex: cursor + 1 };
      }
    }
    cursor += 1;
  }
  return null;
}

function skipDoubleQuotedText(formula: string, startIndex: number): number {
  let cursor = startIndex + 1;
  while (cursor < formula.length) {
    if (formula[cursor] !== '"') {
      cursor += 1;
      continue;
    }
    if (formula[cursor + 1] === '"') {
      cursor += 2;
      continue;
    }
    return cursor + 1;
  }
  return formula.length;
}

function tableRows(table: StructuredTable): { bodyStart: number; bodyEnd: number } {
  return {
    bodyStart: table.range.start.row + (table.headerRow ? 1 : 0),
    bodyEnd: table.range.end.row - (table.totalsRow ? 1 : 0)
  };
}

function formatA1(column: number, row: number): string {
  return `${columnIndexToName(column)}${row + 1}`;
}

function rangeKey(range: CellRange): string {
  const startRow = Math.min(range.start.row, range.end.row);
  const endRow = Math.max(range.start.row, range.end.row);
  const startColumn = Math.min(range.start.column, range.end.column);
  const endColumn = Math.max(range.start.column, range.end.column);
  return `${startRow}:${startColumn}:${endRow}:${endColumn}`;
}

function normalizeName(value: string): string {
  return value.normalize("NFKC").toLocaleLowerCase("en-US");
}

function isIdentifierStart(character: string): boolean {
  return /^[\p{ID_Start}_]$/u.test(character);
}

function isIdentifierContinue(character: string): boolean {
  return /^[\p{ID_Continue}_.]$/u.test(character);
}

function isAsciiLetter(character: string): boolean {
  const code = character.charCodeAt(0);
  return (code >= 65 && code <= 90) || (code >= 97 && code <= 122);
}

function isAsciiDigit(character: string): boolean {
  const code = character.charCodeAt(0);
  return code >= 48 && code <= 57;
}

function isWhitespace(character: string): boolean {
  return /\s/u.test(character);
}

function unsupported(message: string): StructuredFormulaResult {
  return {
    ok: false,
    issue: {
      code: "TABLE_FORMULA_REFERENCE_UNSUPPORTED",
      message
    }
  };
}
