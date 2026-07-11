import type { CellRange } from "../types";
import type { TableIssue } from "../core/commands/types";
import { columnIndexToName, columnNameToIndex, formatCellAddress, normalizeRange } from "./addressing";

export type FormulaReferenceOffset = {
  rowOffset: number;
  columnOffset: number;
};

export type BoundedFormulaRowTranslation = {
  rowOffset: number;
  formulaSheetName: string;
  columnStart: number;
  columnEnd: number;
};

export type ExtractedFormulaReference = {
  label: string;
  range: CellRange;
};

export type FormulaStructureContext = {
  formulaSheetName: string;
  editedSheetName: string;
  axis: "row" | "column";
  mode: "insert" | "delete";
  index: number;
  count: number;
};

export type RectangularRowEditContext = {
  formulaSheetId: string;
  editedSheetId: string;
  tableColumnStart: number;
  tableColumnEnd: number;
  tableRowEnd: number;
  row: number;
  count: number;
  operation: "insert" | "delete";
  sheetBounds: { rowCount: number; columnCount: number };
};

export type FormulaRewriteResult =
  | { ok: true; formula: string }
  | { ok: false; issue: TableIssue };

type ParsedFormulaCellReference = {
  columnLock: string;
  rowLock: string;
  row: number;
  column: number;
};

type FormulaRawToken = {
  kind: "raw";
  raw: string;
};

type FormulaReferenceToken = {
  kind: "reference";
  raw: string;
  sheetName?: string;
  sheetPrefix: string;
  start: ParsedFormulaCellReference;
  end?: ParsedFormulaCellReference;
};

type FormulaToken = FormulaRawToken | FormulaReferenceToken;

type ParsedFormulaReferenceToken = FormulaReferenceToken & {
  endIndex: number;
};

const CELL_REFERENCE_PATTERN = /(?<![A-Z0-9_.])(\$?)([A-Z]+)(\$?)([1-9]\d*)/gi;
const LOCAL_REFERENCE_PATTERN = /(?<![A-Z0-9_.])(\$?[A-Z]+\$?[1-9]\d*)(?::(\$?[A-Z]+\$?[1-9]\d*))?/gi;
const FORMULA_IDENTIFIER_START_PATTERN = /^[\p{ID_Start}_]$/u;
const FORMULA_IDENTIFIER_CONTINUE_PATTERN = /^[\p{ID_Continue}_.]$/u;
const EXTERNAL_WORKBOOK_QUALIFIER_PATTERN = /(?:'(?:[^']|'')*\[[^\]]+\](?:[^']|'')*'|\[[^\]]+\][\p{ID_Start}_][\p{ID_Continue}_.]*)!/u;
const MAX_FORMULA_COLUMN_INDEX = columnNameToIndex("XFD");

export function rewriteFormulaForStructure(formula: string, context: FormulaStructureContext): string {
  if (!formula.startsWith("=") || context.count <= 0) {
    return formula;
  }

  return tokenizeFormulaReferences(formula)
    .map((token) =>
      token.kind === "reference" && referenceTargetsEditedSheet(token, context)
        ? shiftReferenceToken(token, context)
        : token.raw
    )
    .join("");
}

export function rewriteFormulaForRectangularRowEdit(
  formula: string,
  context: RectangularRowEditContext
): FormulaRewriteResult {
  if (!formula.startsWith("=") || context.count <= 0) return { ok: true, formula };
  if (containsUnsupportedRectangularReference(formula)) {
    return {
      ok: false,
      issue: {
        code: "TABLE_FORMULA_REFERENCE_UNSUPPORTED",
        message: "External and 3-D references cannot be rewritten for a table row edit"
      }
    };
  }

  const tokens = tokenizeFormulaReferences(formula);
  const rewritten: string[] = [];
  for (const token of tokens) {
    if (token.kind === "raw" || !rectangularReferenceTargetsEditedSheet(token, context)) {
      rewritten.push(token.raw);
      continue;
    }
    const result = rewriteRectangularReferenceToken(token, context);
    if (!result.ok) return result;
    rewritten.push(result.formula);
  }
  return { ok: true, formula: rewritten.join("") };
}

function containsUnsupportedRectangularReference(formula: string): boolean {
  let unsupported = false;
  transformUnquotedFormulaSegments(formula, (segment) => {
    if (
      EXTERNAL_WORKBOOK_QUALIFIER_PATTERN.test(segment)
      || /(?:'(?:[^']|'')+'|[\p{ID_Start}_][\p{ID_Continue}_.]*):(?:'(?:[^']|'')+'|[\p{ID_Start}_][\p{ID_Continue}_.]*)!/u.test(segment)
    ) {
      unsupported = true;
    }
    return segment;
  });
  return unsupported;
}

export function translateFormulaReferences(content: string, offset: FormulaReferenceOffset): string {
  if (!content.startsWith("=") || (offset.rowOffset === 0 && offset.columnOffset === 0)) {
    return content;
  }

  return transformUnquotedFormulaSegments(content, (segment) =>
    segment.replace(CELL_REFERENCE_PATTERN, (match, columnLock: string, columnName: string, rowLock: string, rowName: string, position: number) => {
      const nextCharacter = segment[position + match.length] ?? "";
      if (nextCharacter === "!" || nextCharacter === "(" || nextCharacter === "[") {
        return match;
      }

      const nextColumn = columnLock ? columnNameToIndex(columnName) : columnNameToIndex(columnName) + offset.columnOffset;
      const nextRow = rowLock ? Number(rowName) - 1 : Number(rowName) - 1 + offset.rowOffset;
      if (nextColumn < 0 || nextRow < 0) {
        return "#REF!";
      }

      return `${columnLock}${columnIndexToName(nextColumn)}${rowLock}${nextRow + 1}`;
    })
  );
}

export function translateFormulaRowsWithinColumns(
  content: string,
  context: BoundedFormulaRowTranslation
): string {
  if (!content.startsWith("=") || context.rowOffset === 0) return content;

  const tokens = tokenizeFormulaReferences(content);
  return tokens
    .map((token, index) => {
      if (
        token.kind === "raw"
        || hasExternalWorkbookPrefix(tokens, index, token)
        || !referenceTargetsFormulaSheet(token, context.formulaSheetName)
      ) {
        return token.raw;
      }
      return translateReferenceRowsWithinColumns(token, context);
    })
    .join("");
}

export function extractFormulaReferences(content: string): ExtractedFormulaReference[] {
  if (!content.startsWith("=")) {
    return [];
  }

  const references: ExtractedFormulaReference[] = [];
  const seenLabels = new Set<string>();

  transformUnquotedFormulaSegments(content, (segment) => {
    segment.replace(LOCAL_REFERENCE_PATTERN, (match, startReference: string, endReference: string | undefined, position: number) => {
      const previousCharacter = segment[position - 1] ?? "";
      const nextCharacter = segment[position + match.length] ?? "";
      if (previousCharacter === "!" || nextCharacter === "!" || nextCharacter === "(" || nextCharacter === "[") {
        return match;
      }

      const range = normalizeRange({
        start: parseFormulaCellReference(startReference),
        end: parseFormulaCellReference(endReference ?? startReference)
      });
      const label = formatFormulaReference(range);
      if (!seenLabels.has(label)) {
        references.push({ label, range });
        seenLabels.add(label);
      }
      return match;
    });
    return segment;
  });

  return references;
}

function transformUnquotedFormulaSegments(formula: string, transform: (segment: string) => string): string {
  let result = "";
  let segment = "";
  let isQuoted = false;

  for (let index = 0; index < formula.length; index += 1) {
    const character = formula[index];
    if (character !== '"') {
      segment += character;
      continue;
    }

    if (isQuoted && formula[index + 1] === '"') {
      segment += '""';
      index += 1;
      continue;
    }

    if (isQuoted) {
      result += segment + character;
      segment = "";
      isQuoted = false;
      continue;
    }

    result += transform(segment) + character;
    segment = "";
    isQuoted = true;
  }

  return result + (isQuoted ? segment : transform(segment));
}

function tokenizeFormulaReferences(formula: string): FormulaToken[] {
  const tokens: FormulaToken[] = [];
  let rawStart = 0;
  let index = 0;

  while (index < formula.length) {
    if (formula[index] === '"') {
      index = skipQuotedFormulaSegment(formula, index, '"');
      continue;
    }
    if (formula[index] === "[") {
      index = skipBracketedFormulaSegment(formula, index);
      continue;
    }

    const reference = parseFormulaReferenceToken(formula, index);
    if (!reference) {
      index =
        formula[index] === "'" ? skipQuotedFormulaSegment(formula, index, "'") : index + 1;
      continue;
    }

    if (rawStart < index) {
      tokens.push({ kind: "raw", raw: formula.slice(rawStart, index) });
    }
    const { endIndex, ...token } = reference;
    tokens.push(token);
    index = endIndex;
    rawStart = endIndex;
  }

  if (rawStart < formula.length) {
    tokens.push({ kind: "raw", raw: formula.slice(rawStart) });
  }

  return tokens;
}

function skipBracketedFormulaSegment(formula: string, index: number): number {
  let depth = 0;
  for (let cursor = index; cursor < formula.length; cursor += 1) {
    if (formula[cursor] === "[") {
      depth += 1;
    } else if (formula[cursor] === "]") {
      depth -= 1;
      if (depth === 0) return cursor + 1;
    }
  }
  return formula.length;
}

function parseFormulaReferenceToken(formula: string, index: number): ParsedFormulaReferenceToken | null {
  const previousCharacter = formulaCodePointBefore(formula, index);
  if (
    isFormulaIdentifierContinue(previousCharacter) ||
    previousCharacter === "!" ||
    previousCharacter === ":"
  ) {
    return null;
  }

  let cellStart = index;
  let sheetName: string | undefined;
  let sheetPrefix = "";

  if (formula[index] === "'") {
    const qualifier = parseQuotedSheetQualifier(formula, index);
    if (!qualifier) {
      return null;
    }
    ({ cellStart, sheetName, sheetPrefix } = qualifier);
  } else {
    const qualifier = parseUnquotedSheetQualifier(formula, index);
    if (qualifier) {
      ({ cellStart, sheetName, sheetPrefix } = qualifier);
    }
  }

  const start = parseFormulaCellReferenceAt(formula, cellStart);
  if (!start) {
    return null;
  }

  let cursor = start.endIndex;
  let end: ParsedFormulaCellReference | undefined;
  if (formula[cursor] === ":") {
    const parsedEnd = parseFormulaCellReferenceAt(formula, cursor + 1);
    if (parsedEnd) {
      end = parsedEnd.reference;
      cursor = parsedEnd.endIndex;
    }
  }

  const nextCharacter = formulaCodePointAt(formula, cursor);
  if (
    isFormulaIdentifierContinue(nextCharacter) ||
    nextCharacter === "(" ||
    nextCharacter === "[" ||
    nextCharacter === "!"
  ) {
    return null;
  }

  return {
    kind: "reference",
    raw: formula.slice(index, cursor),
    sheetName,
    sheetPrefix,
    start: start.reference,
    end,
    endIndex: cursor
  };
}

function parseQuotedSheetQualifier(
  formula: string,
  index: number
): { cellStart: number; sheetName: string; sheetPrefix: string } | null {
  let cursor = index + 1;
  let sheetName = "";

  while (cursor < formula.length) {
    if (formula[cursor] !== "'") {
      sheetName += formula[cursor];
      cursor += 1;
      continue;
    }

    if (formula[cursor + 1] === "'") {
      sheetName += "'";
      cursor += 2;
      continue;
    }

    if (formula[cursor + 1] !== "!") {
      return null;
    }

    return {
      cellStart: cursor + 2,
      sheetName,
      sheetPrefix: formula.slice(index, cursor + 2)
    };
  }

  return null;
}

function parseUnquotedSheetQualifier(
  formula: string,
  index: number
): { cellStart: number; sheetName: string; sheetPrefix: string } | null {
  if (!isFormulaIdentifierStart(formulaCodePointAt(formula, index))) {
    return null;
  }

  let cursor = index + formulaCodePointLengthAt(formula, index);
  while (isFormulaIdentifierContinue(formulaCodePointAt(formula, cursor))) {
    cursor += formulaCodePointLengthAt(formula, cursor);
  }
  if (formula[cursor] !== "!") {
    return null;
  }

  return {
    cellStart: cursor + 1,
    sheetName: formula.slice(index, cursor),
    sheetPrefix: formula.slice(index, cursor + 1)
  };
}

function parseFormulaCellReferenceAt(
  formula: string,
  index: number
): { reference: ParsedFormulaCellReference; endIndex: number } | null {
  let cursor = index;
  const columnLock = formula[cursor] === "$" ? "$" : "";
  cursor += columnLock.length;

  const columnStart = cursor;
  while (isAsciiLetter(formula[cursor] ?? "")) {
    cursor += 1;
  }
  if (cursor === columnStart) {
    return null;
  }
  const columnName = formula.slice(columnStart, cursor);

  const rowLock = formula[cursor] === "$" ? "$" : "";
  cursor += rowLock.length;
  const rowStart = cursor;
  while (isAsciiDigit(formula[cursor] ?? "")) {
    cursor += 1;
  }
  const rowName = formula.slice(rowStart, cursor);
  if (rowName.length === 0 || rowName[0] === "0") {
    return null;
  }

  const column = columnNameToIndex(columnName);
  if (column > MAX_FORMULA_COLUMN_INDEX) {
    return null;
  }

  return {
    reference: {
      columnLock,
      rowLock,
      row: Number(rowName) - 1,
      column
    },
    endIndex: cursor
  };
}

function skipQuotedFormulaSegment(formula: string, index: number, quote: '"' | "'"): number {
  let cursor = index + 1;
  while (cursor < formula.length) {
    if (formula[cursor] !== quote) {
      cursor += 1;
      continue;
    }
    if (formula[cursor + 1] === quote) {
      cursor += 2;
      continue;
    }
    return cursor + 1;
  }
  return formula.length;
}

function isFormulaIdentifierStart(character: string): boolean {
  return FORMULA_IDENTIFIER_START_PATTERN.test(character);
}

function isFormulaIdentifierContinue(character: string): boolean {
  return FORMULA_IDENTIFIER_CONTINUE_PATTERN.test(character);
}

function formulaCodePointAt(formula: string, index: number): string {
  const codePoint = formula.codePointAt(index);
  return codePoint === undefined ? "" : String.fromCodePoint(codePoint);
}

function formulaCodePointBefore(formula: string, index: number): string {
  if (index <= 0) {
    return "";
  }

  const lastCodeUnit = formula.charCodeAt(index - 1);
  if (lastCodeUnit >= 0xdc00 && lastCodeUnit <= 0xdfff && index >= 2) {
    const precedingCodeUnit = formula.charCodeAt(index - 2);
    if (precedingCodeUnit >= 0xd800 && precedingCodeUnit <= 0xdbff) {
      return formula.slice(index - 2, index);
    }
  }

  return formula[index - 1];
}

function formulaCodePointLengthAt(formula: string, index: number): number {
  const codePoint = formula.codePointAt(index);
  return codePoint !== undefined && codePoint > 0xffff ? 2 : 1;
}

function isAsciiLetter(character: string): boolean {
  const code = character.charCodeAt(0);
  return (code >= 65 && code <= 90) || (code >= 97 && code <= 122);
}

function isAsciiDigit(character: string): boolean {
  const code = character.charCodeAt(0);
  return code >= 48 && code <= 57;
}

function referenceTargetsEditedSheet(token: FormulaReferenceToken, context: FormulaStructureContext): boolean {
  const targetSheetName = token.sheetName ?? context.formulaSheetName;
  return targetSheetName.toLowerCase() === context.editedSheetName.toLowerCase();
}

function referenceTargetsFormulaSheet(token: FormulaReferenceToken, formulaSheetName: string): boolean {
  const targetSheetName = token.sheetName ?? formulaSheetName;
  return targetSheetName.normalize("NFKC").toLowerCase()
    === formulaSheetName.normalize("NFKC").toLowerCase();
}

function hasExternalWorkbookPrefix(
  tokens: readonly FormulaToken[],
  index: number,
  token: FormulaReferenceToken
): boolean {
  if (!token.sheetPrefix) return false;
  const previous = tokens[index - 1];
  return previous?.kind === "raw" && /\[[^\]\r\n]+\]$/u.test(previous.raw);
}

function shiftReferenceToken(token: FormulaReferenceToken, context: FormulaStructureContext): string {
  if (!token.end) {
    const nextReference = shiftParsedCellReference(token.start, context);
    return nextReference ? `${token.sheetPrefix}${formatParsedCellReference(nextReference)}` : "#REF!";
  }

  const nextRange = shiftParsedReferenceRange(token.start, token.end, context);
  return nextRange
    ? `${token.sheetPrefix}${formatParsedCellReference(nextRange.start)}:${formatParsedCellReference(nextRange.end)}`
    : "#REF!";
}

function shiftParsedCellReference(
  reference: ParsedFormulaCellReference,
  context: FormulaStructureContext
): ParsedFormulaCellReference | null {
  const value = context.axis === "row" ? reference.row : reference.column;
  const nextValue = shiftReferenceIndex(value, context);
  if (nextValue === null) {
    return null;
  }
  return context.axis === "row" ? { ...reference, row: nextValue } : { ...reference, column: nextValue };
}

function shiftParsedReferenceRange(
  start: ParsedFormulaCellReference,
  end: ParsedFormulaCellReference,
  context: FormulaStructureContext
): { start: ParsedFormulaCellReference; end: ParsedFormulaCellReference } | null {
  if (context.mode === "insert") {
    const nextStart = shiftParsedCellReference(start, context);
    const nextEnd = shiftParsedCellReference(end, context);
    return nextStart && nextEnd ? { start: nextStart, end: nextEnd } : null;
  }

  const startValue = context.axis === "row" ? start.row : start.column;
  const endValue = context.axis === "row" ? end.row : end.column;
  const low = Math.min(startValue, endValue);
  const high = Math.max(startValue, endValue);
  const deleteEnd = context.index + context.count;
  const firstSurvivor = low >= context.index && low < deleteEnd ? deleteEnd : low;
  const lastSurvivor = high >= context.index && high < deleteEnd ? context.index - 1 : high;
  if (firstSurvivor > high || lastSurvivor < low || firstSurvivor > lastSurvivor) {
    return null;
  }

  const nextLow = firstSurvivor >= deleteEnd ? firstSurvivor - context.count : firstSurvivor;
  const nextHigh = lastSurvivor >= deleteEnd ? lastSurvivor - context.count : lastSurvivor;
  const nextStartValue = startValue <= endValue ? nextLow : nextHigh;
  const nextEndValue = startValue <= endValue ? nextHigh : nextLow;

  return context.axis === "row"
    ? { start: { ...start, row: nextStartValue }, end: { ...end, row: nextEndValue } }
    : { start: { ...start, column: nextStartValue }, end: { ...end, column: nextEndValue } };
}

function shiftReferenceIndex(value: number, context: FormulaStructureContext): number | null {
  const end = context.index + context.count;
  if (context.mode === "insert") {
    return value >= context.index ? value + context.count : value;
  }
  if (value >= context.index && value < end) {
    return null;
  }
  return value >= end ? value - context.count : value;
}

function rectangularReferenceTargetsEditedSheet(
  token: FormulaReferenceToken,
  context: RectangularRowEditContext
): boolean {
  const target = (token.sheetName ?? context.formulaSheetId).normalize("NFKC").toLowerCase();
  return target === context.editedSheetId.normalize("NFKC").toLowerCase();
}

function translateReferenceRowsWithinColumns(
  token: FormulaReferenceToken,
  context: BoundedFormulaRowTranslation
): string {
  const originalEnd = token.end ?? token.start;
  const lowColumn = Math.min(token.start.column, originalEnd.column);
  const highColumn = Math.max(token.start.column, originalEnd.column);
  if (highColumn < context.columnStart || lowColumn > context.columnEnd) return token.raw;
  if (token.end && (lowColumn < context.columnStart || highColumn > context.columnEnd)) {
    return token.raw;
  }

  if (!token.end) {
    const translated = translateRelativeReferenceRow(token.start, context.rowOffset);
    return translated
      ? `${token.sheetPrefix}${formatParsedCellReference(translated)}`
      : "#REF!";
  }

  return formatTranslatedReferenceMember(
    token.sheetPrefix,
    translateRelativeReferenceRow(token.start, context.rowOffset),
    translateRelativeReferenceRow(token.end, context.rowOffset)
  );
}

function translateRelativeReferenceRow(
  reference: ParsedFormulaCellReference,
  rowOffset: number
): ParsedFormulaCellReference | null {
  if (reference.rowLock) return reference;
  const row = reference.row + rowOffset;
  return row < 0 ? null : { ...reference, row };
}

function formatTranslatedReferenceMember(
  sheetPrefix: string,
  start: ParsedFormulaCellReference | null,
  end: ParsedFormulaCellReference | null
): string {
  if (!start || !end) return "#REF!";
  const startText = formatParsedCellReference(start);
  const endText = formatParsedCellReference(end);
  if (start.row === end.row && start.column === end.column) {
    return `${sheetPrefix}${startText}`;
  }
  return `${sheetPrefix}${startText}:${endText}`;
}

function rewriteRectangularReferenceToken(
  token: FormulaReferenceToken,
  context: RectangularRowEditContext
): FormulaRewriteResult {
  const originalEnd = token.end ?? token.start;
  const lowColumn = Math.min(token.start.column, originalEnd.column);
  const highColumn = Math.max(token.start.column, originalEnd.column);
  if (highColumn < context.tableColumnStart || lowColumn > context.tableColumnEnd) {
    return { ok: true, formula: token.raw };
  }

  const affectedRowInterval = rewriteRowIntervalForRectangularEdit(token.start.row, originalEnd.row, context);
  if (
    affectedRowInterval
    && affectedRowInterval.start === token.start.row
    && affectedRowInterval.end === originalEnd.row
  ) {
    return { ok: true, formula: token.raw };
  }

  const slices: Array<{ startColumn: number; endColumn: number; affected: boolean }> = [];
  if (lowColumn < context.tableColumnStart) {
    slices.push({
      startColumn: lowColumn,
      endColumn: Math.min(highColumn, context.tableColumnStart - 1),
      affected: false
    });
  }
  const insideStart = Math.max(lowColumn, context.tableColumnStart);
  const insideEnd = Math.min(highColumn, context.tableColumnEnd);
  if (insideStart <= insideEnd) {
    slices.push({ startColumn: insideStart, endColumn: insideEnd, affected: true });
  }
  if (highColumn > context.tableColumnEnd) {
    slices.push({
      startColumn: Math.max(lowColumn, context.tableColumnEnd + 1),
      endColumn: highColumn,
      affected: false
    });
  }

  const members = slices.flatMap((slice) => {
    const rowInterval = slice.affected
      ? affectedRowInterval
      : { start: token.start.row, end: originalEnd.row };
    if (!rowInterval) return [];
    const start = withRectangularCoordinates(token.start, slice.startColumn, rowInterval.start);
    const end = withRectangularCoordinates(originalEnd, slice.endColumn, rowInterval.end);
    return [formatRectangularMember(token.sheetPrefix, start, end)];
  });
  if (members.length === 0) return { ok: true, formula: "#REF!" };
  if (members.length === 1) return { ok: true, formula: members[0] };
  return {
    ok: false,
    issue: {
      code: "TABLE_FORMULA_REFERENCE_UNSUPPORTED",
      message: "Formula ranges spanning table and non-table columns cannot be rewritten for this table row edit"
    }
  };
}

export function rewriteRowIntervalForRectangularEdit(
  startRow: number,
  endRow: number,
  context: Pick<RectangularRowEditContext, "row" | "count" | "operation" | "tableRowEnd">
): { start: number; end: number } | null {
  const ascending = startRow <= endRow;
  const low = Math.min(startRow, endRow);
  const high = Math.max(startRow, endRow);
  if (high < context.row || low > context.tableRowEnd) {
    return { start: startRow, end: endRow };
  }

  let nextLow = low;
  let nextHigh = high;
  if (context.operation === "insert") {
    if (low >= context.row && low <= context.tableRowEnd) {
      nextLow += context.count;
    }
    if (high >= context.row && high <= context.tableRowEnd) {
      nextHigh += context.count;
    }
  } else {
    const deleteEnd = Math.min(
      context.row + context.count - 1,
      context.tableRowEnd
    );
    const lowDeleted = low >= context.row && low <= deleteEnd;
    const highDeleted = high >= context.row && high <= deleteEnd;
    if (lowDeleted && highDeleted) return null;

    if (lowDeleted) nextLow = context.row;
    else if (low > deleteEnd && low <= context.tableRowEnd) {
      nextLow -= context.count;
    }

    if (highDeleted) nextHigh = context.row - 1;
    else if (high > deleteEnd && high <= context.tableRowEnd) {
      nextHigh -= context.count;
    }
  }

  if (nextLow > nextHigh) [nextLow, nextHigh] = [nextHigh, nextLow];
  return ascending ? { start: nextLow, end: nextHigh } : { start: nextHigh, end: nextLow };
}

function withRectangularCoordinates(
  reference: ParsedFormulaCellReference,
  column: number,
  row: number
): ParsedFormulaCellReference {
  return { ...reference, column, row };
}

function formatRectangularMember(
  sheetPrefix: string,
  start: ParsedFormulaCellReference,
  end: ParsedFormulaCellReference
): string {
  const startText = formatParsedCellReference(start);
  const endText = formatParsedCellReference(end);
  return start.row === end.row && start.column === end.column
    ? `${sheetPrefix}${startText}`
    : `${sheetPrefix}${startText}:${endText}`;
}

function formatParsedCellReference(reference: ParsedFormulaCellReference): string {
  return `${reference.columnLock}${columnIndexToName(reference.column)}${reference.rowLock}${reference.row + 1}`;
}

function parseFormulaCellReference(reference: string) {
  const normalized = reference.replace(/\$/g, "").toUpperCase();
  const match = normalized.match(/^([A-Z]+)([1-9]\d*)$/);
  if (!match) {
    throw new Error(`Invalid formula reference: ${reference}`);
  }

  return {
    row: Number(match[2]) - 1,
    column: columnNameToIndex(match[1])
  };
}

function formatFormulaReference(range: CellRange): string {
  const normalized = normalizeRange(range);
  const start = formatCellAddress(normalized.start);
  const end = formatCellAddress(normalized.end);
  return start === end ? start : `${start}:${end}`;
}
