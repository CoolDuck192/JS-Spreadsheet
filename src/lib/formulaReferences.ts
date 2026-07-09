import type { CellRange } from "../types";
import { columnIndexToName, columnNameToIndex, formatCellAddress, normalizeRange } from "./addressing";

export type FormulaReferenceOffset = {
  rowOffset: number;
  columnOffset: number;
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

function parseFormulaReferenceToken(formula: string, index: number): ParsedFormulaReferenceToken | null {
  const previousCharacter = formula[index - 1] ?? "";
  if (isFormulaIdentifierCharacter(previousCharacter) || previousCharacter === "!" || previousCharacter === ":") {
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

  const nextCharacter = formula[cursor] ?? "";
  if (
    isFormulaIdentifierCharacter(nextCharacter) ||
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
  if (!isFormulaIdentifierCharacter(formula[index] ?? "")) {
    return null;
  }

  let cursor = index;
  while (isFormulaIdentifierCharacter(formula[cursor] ?? "")) {
    cursor += 1;
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

function isFormulaIdentifierCharacter(character: string): boolean {
  return isAsciiLetter(character) || isAsciiDigit(character) || character === "_" || character === ".";
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

function shiftReferenceToken(token: FormulaReferenceToken, context: FormulaStructureContext): string {
  if (!token.end) {
    const nextReference = shiftParsedCellReference(token.start, context);
    return nextReference ? `${token.sheetPrefix}${formatParsedCellReference(nextReference)}` : `${token.sheetPrefix}#REF!`;
  }

  const nextRange = shiftParsedReferenceRange(token.start, token.end, context);
  return nextRange
    ? `${token.sheetPrefix}${formatParsedCellReference(nextRange.start)}:${formatParsedCellReference(nextRange.end)}`
    : `${token.sheetPrefix}#REF!`;
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
