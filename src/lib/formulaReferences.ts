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

const CELL_REFERENCE_PATTERN = /(?<![A-Z0-9_.])(\$?)([A-Z]+)(\$?)([1-9]\d*)/gi;
const LOCAL_REFERENCE_PATTERN = /(?<![A-Z0-9_.])(\$?[A-Z]+\$?[1-9]\d*)(?::(\$?[A-Z]+\$?[1-9]\d*))?/gi;

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
