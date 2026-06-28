import type { CellCoord, CellRange } from "../types";

const CELL_ADDRESS_PATTERN = /^([A-Z]+)([1-9]\d*)$/i;

export function columnIndexToName(index: number): string {
  if (!Number.isInteger(index) || index < 0) {
    throw new Error("Column index must be a non-negative integer");
  }

  let remaining = index + 1;
  let name = "";

  while (remaining > 0) {
    const remainder = (remaining - 1) % 26;
    name = String.fromCharCode(65 + remainder) + name;
    remaining = Math.floor((remaining - 1) / 26);
  }

  return name;
}

export function columnNameToIndex(name: string): number {
  const normalized = name.trim().toUpperCase();
  if (!/^[A-Z]+$/.test(normalized)) {
    throw new Error(`Invalid column name: ${name}`);
  }

  let index = 0;
  for (const character of normalized) {
    index = index * 26 + (character.charCodeAt(0) - 64);
  }

  return index - 1;
}

export function parseCellAddress(address: string): CellCoord {
  const match = address.trim().match(CELL_ADDRESS_PATTERN);
  if (!match) {
    throw new Error(`Invalid cell address: ${address}`);
  }

  return {
    row: Number(match[2]) - 1,
    column: columnNameToIndex(match[1])
  };
}

export function formatCellAddress(coord: CellCoord): string {
  if (!Number.isInteger(coord.row) || coord.row < 0 || !Number.isInteger(coord.column) || coord.column < 0) {
    throw new Error("Cell coordinates must be non-negative integers");
  }

  return `${columnIndexToName(coord.column)}${coord.row + 1}`;
}

export function normalizeRange(range: CellRange): CellRange {
  return {
    start: {
      row: Math.min(range.start.row, range.end.row),
      column: Math.min(range.start.column, range.end.column)
    },
    end: {
      row: Math.max(range.start.row, range.end.row),
      column: Math.max(range.start.column, range.end.column)
    }
  };
}

export function parseRangeAddress(address: string): CellRange {
  const [start, end = start] = address.split(":");
  return normalizeRange({
    start: parseCellAddress(start),
    end: parseCellAddress(end)
  });
}

export function getRangeAddresses(range: CellRange): string[] {
  const normalized = normalizeRange(range);
  const addresses: string[] = [];

  for (let row = normalized.start.row; row <= normalized.end.row; row += 1) {
    for (let column = normalized.start.column; column <= normalized.end.column; column += 1) {
      addresses.push(formatCellAddress({ row, column }));
    }
  }

  return addresses;
}
