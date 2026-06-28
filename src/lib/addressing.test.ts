import { describe, expect, it } from "vitest";
import {
  columnIndexToName,
  formatCellAddress,
  getRangeAddresses,
  normalizeRange,
  parseCellAddress,
  parseRangeAddress
} from "./addressing";

describe("addressing", () => {
  it("converts zero-based column indexes into spreadsheet column names", () => {
    expect(columnIndexToName(0)).toBe("A");
    expect(columnIndexToName(25)).toBe("Z");
    expect(columnIndexToName(26)).toBe("AA");
    expect(columnIndexToName(54)).toBe("BC");
  });

  it("parses and formats A1-style addresses", () => {
    expect(parseCellAddress("BC23")).toEqual({ row: 22, column: 54 });
    expect(formatCellAddress({ row: 4, column: 2 })).toBe("C5");
  });

  it("parses ranges and normalizes reverse selections", () => {
    expect(parseRangeAddress("A1:C3")).toEqual({
      start: { row: 0, column: 0 },
      end: { row: 2, column: 2 }
    });

    expect(normalizeRange({ start: { row: 5, column: 4 }, end: { row: 2, column: 1 } })).toEqual({
      start: { row: 2, column: 1 },
      end: { row: 5, column: 4 }
    });
  });

  it("iterates addresses in row-major order", () => {
    expect(
      getRangeAddresses({ start: { row: 0, column: 0 }, end: { row: 1, column: 1 } })
    ).toEqual(["A1", "B1", "A2", "B2"]);
  });

  it("rejects invalid addresses", () => {
    expect(() => parseCellAddress("1A")).toThrow("Invalid cell address");
    expect(() => parseRangeAddress("A1:B")).toThrow("Invalid cell address");
    expect(() => columnIndexToName(-1)).toThrow("Column index");
  });
});
