import { describe, expect, it } from "vitest";
import { parseCsv, serializeCsv } from "./csv";

describe("csv", () => {
  it("parses simple comma-separated rows", () => {
    expect(parseCsv("A,B,C\n1,2,3")).toEqual([
      ["A", "B", "C"],
      ["1", "2", "3"]
    ]);
  });

  it("parses quoted commas, escaped quotes, blanks, and multiline fields", () => {
    const csv = '"Name","Note","Blank"\n"Riya","hello, ""spreadsheet""",""\n"Line","first\nsecond",42';

    expect(parseCsv(csv)).toEqual([
      ["Name", "Note", "Blank"],
      ["Riya", 'hello, "spreadsheet"', ""],
      ["Line", "first\nsecond", "42"]
    ]);
  });

  it("serializes rows with quotes only where needed", () => {
    expect(
      serializeCsv([
        ["A", "hello, world", 'say "yes"'],
        ["Line", "first\nsecond", ""]
      ])
    ).toBe('A,"hello, world","say ""yes"""\nLine,"first\nsecond",');
  });

  it("round trips parsed rows", () => {
    const rows = [
      ["Name", "Amount", "Formula"],
      ["Asha", "10", "=SUM(A1:A3)"],
      ["Quoted", 'A "quote"', "comma, value"]
    ];

    expect(parseCsv(serializeCsv(rows))).toEqual(rows);
  });
});
