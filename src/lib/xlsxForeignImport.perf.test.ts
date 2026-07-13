import { readFile } from "node:fs/promises";
import { describe, expect, it } from "vitest";
import { importWorkbookFromXlsx } from "./xlsx";

const fixturePath = process.env.XLSX_FULL_IMPORT_FIXTURE ?? "";

describe.runIf(fixturePath.length > 0)("full foreign XLSX import", () => {
  it("imports the generated 80,000-cell kitchen sink", async () => {
    const source = await readFile(fixturePath);
    const workbook = await importWorkbookFromXlsx(
      source.buffer.slice(source.byteOffset, source.byteOffset + source.byteLength)
    );
    const sales = workbook.sheets.find((sheet) => sheet.name === "Sales")!;
    const dashboard = workbook.sheets.find((sheet) => sheet.name === "Dashboard")!;
    const dense = workbook.sheets.find((sheet) => sheet.name === "Data10k")!;
    const table = workbook.tables.find((candidate) => candidate.name === "SalesTable")!;

    expect(workbook.sheets.map((sheet) => sheet.name)).toEqual(["Sales", "Dashboard", "Data10k"]);
    expect(table.range).toEqual({ start: { row: 0, column: 0 }, end: { row: 50, column: 5 } });
    expect(table.rowIds).toHaveLength(50);
    expect(Array.from({ length: 50 }, (_, index) => sales.cells[`F${index + 2}`])).toEqual(
      Array.from({ length: 50 }, (_, index) => `=D${index + 2}*E${index + 2}`)
    );
    expect(dashboard.comments.A10).toBe("Checked by finance");
    expect(dashboard.cells.A9).toBe("Docs");
    expect(Object.keys(dense.cells)).toHaveLength(80_000);
    expect([dense.cells.A1, dense.cells.H10000]).toEqual([1, 1]);
    expect(dense.rowCount).toBe(10_000);
  }, 60_000);
});
