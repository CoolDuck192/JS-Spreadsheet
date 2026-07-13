import { readFile } from "node:fs/promises";
import { resolve } from "node:path";
import { describe, expect, it } from "vitest";
import { importWorkbookFromXlsx } from "./xlsx";

const fixturePath = resolve("src/test/fixtures/xlsx/real-kitchen-sink.xlsx");
const FOREIGN_IMPORT_BUDGET_MS = 30_000;

describe("full foreign XLSX import", () => {
  it("imports the generated 80,000-cell kitchen sink", async () => {
    const source = await readFile(fixturePath);
    expect(source.byteLength).toBeLessThanOrEqual(500 * 1024);

    const startedAt = performance.now();
    const workbook = await importWorkbookFromXlsx(source);
    const importDurationMs = performance.now() - startedAt;
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
    // eslint-disable-next-line no-console
    console.log(`foreign-xlsx-import(80k cells)=${Math.round(importDurationMs)}ms`);
    if (!process.env.SKIP_PERF_ASSERT) {
      expect(importDurationMs).toBeLessThan(FOREIGN_IMPORT_BUDGET_MS);
    }
  }, 60_000);
});
