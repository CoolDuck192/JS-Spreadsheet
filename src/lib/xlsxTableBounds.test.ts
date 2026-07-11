import { describe, expect, it, vi } from "vitest";
import { createBlankWorkbook } from "./workbook";

vi.mock("./xlsxTables", async (importOriginal) => {
  const actual = await importOriginal<typeof import("./xlsxTables")>();
  return {
    ...actual,
    importStructuredTablesFromWorksheet: vi.fn(async (_worksheet, sheetId: string) => [{
      id: "table-trailing-empty-rows",
      name: "TrailingEmptyRows",
      sheetId,
      range: {
        start: { row: 0, column: 0 },
        end: { row: 299, column: 1 }
      },
      headerRow: true,
      totalsRow: false,
      columns: [
        { id: "column-a", name: "A", sheetColumn: 0 },
        { id: "column-b", name: "B", sheetColumn: 1 }
      ],
      rowIds: Array.from({ length: 299 }, (_, index) => `row-${index + 1}`)
    }])
  };
});

import { exportWorkbookToXlsx, importWorkbookFromXlsx } from "./xlsx";

describe("XLSX table worksheet bounds", () => {
  it("includes separately imported table ranges in the sheet dimensions", async () => {
    const bytes = await exportWorkbookToXlsx(createBlankWorkbook());

    const imported = await importWorkbookFromXlsx(bytes);

    expect(imported.sheets[0].rowCount).toBe(300);
    await expect(exportWorkbookToXlsx(imported)).resolves.toBeInstanceOf(ArrayBuffer);
  });
});
