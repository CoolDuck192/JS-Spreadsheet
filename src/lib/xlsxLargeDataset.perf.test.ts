import { describe, expect, it } from "vitest";
import type { SheetModel, WorkbookModel } from "../types";
import { createBlankWorkbook } from "./workbook";
import { exportWorkbookToXlsx, importWorkbookFromXlsx } from "./xlsx";

const DATA_ROWS = 100_001;

function largeWorkbook(): WorkbookModel {
  const workbook = createBlankWorkbook();
  const sheet = workbook.sheets[0];
  const cells: SheetModel["cells"] = {
    A1: "First",
    B1: "Second",
    C1: "Total"
  };
  for (let row = 1; row <= DATA_ROWS; row += 1) {
    const worksheetRow = row + 1;
    cells[`A${worksheetRow}`] = row;
    cells[`B${worksheetRow}`] = row * 2;
    cells[`C${worksheetRow}`] = row * 3;
  }
  return {
    ...workbook,
    sheets: [{
      ...sheet,
      rowCount: DATA_ROWS + 1,
      cells
    }]
  };
}

describe("large XLSX round-trip", () => {
  it("exports and imports more than 100k rows under the default security limits", async () => {
    const exported = await exportWorkbookToXlsx(largeWorkbook());

    const imported = await importWorkbookFromXlsx(exported);

    expect(imported.sheets[0].rowCount).toBe(DATA_ROWS + 1);
    expect(imported.sheets[0].cells[`C${DATA_ROWS + 1}`]).toBe(DATA_ROWS * 3);
  }, 60_000);
});
