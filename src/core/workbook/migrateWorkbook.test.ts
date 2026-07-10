import { describe, expect, it } from "vitest";
import { createBlankWorkbook } from "../../lib/workbook";
import { migrateWorkbookModel } from "./migrateWorkbook";

describe("workbook model migration", () => {
  it("migrates a valid version 1 workbook to version 2", () => {
    const current = createBlankWorkbook();
    const versionOneFixture = { ...current, version: 1 };
    const migrated = migrateWorkbookModel(versionOneFixture);
    expect(migrated).toMatchObject({ version: 2, tables: [] });
  });

  it("preserves app-native structured IDs in version 2", () => {
    const versionTwoFixture = createVersionTwoFixture();
    const migrated = migrateWorkbookModel(versionTwoFixture);
    expect(migrated?.tables[0]).toMatchObject({
      id: "table-fixed",
      columns: [{ id: "column-fixed" }],
      rowIds: ["row-fixed"]
    });
  });

  it.each([
    null,
    {},
    { version: 99 },
    { ...createVersionTwoFixture(), tables: [{ ...createVersionTwoFixture().tables[0], rowIds: [] }] }
  ])("rejects invalid persisted input", (value) => {
    expect(migrateWorkbookModel(value)).toBeNull();
  });

  it("rejects normalized duplicate names, overlapping ranges, merges, and broken references", () => {
    const fixture = createVersionTwoFixture();
    const duplicate = {
      ...fixture.tables[0],
      id: "table-duplicate",
      name: "ＥＭＰＬＯＹＥＥＳ"
    };
    expect(migrateWorkbookModel({ ...fixture, tables: [...fixture.tables, duplicate] })).toBeNull();
    expect(migrateWorkbookModel({
      ...fixture,
      sheets: [{ ...fixture.sheets[0], merges: [{ id: "merge-1", range: fixture.tables[0].range }] }]
    })).toBeNull();
    expect(migrateWorkbookModel({
      ...fixture,
      tables: [{ ...fixture.tables[0], keyColumnId: "missing-column" }]
    })).toBeNull();
  });
});

function createVersionTwoFixture() {
  const current = createBlankWorkbook();
  const sheet = {
    ...current.sheets[0],
    cells: { A1: "Employee ID", A2: 101 }
  };
  return {
    ...current,
    version: 2,
    sheets: [sheet],
    tables: [{
      id: "table-fixed",
      name: "Employees",
      sheetId: sheet.id,
      range: { start: { row: 0, column: 0 }, end: { row: 1, column: 0 } },
      headerRow: true,
      totalsRow: false,
      columns: [{ id: "column-fixed", name: "Employee ID", sheetColumn: 0 }],
      rowIds: ["row-fixed"],
      keyColumnId: "column-fixed"
    }]
  };
}
