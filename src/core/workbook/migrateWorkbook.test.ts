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

  it.each([
    ["comments", { comments: { A1: 42 } }],
    ["hyperlinks", { hyperlinks: { A1: { href: "https://example.test" } } }],
    ["validations", { validations: { A1: { type: "list" } } }],
    ["conditional formats", { conditionalFormats: [{ id: "rule-1", range: cellRange(), format: {} }] }],
    ["filters", { filters: [{ id: "filter-1", range: cellRange(), operator: "equals", value: "Open" }] }],
    ["charts", { charts: [{ id: "chart-1", title: "Chart", type: "bar", range: cellRange() }] }],
    ["merges", { merges: [{ id: "merge-1" }] }]
  ])("rejects malformed %s instead of casting persisted values", (_label, sheetPatch) => {
    const fixture = createBlankWorkbook();
    expect(migrateWorkbookModel({
      ...fixture,
      sheets: [{ ...fixture.sheets[0], ...sheetPatch }]
    })).toBeNull();
  });

  it("canonicalizes accepted aliases for cell-address keyed metadata", () => {
    const fixture = createBlankWorkbook();
    const migrated = migrateWorkbookModel({
      ...fixture,
      sheets: [{
        ...fixture.sheets[0],
        comments: { " a1 ": "note" },
        hyperlinks: { b2: "https://example.test" },
        validations: { " c3 ": { type: "list", values: ["Open", "Closed"] } }
      }]
    });

    expect(migrated?.sheets[0]).toMatchObject({
      comments: { A1: "note" },
      hyperlinks: { B2: "https://example.test" },
      validations: { C3: { type: "list", values: ["Open", "Closed"] } }
    });
    expect(migrated?.sheets[0].comments).not.toHaveProperty(" a1 ");
    expect(migrated?.sheets[0].hyperlinks).not.toHaveProperty("b2");
    expect(migrated?.sheets[0].validations).not.toHaveProperty(" c3 ");
  });

  it.each([
    ["comments", { comments: { A1: "first", " a1 ": "second" } }],
    ["hyperlinks", { hyperlinks: { A1: "https://first.test", a1: "https://second.test" } }],
    ["validations", {
      validations: {
        A1: { type: "list", values: ["First"] },
        " a1 ": { type: "list", values: ["Second"] }
      }
    }]
  ])("rejects normalized duplicate %s addresses", (_label, sheetPatch) => {
    const fixture = createBlankWorkbook();
    expect(migrateWorkbookModel({
      ...fixture,
      sheets: [{ ...fixture.sheets[0], ...sheetPatch }]
    })).toBeNull();
  });

  it("accepts finite text-length bounds supported by the public workbook model", () => {
    const fixture = createBlankWorkbook();
    const migrated = migrateWorkbookModel({
      ...fixture,
      sheets: [{
        ...fixture.sheets[0],
        validations: { A1: { type: "textLength", min: -1.5, max: 3.5 } }
      }]
    });

    expect(migrated?.sheets[0].validations.A1).toEqual({
      type: "textLength",
      min: -1.5,
      max: 3.5
    });
  });

  it("validates and deeply clones persisted sheet collections", () => {
    const fixture = createBlankWorkbook();
    const range = cellRange();
    const sheet = {
      ...fixture.sheets[0],
      comments: { A1: "note" },
      hyperlinks: { A1: "https://example.test" },
      validations: { A1: { type: "list" as const, values: ["Open", "Closed"], allowBlank: false } },
      conditionalFormats: [{
        id: "rule-1",
        range,
        condition: { type: "between" as const, value: "1", secondValue: "5" },
        format: { bold: true, borders: { top: { style: "thin" as const, color: "#123456" } } }
      }],
      filters: [{
        id: "filter-1",
        range,
        column: 0,
        operator: "equals" as const,
        value: "Open",
        values: ["Open"],
        hasHeader: true
      }],
      charts: [{ id: "chart-1", title: "Chart", type: "bar" as const, range, anchor: { row: 2, column: 2 } }],
      merges: [{ id: "merge-1", range }]
    };

    const migratedSheet = migrateWorkbookModel({ ...fixture, sheets: [sheet] })?.sheets[0];

    expect(migratedSheet).toBeDefined();
    expect(migratedSheet?.comments).toEqual(sheet.comments);
    expect(migratedSheet?.comments).not.toBe(sheet.comments);
    expect(migratedSheet?.hyperlinks).not.toBe(sheet.hyperlinks);
    expect(migratedSheet?.validations.A1).not.toBe(sheet.validations.A1);
    expect((migratedSheet?.validations.A1 as { values: readonly string[] }).values)
      .not.toBe(sheet.validations.A1.values);
    expect(migratedSheet?.conditionalFormats[0]).not.toBe(sheet.conditionalFormats[0]);
    expect(migratedSheet?.conditionalFormats[0].range).not.toBe(range);
    expect(migratedSheet?.conditionalFormats[0].condition).not.toBe(sheet.conditionalFormats[0].condition);
    expect(migratedSheet?.conditionalFormats[0].format.borders).not.toBe(sheet.conditionalFormats[0].format.borders);
    expect(migratedSheet?.filters[0]).not.toBe(sheet.filters[0]);
    expect(migratedSheet?.filters[0].values).not.toBe(sheet.filters[0].values);
    expect(migratedSheet?.charts[0].anchor).not.toBe(sheet.charts[0].anchor);
    expect(migratedSheet?.merges[0].range).not.toBe(range);
  });
});

function cellRange() {
  return { start: { row: 0, column: 0 }, end: { row: 1, column: 1 } };
}

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
