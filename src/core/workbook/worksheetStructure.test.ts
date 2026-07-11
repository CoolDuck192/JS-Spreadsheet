import { beforeEach, describe, expect, it, vi } from "vitest";
import type { CellRange, StructuredTable, WorkbookModel } from "../../types";
import {
  addSheet,
  createBlankWorkbook,
  getCellContent,
  setCellContent,
  setSheetProtection
} from "../../lib/workbook";
import { isWorksheetStructureCommand, reduceWorksheetStructureCommand } from "./worksheetStructure";

function deterministicServices(returnedIds?: readonly string[]) {
  let sequence = 0;
  return {
    createId: vi.fn((kind: string) => {
      const index = sequence;
      sequence += 1;
      return returnedIds?.[index] ?? `${kind}-${index + 1}`;
    })
  };
}

const services = deterministicServices();

function range(startRow: number, startColumn: number, endRow: number, endColumn: number): CellRange {
  return {
    start: { row: startRow, column: startColumn },
    end: { row: endRow, column: endColumn }
  };
}

function structuredWorkbook(options: {
  costStart?: number;
  headerRow?: boolean;
  salesNames?: readonly [string, string];
} = {}): WorkbookModel {
  const costStart = options.costStart ?? 3;
  const headerRow = options.headerRow ?? true;
  const salesNames = options.salesNames ?? ["Region", "Sales"];
  let workbook = createBlankWorkbook();
  const sheetId = workbook.activeSheetId;
  const values: readonly (readonly [string, string | number])[] = [
    ["A1", salesNames[0]], ["B1", salesNames[1]],
    ["A2", "West"], ["B2", 10],
    ["A3", "East"], ["B3", 20],
    [`${costStart === 2 ? "C" : "D"}1`, "Category"],
    [`${costStart === 2 ? "D" : "E"}1`, "Costs"]
  ];
  for (const [address, value] of values) {
    workbook = setCellContent(workbook, sheetId, address, value);
  }
  const sales: StructuredTable = {
    id: "table-sales",
    name: "SalesTable",
    sheetId,
    range: range(0, 0, 2, 1),
    headerRow,
    totalsRow: false,
    columns: [
      {
        id: "sales-region",
        name: salesNames[0],
        sheetColumn: 0,
        dataType: "text",
        totalsLabel: "Total"
      },
      {
        id: "sales-value",
        name: salesNames[1],
        sheetColumn: 1,
        dataType: "number",
        calculatedFormula: "=A2*2",
        totalsFunction: "sum"
      }
    ],
    rowIds: headerRow ? ["sales-row-1", "sales-row-2"] : ["sales-row-1", "sales-row-2", "sales-row-3"],
    keyColumnId: "sales-region",
    style: { theme: "TableStyleMedium2", showRowStripes: true, showFirstColumn: true },
    sort: [{ columnId: "sales-value", direction: "desc", nulls: "last" }],
    filter: {
      kind: "comparison",
      columnId: "sales-region",
      operator: "eq",
      value: { type: "string", value: "West" }
    }
  };
  const costs: StructuredTable = {
    id: "table-costs",
    name: "CostsTable",
    sheetId,
    range: range(0, costStart, 2, costStart + 1),
    headerRow: true,
    totalsRow: false,
    columns: [
      { id: "costs-category", name: "Category", sheetColumn: costStart },
      { id: "costs-value", name: "Costs", sheetColumn: costStart + 1 }
    ],
    rowIds: ["costs-row-1", "costs-row-2"],
    style: { theme: "TableStyleLight1", showColumnStripes: true }
  };
  return { ...workbook, tables: [sales, costs] };
}

function withCrossSheetTable(workbook: WorkbookModel): WorkbookModel {
  const primarySheetId = workbook.activeSheetId;
  const withSheet = addSheet(workbook, "Other");
  const otherSheetId = withSheet.activeSheetId;
  const crossSheetTable: StructuredTable = {
    ...workbook.tables[0],
    id: "table-cross-sheet",
    name: "CrossSheetTable",
    sheetId: otherSheetId,
    range: range(0, 2, 2, 3),
    columns: workbook.tables[0].columns.map((column, index) => ({
      ...column,
      id: `cross-column-${index + 1}`,
      sheetColumn: index + 2
    }))
  };
  return {
    ...withSheet,
    activeSheetId: primarySheetId,
    tables: [...workbook.tables, crossSheetTable]
  };
}

function threeColumnCalculatedWorkbook(startColumn = 0): WorkbookModel {
  const workbook = createBlankWorkbook();
  const table: StructuredTable = {
    id: "table-calculated",
    name: "CalculatedTable",
    sheetId: workbook.activeSheetId,
    range: range(0, startColumn, 3, startColumn + 2),
    headerRow: true,
    totalsRow: false,
    columns: [
      {
        id: "column-a",
        name: "A",
        sheetColumn: startColumn,
        dataType: "text",
        totalsLabel: "Total"
      },
      {
        id: "column-b",
        name: "B",
        sheetColumn: startColumn + 1,
        dataType: "number",
        totalsFunction: "sum"
      },
      {
        id: "column-c",
        name: "C",
        sheetColumn: startColumn + 2,
        calculatedFormula: "=A2*2"
      }
    ],
    rowIds: ["row-1", "row-2", "row-3"],
    keyColumnId: "column-b",
    sort: [{ columnId: "column-b", direction: "asc" }],
    filter: {
      kind: "logical",
      operator: "and",
      operands: [
        {
          kind: "comparison",
          columnId: "column-a",
          operator: "eq",
          value: { type: "string", value: "keep" }
        },
        { kind: "blank", columnId: "column-b", operator: "isBlank" }
      ]
    }
  };
  return { ...workbook, tables: [table] };
}

function adjacentSingleColumnTables(): WorkbookModel {
  const workbook = createBlankWorkbook();
  const table = (id: string, column: number): StructuredTable => ({
    id,
    name: id === "table-left" ? "LeftTable" : "RightTable",
    sheetId: workbook.activeSheetId,
    range: range(0, column, 2, column),
    headerRow: true,
    totalsRow: false,
    columns: [{ id: `${id}-column`, name: "Value", sheetColumn: column }],
    rowIds: [`${id}-row-1`, `${id}-row-2`]
  });
  return { ...workbook, tables: [table("table-left", 0), table("table-right", 1)] };
}

function tableStartingAtRowFive(): WorkbookModel {
  const workbook = createBlankWorkbook();
  const table: StructuredTable = {
    id: "table-row-guard",
    name: "RowGuardTable",
    sheetId: workbook.activeSheetId,
    range: range(4, 0, 6, 1),
    headerRow: true,
    totalsRow: false,
    columns: [
      { id: "column-label", name: "Label", sheetColumn: 0 },
      { id: "column-total", name: "Total", sheetColumn: 1, calculatedFormula: "=A6*2" }
    ],
    rowIds: ["row-1", "row-2"]
  };
  return { ...workbook, tables: [table] };
}

function calculatedFormulaWorkbook(formula: string): WorkbookModel {
  const workbook = createBlankWorkbook();
  const table: StructuredTable = {
    id: "table-formulas",
    name: "FormulaTable",
    sheetId: workbook.activeSheetId,
    range: range(0, 0, 3, 2),
    headerRow: true,
    totalsRow: false,
    columns: [
      { id: "source-a", name: "A", sheetColumn: 0 },
      { id: "source-b", name: "B", sheetColumn: 1 },
      { id: "calculated", name: "Calculated", sheetColumn: 2, calculatedFormula: formula }
    ],
    rowIds: ["formula-row-1", "formula-row-2", "formula-row-3"]
  };
  return { ...workbook, tables: [table] };
}

describe("worksheet structure reducer", () => {
  beforeEach(() => {
    services.createId.mockClear();
  });

  it.each([
    { type: "columns.insert", index: 27, count: 1 },
    { type: "columns.delete", index: 25, count: 2 },
    { type: "rows.insert", index: -1, count: 1 },
    { type: "rows.delete", index: 0, count: 0 }
  ] as const)("rejects strict structural bounds: $type", (command) => {
    const workbook = createBlankWorkbook();
    const result = reduceWorksheetStructureCommand(workbook, {
      ...command,
      sheetId: workbook.activeSheetId
    }, services);
    expect(result).toMatchObject({
      status: "rejected",
      reason: "validation",
      workbook,
      issues: [{ code: expect.stringMatching(/^SHEET_STRUCTURE_/) }]
    });
    expect(services.createId).not.toHaveBeenCalled();
  });

  it("rejects protected sheets without mutation", () => {
    let workbook = createBlankWorkbook();
    workbook = setSheetProtection(workbook, workbook.activeSheetId, true);
    const result = reduceWorksheetStructureCommand(workbook, {
      type: "columns.insert",
      sheetId: workbook.activeSheetId,
      index: 0,
      count: 1
    }, services);
    expect(result).toEqual({
      status: "rejected",
      reason: "permission",
      workbook,
      issues: [{ code: "TABLE_PROTECTED", message: "Protected sheets cannot change worksheet structure" }]
    });
  });

  it("commits a valid table-free plane shift", () => {
    let workbook = createBlankWorkbook();
    const sheetId = workbook.activeSheetId;
    workbook = setCellContent(workbook, sheetId, "A1", "Name");
    const result = reduceWorksheetStructureCommand(workbook, {
      type: "columns.insert", sheetId, index: 0, count: 1
    }, services);
    expect(result.status).toBe("committed");
    if (result.status !== "committed") return;
    expect(getCellContent(result.workbook, sheetId, "B1")).toBe("Name");
  });

  it("rejects generic row edits inside a structured table", () => {
    let workbook = createBlankWorkbook();
    const sheetId = workbook.activeSheetId;
    const table: StructuredTable = {
      id: "table-1",
      name: "TableOne",
      sheetId,
      range: { start: { row: 0, column: 0 }, end: { row: 1, column: 0 } },
      headerRow: true,
      totalsRow: false,
      columns: [{ id: "table-column-1", name: "Name", sheetColumn: 0 }],
      rowIds: ["table-row-1"]
    };
    workbook = { ...workbook, tables: [table] };

    const result = reduceWorksheetStructureCommand(workbook, {
      type: "rows.insert",
      sheetId,
      index: 1,
      count: 1
    }, services);

    expect(result).toEqual({
      status: "rejected",
      reason: "validation",
      workbook,
      issues: [{
        code: "TABLE_PARTIAL_STRUCTURAL_EDIT",
        message: "Worksheet row edits cannot change a structured table body"
      }]
    });
    expect(services.createId).not.toHaveBeenCalled();
  });

  it("recognizes only worksheet structure command types", () => {
    expect(isWorksheetStructureCommand({
      type: "columns.insert",
      sheetId: "sheet-1",
      index: 0,
      count: 1,
      expandTableIds: ["table-1"]
    })).toBe(true);
    expect(isWorksheetStructureCommand({ type: "columns.resize" })).toBe(false);
  });

  it("shifts later tables and expands only the intended boundary table", () => {
    const workbook = structuredWorkbook();
    const result = reduceWorksheetStructureCommand(workbook, {
      type: "columns.insert",
      sheetId: workbook.activeSheetId,
      index: 2,
      count: 1,
      expandTableIds: ["table-sales"]
    }, deterministicServices());

    expect(result.status).toBe("committed");
    if (result.status !== "committed") return;
    expect(result.workbook.tables[0].range).toEqual(range(0, 0, 2, 2));
    expect(result.workbook.tables[0].columns.map((column) => column.name)).toEqual([
      "Region", "Sales", "Column3"
    ]);
    expect(result.workbook.tables[1].range).toEqual(range(0, 4, 2, 5));
    expect(result.workbook.sheets[0].cells.C1).toBe("Column3");
  });

  it("expands an internal insertion without boundary context", () => {
    const workbook = structuredWorkbook();
    const result = reduceWorksheetStructureCommand(workbook, {
      type: "columns.insert",
      sheetId: workbook.activeSheetId,
      index: 1,
      count: 2
    }, deterministicServices());

    expect(result.status).toBe("committed");
    if (result.status !== "committed") return;
    expect(result.workbook.tables[0].columns.map((column) => [column.id, column.name, column.sheetColumn])).toEqual([
      ["sales-region", "Region", 0],
      ["table-column-1", "Column2", 1],
      ["table-column-2", "Column3", 2],
      ["sales-value", "Sales", 3]
    ]);
  });

  it.each([
    {
      case: "duplicate",
      prepare: (workbook: WorkbookModel) => workbook,
      index: 2,
      expandTableIds: ["table-sales", "table-sales"]
    },
    {
      case: "stale",
      prepare: (workbook: WorkbookModel) => workbook,
      index: 2,
      expandTableIds: ["table-missing"]
    },
    {
      case: "cross-sheet",
      prepare: withCrossSheetTable,
      index: 2,
      expandTableIds: ["table-cross-sheet"]
    },
    {
      case: "non-boundary",
      prepare: (workbook: WorkbookModel) => workbook,
      index: 1,
      expandTableIds: ["table-sales"]
    }
  ])("rejects $case expansion ids before generating ids", ({ prepare, index, expandTableIds }) => {
    const workbook = prepare(structuredWorkbook());
    const insertionServices = deterministicServices();
    const result = reduceWorksheetStructureCommand(workbook, {
      type: "columns.insert",
      sheetId: workbook.activeSheetId,
      index,
      count: 1,
      expandTableIds
    }, insertionServices);

    expect(result).toMatchObject({
      status: "rejected",
      issues: [{ code: "TABLE_EXPANSION_CONTEXT_INVALID" }]
    });
    expect(result.workbook).toBe(workbook);
    expect(insertionServices.createId).not.toHaveBeenCalled();
  });

  it("expands one adjacent boundary table and shifts its neighbor without expanding both", () => {
    const workbook = structuredWorkbook({ costStart: 2 });
    const result = reduceWorksheetStructureCommand(workbook, {
      type: "columns.insert",
      sheetId: workbook.activeSheetId,
      index: 2,
      count: 1,
      expandTableIds: ["table-sales"]
    }, deterministicServices());

    expect(result.status).toBe("committed");
    if (result.status !== "committed") return;
    expect(result.workbook.tables.map((table) => table.range)).toEqual([
      range(0, 0, 2, 2),
      range(0, 3, 2, 4)
    ]);
    expect(result.workbook.tables[1].columns.map((column) => column.sheetColumn)).toEqual([3, 4]);
    expect(result.workbook.tables[1].columns).toHaveLength(2);
  });

  it("rejects projected overlap at an adjacent boundary before generating ids", () => {
    const workbook = structuredWorkbook({ costStart: 2 });
    const insertionServices = deterministicServices();
    const result = reduceWorksheetStructureCommand(workbook, {
      type: "columns.insert",
      sheetId: workbook.activeSheetId,
      index: 2,
      count: 1,
      expandTableIds: ["table-sales", "table-costs"]
    }, insertionServices);

    expect(result).toMatchObject({ status: "rejected", issues: [{ code: "TABLE_RANGE_OVERLAP" }] });
    expect(result.workbook).toBe(workbook);
    expect(insertionServices.createId).not.toHaveBeenCalled();
  });

  it.each([
    {
      case: "expands a selected left boundary",
      index: 0,
      expandTableIds: ["table-sales"],
      expectedRange: range(0, 0, 2, 2),
      expectedColumns: [
        ["table-column-1", "Column1", 0],
        ["sales-region", "Region", 1],
        ["sales-value", "Sales", 2]
      ],
      generated: 1
    },
    {
      case: "shifts an unselected left boundary",
      index: 0,
      expandTableIds: undefined,
      expectedRange: range(0, 1, 2, 2),
      expectedColumns: [
        ["sales-region", "Region", 1],
        ["sales-value", "Sales", 2]
      ],
      generated: 0
    },
    {
      case: "leaves an unselected right boundary unchanged",
      index: 2,
      expandTableIds: undefined,
      expectedRange: range(0, 0, 2, 1),
      expectedColumns: [
        ["sales-region", "Region", 0],
        ["sales-value", "Sales", 1]
      ],
      generated: 0
    }
  ])("handles boundary insertion: $case", ({ index, expandTableIds, expectedRange, expectedColumns, generated }) => {
    const workbook = structuredWorkbook();
    const insertionServices = deterministicServices();
    const result = reduceWorksheetStructureCommand(workbook, {
      type: "columns.insert",
      sheetId: workbook.activeSheetId,
      index,
      count: 1,
      ...(expandTableIds ? { expandTableIds } : {})
    }, insertionServices);

    expect(result.status).toBe("committed");
    if (result.status !== "committed") return;
    expect(result.workbook.tables[0].range).toEqual(expectedRange);
    expect(result.workbook.tables[0].columns.map((column) => [column.id, column.name, column.sheetColumn]))
      .toEqual(expectedColumns);
    expect(insertionServices.createId).toHaveBeenCalledTimes(generated);
  });

  it("keeps insertion metadata on existing column ids and gives new columns no optional roles", () => {
    const workbook = structuredWorkbook();
    const original = workbook.tables[0];
    const result = reduceWorksheetStructureCommand(workbook, {
      type: "columns.insert",
      sheetId: workbook.activeSheetId,
      index: 1,
      count: 2
    }, deterministicServices());

    expect(result.status).toBe("committed");
    if (result.status !== "committed") return;
    const table = result.workbook.tables[0];
    expect(table.columns.find((column) => column.id === "sales-region")).toEqual(original.columns[0]);
    expect(table.columns.find((column) => column.id === "sales-value")).toEqual({
      ...original.columns[1],
      sheetColumn: 3
    });
    expect(table).toMatchObject({
      keyColumnId: original.keyColumnId,
      sort: original.sort,
      filter: original.filter,
      style: original.style
    });
    expect(table.columns.filter((column) => column.id.startsWith("table-column-"))).toEqual([
      { id: "table-column-1", name: "Column2", sheetColumn: 1 },
      { id: "table-column-2", name: "Column3", sheetColumn: 2 }
    ]);
  });

  it("advances insertion names until their normalized headers are unique", () => {
    const workbook = structuredWorkbook({ salesNames: ["column2", "Sales"] });
    const result = reduceWorksheetStructureCommand(workbook, {
      type: "columns.insert",
      sheetId: workbook.activeSheetId,
      index: 1,
      count: 2
    }, deterministicServices());

    expect(result.status).toBe("committed");
    if (result.status !== "committed") return;
    expect(result.workbook.tables[0].columns.map((column) => column.name)).toEqual([
      "column2", "Column3", "Column4", "Sales"
    ]);
    expect(result.workbook.sheets[0].cells).toMatchObject({ B1: "Column3", C1: "Column4" });
  });

  it("generates insertion ids in workbook-table and ascending physical-column order", () => {
    const workbook = structuredWorkbook();
    const lowerTable: StructuredTable = {
      ...workbook.tables[0],
      id: "table-lower",
      name: "LowerTable",
      range: range(5, 0, 7, 1),
      columns: workbook.tables[0].columns.map((column, index) => ({
        ...column,
        id: `lower-column-${index + 1}`
      })),
      rowIds: ["lower-row-1", "lower-row-2"]
    };
    const orderedWorkbook = { ...workbook, tables: [lowerTable, workbook.tables[0], workbook.tables[1]] };
    const result = reduceWorksheetStructureCommand(orderedWorkbook, {
      type: "columns.insert",
      sheetId: orderedWorkbook.activeSheetId,
      index: 1,
      count: 2
    }, deterministicServices());

    expect(result.status).toBe("committed");
    if (result.status !== "committed") return;
    expect(result.workbook.tables[0].columns.slice(1, 3).map((column) => column.id)).toEqual([
      "table-column-1", "table-column-2"
    ]);
    expect(result.workbook.tables[1].columns.slice(1, 3).map((column) => column.id)).toEqual([
      "table-column-3", "table-column-4"
    ]);
  });

  it("rejects an insertion id collision without mutating the workbook", () => {
    const workbook = structuredWorkbook();
    const insertionServices = deterministicServices(["sales-region"]);
    const result = reduceWorksheetStructureCommand(workbook, {
      type: "columns.insert",
      sheetId: workbook.activeSheetId,
      index: 1,
      count: 1
    }, insertionServices);

    expect(result).toMatchObject({ status: "rejected", issues: [{ code: "TABLE_COLUMN_ID_CONFLICT" }] });
    expect(result.workbook).toBe(workbook);
    expect(insertionServices.createId).toHaveBeenCalledTimes(1);
  });

  it.each([
    { label: "blank", ids: ["   "] },
    { label: "duplicate", ids: ["new-column", "new-column"] }
  ])("rejects a $label generated insertion id before publishing", ({ ids }) => {
    const workbook = structuredWorkbook();
    const insertionServices = deterministicServices(ids);
    const result = reduceWorksheetStructureCommand(workbook, {
      type: "columns.insert",
      sheetId: workbook.activeSheetId,
      index: 1,
      count: ids.length
    }, insertionServices);

    expect(result).toMatchObject({
      status: "rejected",
      issues: [{ code: "TABLE_COLUMN_ID_CONFLICT" }]
    });
    expect(result.workbook).toBe(workbook);
  });

  it("creates insertion metadata without writing a header when the table header row is disabled", () => {
    const workbook = structuredWorkbook({ headerRow: false });
    const result = reduceWorksheetStructureCommand(workbook, {
      type: "columns.insert",
      sheetId: workbook.activeSheetId,
      index: 1,
      count: 1
    }, deterministicServices());

    expect(result.status).toBe("committed");
    if (result.status !== "committed") return;
    expect(result.workbook.tables[0].columns.map((column) => column.name)).toEqual([
      "Region", "Column2", "Sales"
    ]);
    expect(result.workbook.sheets[0].cells.B1).toBeUndefined();
    expect(result.workbook.sheets[0].cells.C1).toBe("Sales");
  });

  it("preserves ordinary A1 formula insertion rewriting while projecting tables", () => {
    let workbook = structuredWorkbook();
    workbook = setCellContent(workbook, workbook.activeSheetId, "G1", "=B2+E2");
    const result = reduceWorksheetStructureCommand(workbook, {
      type: "columns.insert",
      sheetId: workbook.activeSheetId,
      index: 1,
      count: 2
    }, deterministicServices());

    expect(result.status).toBe("committed");
    if (result.status !== "committed") return;
    expect(getCellContent(result.workbook, workbook.activeSheetId, "I1")).toBe("=D2+G2");
  });

  it("deletes an internal table column and preserves surviving logical identities", () => {
    const base = threeColumnCalculatedWorkbook();
    const workbook = {
      ...base,
      tables: [{
        ...base.tables[0],
        columns: base.tables[0].columns.map((column) =>
          column.id === "column-c" ? { ...column, calculatedFormula: "=C2*2" } : column
        )
      }]
    };
    const result = reduceWorksheetStructureCommand(workbook, {
      type: "columns.delete", sheetId: workbook.activeSheetId, index: 1, count: 1
    }, deterministicServices());

    expect(result.status).toBe("committed");
    if (result.status !== "committed") return;
    const table = result.workbook.tables[0];
    expect(table.range).toEqual(range(0, 0, 3, 1));
    expect(table.columns.map((column) => [column.id, column.sheetColumn])).toEqual([
      ["column-a", 0], ["column-c", 1]
    ]);
    expect(table.columns[1]).toMatchObject({
      id: "column-c",
      name: "C",
      calculatedFormula: "=B2*2"
    });
    expect(table.rowIds).toEqual(workbook.tables[0].rowIds);
    expect(table.sort).toEqual([]);
    expect(table.filter).toBeUndefined();
    expect(table.keyColumnId).toBeUndefined();
  });

  it("rejects atomically when any affected table would lose every column", () => {
    const workbook = adjacentSingleColumnTables();
    const deletionServices = deterministicServices();
    const result = reduceWorksheetStructureCommand(workbook, {
      type: "columns.delete", sheetId: workbook.activeSheetId, index: 0, count: 2
    }, deletionServices);

    expect(result).toMatchObject({
      status: "rejected",
      issues: [{ code: "TABLE_PARTIAL_STRUCTURAL_EDIT" }]
    });
    expect(result.workbook).toBe(workbook);
    expect(deletionServices.createId).not.toHaveBeenCalled();
  });

  it.each([
    {
      case: "wholly before a table",
      workbook: threeColumnCalculatedWorkbook(3),
      index: 0,
      count: 2,
      expectedRange: range(0, 1, 3, 3),
      expectedColumns: [["column-a", 1], ["column-b", 2], ["column-c", 3]]
    },
    {
      case: "ending exactly at a table start",
      workbook: threeColumnCalculatedWorkbook(3),
      index: 1,
      count: 2,
      expectedRange: range(0, 1, 3, 3),
      expectedColumns: [["column-a", 1], ["column-b", 2], ["column-c", 3]]
    },
    {
      case: "wholly after a table",
      workbook: threeColumnCalculatedWorkbook(),
      index: 5,
      count: 1,
      expectedRange: range(0, 0, 3, 2),
      expectedColumns: [["column-a", 0], ["column-b", 1], ["column-c", 2]]
    },
    {
      case: "starting exactly at a table exclusive end",
      workbook: threeColumnCalculatedWorkbook(),
      index: 3,
      count: 1,
      expectedRange: range(0, 0, 3, 2),
      expectedColumns: [["column-a", 0], ["column-b", 1], ["column-c", 2]]
    },
    {
      case: "entering a table from the left",
      workbook: threeColumnCalculatedWorkbook(2),
      index: 1,
      count: 2,
      expectedRange: range(0, 1, 3, 2),
      expectedColumns: [["column-b", 1], ["column-c", 2]]
    },
    {
      case: "leaving a table to the right",
      workbook: threeColumnCalculatedWorkbook(),
      index: 1,
      count: 3,
      expectedRange: range(0, 0, 3, 0),
      expectedColumns: [["column-a", 0]]
    }
  ])("deletes columns $case with half-open interval projection", ({
    workbook, index, count, expectedRange, expectedColumns
  }) => {
    const result = reduceWorksheetStructureCommand(workbook, {
      type: "columns.delete", sheetId: workbook.activeSheetId, index, count
    }, deterministicServices());

    expect(result.status).toBe("committed");
    if (result.status !== "committed") return;
    expect(result.workbook.tables[0].range).toEqual(expectedRange);
    expect(result.workbook.tables[0].columns.map((column) => [column.id, column.sheetColumn]))
      .toEqual(expectedColumns);
    expect(result.workbook.tables[0].rowIds).toEqual(workbook.tables[0].rowIds);
  });

  it("deletes columns spanning multiple tables without disturbing surviving identities", () => {
    const workbook = structuredWorkbook();
    const result = reduceWorksheetStructureCommand(workbook, {
      type: "columns.delete", sheetId: workbook.activeSheetId, index: 1, count: 3
    }, deterministicServices());

    expect(result.status).toBe("committed");
    if (result.status !== "committed") return;
    expect(result.workbook.tables.map((table) => table.range)).toEqual([
      range(0, 0, 2, 0),
      range(0, 1, 2, 1)
    ]);
    expect(result.workbook.tables.map((table) => table.columns.map((column) => [column.id, column.sheetColumn])))
      .toEqual([
        [["sales-region", 0]],
        [["costs-value", 1]]
      ]);
    expect(result.workbook.tables[0]).toMatchObject({
      rowIds: workbook.tables[0].rowIds,
      keyColumnId: "sales-region",
      filter: workbook.tables[0].filter,
      sort: []
    });
  });

  it("deletes only removed sort entries and retains a filter whose recursive leaves survive", () => {
    const workbook = threeColumnCalculatedWorkbook();
    const filter: StructuredTable["filter"] = {
      kind: "logical",
      operator: "or",
      operands: [
        {
          kind: "comparison",
          columnId: "column-a",
          operator: "eq",
          value: { type: "string", value: "A" }
        },
        { kind: "blank", columnId: "column-c", operator: "isNotBlank" }
      ]
    };
    const table = {
      ...workbook.tables[0],
      sort: [
        { columnId: "column-a", direction: "asc" as const },
        { columnId: "column-b", direction: "desc" as const },
        { columnId: "column-c", direction: "asc" as const }
      ],
      filter
    };
    const prepared = { ...workbook, tables: [table] };
    const result = reduceWorksheetStructureCommand(prepared, {
      type: "columns.delete", sheetId: prepared.activeSheetId, index: 1, count: 1
    }, deterministicServices());

    expect(result.status).toBe("committed");
    if (result.status !== "committed") return;
    expect(result.workbook.tables[0].sort).toEqual([
      { columnId: "column-a", direction: "asc" },
      { columnId: "column-c", direction: "asc" }
    ]);
    expect(result.workbook.tables[0].filter).toBe(filter);
  });

  it("rejects a projected table overlap before mutating worksheet planes", () => {
    const workbook = threeColumnCalculatedWorkbook();
    const escapedColumnTable: StructuredTable = {
      ...workbook.tables[0],
      columns: [workbook.tables[0].columns[0], {
        ...workbook.tables[0].columns[2],
        sheetColumn: 4
      }]
    };
    const neighbor: StructuredTable = {
      id: "table-neighbor",
      name: "NeighborTable",
      sheetId: workbook.activeSheetId,
      range: range(0, 3, 3, 3),
      headerRow: true,
      totalsRow: false,
      columns: [{ id: "neighbor-column", name: "Neighbor", sheetColumn: 3 }],
      rowIds: ["neighbor-row-1", "neighbor-row-2", "neighbor-row-3"]
    };
    const prepared = { ...workbook, tables: [escapedColumnTable, neighbor] };
    const result = reduceWorksheetStructureCommand(prepared, {
      type: "columns.delete", sheetId: prepared.activeSheetId, index: 1, count: 1
    }, deterministicServices());

    expect(result).toMatchObject({ status: "rejected", issues: [{ code: "TABLE_RANGE_OVERLAP" }] });
    expect(result.workbook).toBe(prepared);
  });

  it("shifts tables for generic row edits before them and rejects edits inside them", () => {
    const workbook = tableStartingAtRowFive();
    const before = reduceWorksheetStructureCommand(workbook, {
      type: "rows.insert", sheetId: workbook.activeSheetId, index: 2, count: 2
    }, deterministicServices());
    expect(before.status).toBe("committed");
    if (before.status !== "committed") return;
    expect(before.workbook.tables[0].range.start.row).toBe(6);
    expect(before.workbook.tables[0].rowIds).toEqual(workbook.tables[0].rowIds);
    expect(before.workbook.tables[0].columns[1].calculatedFormula).toBe("=A8*2");

    const inside = reduceWorksheetStructureCommand(workbook, {
      type: "rows.delete", sheetId: workbook.activeSheetId, index: 5, count: 1
    }, deterministicServices());
    expect(inside).toMatchObject({
      status: "rejected",
      issues: [{ code: "TABLE_PARTIAL_STRUCTURAL_EDIT" }]
    });
    expect(inside.workbook).toBe(workbook);
  });

  it.each([
    { case: "inserts at the table start", type: "rows.insert" as const, index: 4, count: 1, start: 5 },
    { case: "inserts after the table", type: "rows.insert" as const, index: 7, count: 2, start: 4 },
    { case: "deletes wholly before the table", type: "rows.delete" as const, index: 1, count: 2, start: 2 },
    { case: "deletes wholly after the table", type: "rows.delete" as const, index: 7, count: 1, start: 4 }
  ])("allows a generic row edit that $case", ({ type, index, count, start }) => {
    const workbook = tableStartingAtRowFive();
    const result = reduceWorksheetStructureCommand(workbook, {
      type, sheetId: workbook.activeSheetId, index, count
    }, deterministicServices());

    expect(result.status).toBe("committed");
    if (result.status !== "committed") return;
    expect(result.workbook.tables[0].range).toEqual(range(start, 0, start + 2, 1));
    expect(result.workbook.tables[0].rowIds).toEqual(workbook.tables[0].rowIds);
  });

  it.each([
    { case: "inserts inside the table", type: "rows.insert" as const, index: 5, count: 1 },
    { case: "inserts at the inclusive table end", type: "rows.insert" as const, index: 6, count: 1 },
    { case: "deletes into the table from above", type: "rows.delete" as const, index: 3, count: 2 },
    { case: "deletes out of the table to below", type: "rows.delete" as const, index: 6, count: 2 }
  ])("rejects atomically when a generic row edit $case", ({ type, index, count }) => {
    const workbook = tableStartingAtRowFive();
    const result = reduceWorksheetStructureCommand(workbook, {
      type, sheetId: workbook.activeSheetId, index, count
    }, deterministicServices());

    expect(result).toMatchObject({
      status: "rejected",
      issues: [{ code: "TABLE_PARTIAL_STRUCTURAL_EDIT" }]
    });
    expect(result.workbook).toBe(workbook);
  });

  it("rewrites canonical calculated formulas before later row regeneration", () => {
    const workbook = calculatedFormulaWorkbook("=B2*$C$2");
    const inserted = reduceWorksheetStructureCommand(workbook, {
      type: "columns.insert", sheetId: workbook.activeSheetId, index: 1, count: 1
    }, deterministicServices());

    expect(inserted.status).toBe("committed");
    if (inserted.status !== "committed") return;
    expect(inserted.workbook.tables[0].columns.find((column) => column.id === "calculated")?.calculatedFormula)
      .toBe("=C2*$D$2");
  });

  it("rewrites calculated formulas in every table while leaving non-formula metadata unchanged", () => {
    const primary = calculatedFormulaWorkbook("=B2*$C$2");
    const withOther = addSheet(primary, "Other");
    const otherSheetId = withOther.activeSheetId;
    const otherTable: StructuredTable = {
      id: "table-other-formulas",
      name: "OtherFormulaTable",
      sheetId: otherSheetId,
      range: range(0, 0, 2, 1),
      headerRow: true,
      totalsRow: false,
      columns: [
        {
          id: "other-formula",
          name: "Formula",
          sheetColumn: 0,
          calculatedFormula: "=Sheet1!B2+A2"
        },
        {
          id: "other-text",
          name: "Text",
          sheetColumn: 1,
          calculatedFormula: "B2"
        }
      ],
      rowIds: ["other-row-1", "other-row-2"]
    };
    const workbook = {
      ...withOther,
      activeSheetId: primary.activeSheetId,
      tables: [...primary.tables, otherTable]
    };
    const result = reduceWorksheetStructureCommand(workbook, {
      type: "columns.insert", sheetId: primary.activeSheetId, index: 1, count: 1
    }, deterministicServices());

    expect(result.status).toBe("committed");
    if (result.status !== "committed") return;
    expect(result.workbook.tables[1].columns.map((column) => column.calculatedFormula)).toEqual([
      "=Sheet1!C2+A2",
      "B2"
    ]);
  });
});
