import { describe, expect, it } from "vitest";
import {
  createBlankWorkbook,
  getCellComment,
  getCellContent,
  getCellFormat,
  getCellReadOnly,
  getCellValidation,
  isColumnHidden
} from "../../lib/workbook";
import type { StructuredTable, WorkbookModel } from "../../types";
import { createWorkbookSession } from "../../core/workbook/WorkbookSession";
import type { TableIntent } from "../core/types";

describe("WorkbookTableSession", () => {
  it("memoizes table sessions and projects table cells without copying a value matrix", () => {
    const parent = createWorkbookSession({ workbook: workbookFixture() });
    const first = parent.table("table-people");
    const second = parent.table("table-people");

    expect(second).toBe(first);
    expect(first.getSnapshot()).toBe(first.getSnapshot());
    expect(first.getSnapshot()).not.toHaveProperty("values");
    expect(first.getSnapshot().rows.map((row) => row.id)).toEqual(["row-ada", "row-grace"]);
    expect(first.getSnapshot().columns.map((column) => column.id)).toEqual(["column-name", "column-score"]);
    expect(first.getSnapshot().getCell("row-grace", "column-score")).toMatchObject({
      storedValue: 20,
      evaluatedValue: 20,
      displayValue: "20",
      editable: true
    });

    parent.destroy();
  });

  it("maps edits and row operations to one atomic parent publication each", async () => {
    const parent = createWorkbookSession({
      workbook: workbookFixture(),
      createId: (() => {
        let next = 0;
        return (kind: string) => `${kind}-new-${++next}`;
      })()
    });
    const table = parent.table("table-people");
    let parentPublications = 0;
    let tablePublications = 0;
    let secondTableSubscriberPublications = 0;
    parent.subscribe(() => { parentPublications += 1; });
    table.subscribe(() => { tablePublications += 1; });
    table.subscribe(() => { secondTableSubscriberPublications += 1; });

    expect(await table.dispatch({
      type: "edit-cells",
      edits: [
        { rowId: "row-ada", columnId: "column-name", rawText: "Ada Lovelace" },
        { rowId: "row-ada", columnId: "column-score", rawText: "15" }
      ]
    })).toMatchObject({ status: "committed", changed: true });
    expect(parentPublications).toBe(1);
    expect(tablePublications).toBe(1);
    expect(secondTableSubscriberPublications).toBe(1);
    expect(getCellContent(parent.getSnapshot().workbook, "sheet-1", "A2")).toBe("Ada Lovelace");
    expect(getCellContent(parent.getSnapshot().workbook, "sheet-1", "B2")).toBe(15);

    expect(await table.dispatch({ type: "insert-rows", count: 1, afterRowId: "row-ada" }))
      .toMatchObject({ status: "committed", changed: true });
    expect(parentPublications).toBe(2);
    expect(tablePublications).toBe(2);
    expect(secondTableSubscriberPublications).toBe(2);
    expect(table.getSnapshot().rows).toHaveLength(3);

    const inserted = table.getSnapshot().rows.find((row) => row.id.startsWith("table-row-new-"));
    expect(inserted).toBeDefined();
    expect(await table.dispatch({ type: "delete-rows", rowIds: [inserted!.id] }))
      .toMatchObject({ status: "committed", changed: true });
    expect(parentPublications).toBe(3);
    expect(tablePublications).toBe(3);
    expect(secondTableSubscriberPublications).toBe(3);
    expect(table.getSnapshot().rows.map((row) => row.id)).toEqual(["row-ada", "row-grace"]);

    parent.destroy();
  });

  it("reads current values lazily and retains column definitions across value-only revisions", () => {
    const parent = createWorkbookSession({ workbook: workbookFixture() });
    const table = parent.table("table-people");
    const before = table.getSnapshot();
    const scoreColumn = before.columns[1];

    expect(parent.dispatch({ type: "cell.set", sheetId: "sheet-1", address: "B2", input: "30" }))
      .toMatchObject({ status: "committed" });
    const afterEdit = table.getSnapshot();
    expect(afterEdit.revision).not.toBe(before.revision);
    expect(afterEdit.getCell("row-ada", "column-score").storedValue).toBe(30);
    expect(afterEdit.columns[1]).toBe(scoreColumn);

    expect(parent.dispatch({
      type: "table.renameColumn",
      tableId: "table-people",
      columnId: "column-score",
      name: "Points"
    })).toMatchObject({ status: "committed" });
    expect(table.getSnapshot().columns[1]).not.toBe(scoreColumn);
    expect(table.getSnapshot().columns[1].header).toBe("Points");

    parent.destroy();
  });

  it("preserves row IDs through sorting, filtering, and undo", async () => {
    const parent = createWorkbookSession({ workbook: workbookFixture() });
    const table = parent.table("table-people");

    expect(await table.dispatch({
      type: "set-sorting",
      sorting: [{ columnId: "column-score", direction: "desc" }]
    })).toMatchObject({ status: "committed" });
    expect(table.getSnapshot().rows.map((row) => row.id)).toEqual(["row-grace", "row-ada"]);
    expect(table.getSnapshot().getCell("row-grace", "column-score").storedValue).toBe(20);

    expect(await table.undo()).toMatchObject({ status: "committed" });
    expect(table.getSnapshot().rows.map((row) => row.id)).toEqual(["row-ada", "row-grace"]);
    expect(table.getSnapshot().getCell("row-grace", "column-score").storedValue).toBe(20);

    expect(await table.dispatch({
      type: "set-filter",
      filter: {
        kind: "comparison",
        columnId: "column-score",
        operator: "gt",
        value: { type: "number", value: 10 }
      }
    })).toMatchObject({ status: "committed" });
    expect(table.getSnapshot().rows.map((row) => row.id)).toEqual(["row-grace"]);
    expect(table.getSnapshot().rowCount).toBe(1);
    expect(table.getSnapshot().totalRowCount).toEqual({ kind: "known", value: 2 });

    parent.destroy();
  });

  it("validates metadata IDs before dispatch and commits a valid metadata batch once", async () => {
    const parent = createWorkbookSession({ workbook: workbookFixture() });
    const table = parent.table("table-people");
    const before = parent.getSnapshot();

    expect(await table.dispatch({
      type: "update-cell-metadata",
      updates: [
        { rowId: "row-ada", columnId: "column-name", patch: { comment: "must not commit" } },
        { rowId: "missing", columnId: "column-name", patch: { readOnly: true } }
      ]
    })).toMatchObject({ status: "rejected", reason: "validation" });
    expect(parent.getSnapshot()).toBe(before);
    expect(getCellComment(parent.getSnapshot().workbook, "sheet-1", "A2")).toBeNull();

    let publications = 0;
    parent.subscribe(() => { publications += 1; });
    expect(await table.dispatch({
      type: "update-cell-metadata",
      updates: [
        {
          rowId: "row-ada",
          columnId: "column-name",
          patch: { formula: "=B2", comment: "calculated", format: { bold: true } }
        },
        {
          rowId: "row-grace",
          columnId: "column-score",
          patch: {
            validation: { kind: "number", min: 0, max: 100 },
            readOnly: true
          }
        }
      ]
    })).toMatchObject({ status: "committed", changed: true });
    expect(publications).toBe(1);
    const workbook = parent.getSnapshot().workbook;
    expect(getCellContent(workbook, "sheet-1", "A2")).toBe("=B2");
    expect(getCellComment(workbook, "sheet-1", "A2")).toBe("calculated");
    expect(getCellFormat(workbook, "sheet-1", "A2")).toMatchObject({ bold: true });
    expect(getCellValidation(workbook, "sheet-1", "B3")).toEqual({
      type: "number",
      min: 0,
      max: 100,
      allowBlank: undefined
    });
    expect(getCellReadOnly(workbook, "sheet-1", "B3")).toBe(true);

    parent.destroy();
  });

  it("routes selection, clearing, resizing, visibility, and replacement through stable IDs", async () => {
    const parent = createWorkbookSession({ workbook: workbookFixture() });
    const table = parent.table("table-people");

    expect(await table.dispatch({
      type: "set-selection",
      selection: {
        anchor: { rowId: "row-ada", columnId: "column-name" },
        focus: { rowId: "row-grace", columnId: "column-score" }
      }
    })).toMatchObject({ status: "committed" });
    expect(parent.getSnapshot().selection).toEqual({
      start: { row: 1, column: 0 },
      end: { row: 2, column: 1 }
    });
    expect(table.getSnapshot().selection).not.toBeNull();
    expect(await table.dispatch({ type: "set-selection", selection: null }))
      .toMatchObject({ status: "committed", changed: true });
    expect(table.getSnapshot().selection).toBeNull();

    expect(await table.dispatch({
      type: "clear-cells",
      cells: [{ rowId: "row-ada", columnId: "column-name" }]
    })).toMatchObject({ status: "committed" });
    expect(getCellContent(parent.getSnapshot().workbook, "sheet-1", "A2")).toBeNull();

    expect(await table.dispatch({ type: "resize-column", columnId: "column-score", width: 144 }))
      .toMatchObject({ status: "committed" });
    expect(await table.dispatch({
      type: "set-column-visibility",
      columnId: "column-score",
      visible: false
    })).toMatchObject({ status: "committed" });
    expect(table.getSnapshot().state.columnWidths["column-score"]).toBe(144);
    expect(isColumnHidden(parent.getSnapshot().workbook, "sheet-1", 1)).toBe(true);

    const replacement = workbookFixture({ B2: 99 });
    expect(parent.replaceWorkbook(replacement, { history: "reset", origin: "external" }))
      .toMatchObject({ status: "committed" });
    expect(parent.table("table-people")).toBe(table);
    expect(table.getSnapshot().getCell("row-ada", "column-score").storedValue).toBe(99);

    parent.destroy();
  });

  it("keeps local view state local and exposes unsupported features honestly", async () => {
    const parent = createWorkbookSession({ workbook: workbookFixture() });
    const table = parent.table("table-people");
    let parentPublications = 0;
    parent.subscribe(() => { parentPublications += 1; });

    expect(await table.dispatch({
      type: "set-column-pinning",
      columnId: "column-name",
      pin: "left"
    })).toMatchObject({ status: "committed", changed: true });
    expect(table.getSnapshot().state.columnPinning.left).toEqual(["column-name"]);
    expect(parentPublications).toBe(0);

    expect(await table.dispatch({
      type: "set-column-order",
      columnIds: ["column-name", "column-name"]
    })).toMatchObject({ status: "rejected", reason: "validation" });
    expect(table.getSnapshot().state.columnOrder).toEqual(["column-name", "column-score"]);

    for (const intent of [
      { type: "set-grouping", grouping: [{ columnId: "column-name" }] },
      {
        type: "set-aggregates",
        aggregates: [{ id: "sum-score", columnId: "column-score", function: "sum" }]
      },
      { type: "set-pagination", pagination: { kind: "offset", offset: 0, limit: 25 } }
    ] satisfies readonly TableIntent[]) {
      expect(await table.dispatch(intent)).toMatchObject({ status: "rejected", reason: "unsupported" });
    }

    parent.destroy();
  });

  it("turns a converted table into an error snapshot without breaking the parent", () => {
    const parent = createWorkbookSession({ workbook: workbookFixture() });
    const table = parent.table("table-people");

    expect(parent.dispatch({ type: "table.convertToRange", tableId: "table-people" }))
      .toMatchObject({ status: "committed" });
    expect(table.getSnapshot()).toMatchObject({
      status: { phase: "error", message: "Structured table does not exist" },
      rowCount: 0
    });
    expect(parent.getSnapshot().workbook.tables).toEqual([]);

    parent.destroy();
  });

  it("destroying a child leaves the parent usable while destroying the parent tears down children", async () => {
    const parent = createWorkbookSession({ workbook: workbookFixture() });
    const child = parent.table("table-people");
    child.destroy();

    expect(parent.dispatch({ type: "cell.set", sheetId: "sheet-1", address: "D1", input: "alive" }))
      .toMatchObject({ status: "committed" });
    expect(await child.dispatch({ type: "refresh" })).toMatchObject({
      status: "rejected",
      reason: "unsupported"
    });

    const replacement = parent.table("table-people");
    expect(replacement).not.toBe(child);
    parent.destroy();
    expect(await replacement.dispatch({ type: "refresh" })).toMatchObject({
      status: "rejected",
      reason: "unsupported"
    });
  });
});

function workbookFixture(cellOverrides: Record<string, string | number> = {}): WorkbookModel {
  const workbook = createBlankWorkbook();
  const table: StructuredTable = {
    id: "table-people",
    name: "People",
    sheetId: "sheet-1",
    range: { start: { row: 0, column: 0 }, end: { row: 2, column: 1 } },
    headerRow: true,
    totalsRow: false,
    columns: [
      { id: "column-name", name: "Name", sheetColumn: 0, dataType: "text" },
      { id: "column-score", name: "Score", sheetColumn: 1, dataType: "number" }
    ],
    rowIds: ["row-ada", "row-grace"]
  };
  return {
    ...workbook,
    activeSheetId: "sheet-1",
    tables: [table],
    sheets: [{
      ...workbook.sheets[0],
      id: "sheet-1",
      cells: {
        ...workbook.sheets[0].cells,
        A1: "Name",
        B1: "Score",
        A2: "Ada",
        B2: 10,
        A3: "Grace",
        B3: 20,
        ...cellOverrides
      }
    }]
  };
}
