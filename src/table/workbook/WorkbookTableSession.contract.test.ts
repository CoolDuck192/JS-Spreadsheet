import { expect } from "vitest";
import type { CommandResult } from "../../core/commands/types";
import { createWorkbookSession } from "../../core/workbook/WorkbookSession";
import { createBlankWorkbook } from "../../lib/workbook";
import type { StructuredTable, WorkbookModel } from "../../types";
import type { TableFeature } from "../core/capabilities";
import { defineTableSessionContract } from "../core/session.contract";

defineTableSessionContract("workbook", () => {
  const parent = createWorkbookSession({ workbook: workbookFixture() });
  const session = parent.table("table-people");
  const featureOperations: Partial<Record<TableFeature, () => Promise<CommandResult>>> = {
    sort: () => session.dispatch({
      type: "set-sorting",
      sorting: [{ columnId: "column-score", direction: "asc" }]
    }),
    filter: () => session.dispatch({
      type: "set-filter",
      filter: { kind: "blank", columnId: "column-name", operator: "isNotBlank" }
    }),
    group: () => session.dispatch({ type: "set-grouping", grouping: [{ columnId: "column-name" }] }),
    aggregate: () => session.dispatch({
      type: "set-aggregates",
      aggregates: [{ id: "sum-score", columnId: "column-score", function: "sum" }]
    }),
    pagination: () => session.dispatch({
      type: "set-pagination",
      pagination: { kind: "offset", offset: 0, limit: 25 }
    }),
    edit: () => session.dispatch({
      type: "edit-cells",
      edits: [{ rowId: "row-ada", columnId: "column-name", rawText: "Ada Lovelace" }]
    }),
    bulkEdit: () => session.dispatch({
      type: "edit-cells",
      edits: [
        { rowId: "row-ada", columnId: "column-name", rawText: "Ada Lovelace" },
        { rowId: "row-grace", columnId: "column-name", rawText: "Grace Hopper" }
      ]
    }),
    metadata: () => session.dispatch({
      type: "update-cell-metadata",
      updates: [{ rowId: "row-ada", columnId: "column-name", patch: { comment: "note" } }]
    }),
    validation: () => session.dispatch({
      type: "update-cell-metadata",
      updates: [{
        rowId: "row-ada",
        columnId: "column-score",
        patch: { validation: { kind: "number", min: 0 } }
      }]
    }),
    formula: () => session.dispatch({
      type: "update-cell-metadata",
      updates: [{ rowId: "row-ada", columnId: "column-name", patch: { formula: "=B2" } }]
    }),
    subscription: () => session.dispatch({ type: "refresh" }),
    undo: () => session.dispatch({ type: "undo" }),
    export: async () => {
      const artifact = await session.export({ format: "csv", scope: "completeDataset" });
      expect(artifact).toMatchObject({
        bytes: expect.any(Uint8Array),
        mediaType: "text/csv;charset=utf-8",
        fileName: "People.csv"
      });
      return { status: "committed" as const, revision: session.getSnapshot().revision };
    }
  };
  return {
    session,
    editableCell: { rowId: "row-ada", columnId: "column-name" },
    validRawText: "Ada Lovelace",
    featureOperations,
    cleanup: () => parent.destroy()
  };
});

function workbookFixture(): WorkbookModel {
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
      cells: { A1: "Name", B1: "Score", A2: "Ada", B2: 10, A3: "Grace", B3: 20 }
    }]
  };
}
