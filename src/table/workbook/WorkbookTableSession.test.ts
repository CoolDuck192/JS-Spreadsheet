import { describe, expect, it, vi } from "vitest";
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

  it("allows row insertion past existing read-only cells but protects the expansion band", async () => {
    const parent = createWorkbookSession({ workbook: workbookFixture() });
    const table = parent.table("table-people");
    expect(await table.dispatch({
      type: "update-cell-metadata",
      updates: [{ rowId: "row-ada", columnId: "column-score", patch: { readOnly: true } }]
    })).toMatchObject({ status: "committed" });

    expect(await table.dispatch({ type: "insert-rows", count: 1 }))
      .toMatchObject({ status: "committed", changed: true });
    expect(table.getSnapshot().rowCount).toBe(3);
    expect(table.getSnapshot().getCell("row-ada", "column-score").editable).toBe(false);
    parent.destroy();

    const protectedBandParent = createWorkbookSession({ workbook: workbookFixture() });
    const protectedBandTable = protectedBandParent.table("table-people");
    expect(protectedBandParent.dispatch({
      type: "range.readOnly.set",
      sheetId: "sheet-1",
      range: { start: { row: 3, column: 1 }, end: { row: 3, column: 1 } },
      readOnly: true
    })).toMatchObject({ status: "committed" });

    expect(await protectedBandTable.dispatch({ type: "insert-rows", count: 1 }))
      .toMatchObject({
        status: "rejected",
        issues: [{ code: "TABLE_PROTECTED" }]
      });
    expect(protectedBandTable.getSnapshot().rowCount).toBe(2);
    protectedBandParent.destroy();
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

  it("keeps read-only metadata attached to stable row IDs through sorting", async () => {
    const parent = createWorkbookSession({ workbook: workbookFixture() });
    const table = parent.table("table-people");
    expect(await table.dispatch({
      type: "update-cell-metadata",
      updates: [{ rowId: "row-grace", columnId: "column-score", patch: { readOnly: true } }]
    })).toMatchObject({ status: "committed" });

    expect(await table.dispatch({
      type: "set-sorting",
      sorting: [{ columnId: "column-score", direction: "desc" }]
    })).toMatchObject({ status: "committed" });

    expect(table.getSnapshot().getCell("row-grace", "column-score").editable).toBe(false);
    expect(table.getSnapshot().getCell("row-ada", "column-score").editable).toBe(true);
    expect(getCellReadOnly(parent.getSnapshot().workbook, "sheet-1", "B2")).toBe(true);
    expect(getCellReadOnly(parent.getSnapshot().workbook, "sheet-1", "B3")).toBe(false);
    parent.destroy();
  });

  it("keeps read-only metadata attached to stable row IDs through deletion", async () => {
    const parent = createWorkbookSession({ workbook: workbookFixture() });
    const table = parent.table("table-people");
    expect(await table.dispatch({
      type: "update-cell-metadata",
      updates: [{ rowId: "row-grace", columnId: "column-score", patch: { readOnly: true } }]
    })).toMatchObject({ status: "committed" });

    expect(await table.dispatch({ type: "delete-rows", rowIds: ["row-ada"] }))
      .toMatchObject({ status: "committed" });

    expect(table.getSnapshot().getCell("row-grace", "column-score").editable).toBe(false);
    expect(getCellReadOnly(parent.getSnapshot().workbook, "sheet-1", "B2")).toBe(true);
    expect(getCellReadOnly(parent.getSnapshot().workbook, "sheet-1", "B3")).toBe(false);
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

  it("exports injection-safe UTF-8 CSV for the current view and complete dataset", async () => {
    const parent = createWorkbookSession({ workbook: exportWorkbookFixture() });
    const table = parent.table("table-export");

    const current = await table.export({
      format: "csv",
      scope: "currentView",
      fileName: " ../unsafe:name?.csv "
    });
    const complete = await table.export({ format: "csv", scope: "completeDataset" });

    const header = "Formula Result,Leading Space,Minus Text,At Text,Tab Text,CR Text,LF Text,Safe Apostrophe,Comma,Quote,Unicode Space,Number,Date\r\n";
    const firstRow = [
      "'=1+1",
      "' +cmd",
      "'-2+3",
      "'@SUM",
      "'\tplain",
      "\"'\rplain\"",
      "\"'\nplain\"",
      "'=already-safe",
      "\"a,b\"",
      "\"a\"\"b\"",
      "'\u2003=unicode",
      "-42",
      "\"Jan 15, 2026\""
    ].join(",") + "\r\n";
    const secondRow = [
      "done",
      "plain",
      "text",
      "safe",
      "tabless",
      "crless",
      "lfless",
      "apostrophe",
      "comma-less",
      "quote-less",
      "unicode-less",
      "7",
      "\"Jan 16, 2026\""
    ].join(",") + "\r\n";

    expect(current.bytes).toBeInstanceOf(Uint8Array);
    expect(current).toMatchObject({
      mediaType: "text/csv;charset=utf-8",
      fileName: "unsafe_name_.csv"
    });
    expect(new TextDecoder().decode(current.bytes)).toBe(header + firstRow);
    expect(new TextDecoder().decode(complete.bytes)).toBe(header + firstRow + secondRow);
    expect([...complete.bytes]).toEqual([...new TextEncoder().encode(header + firstRow + secondRow)]);

    parent.destroy();
  });

  it("exports stable native XLSX bytes only for the complete dataset", async () => {
    const parent = createWorkbookSession({ workbook: workbookFixture() });
    const table = parent.table("table-people");

    await expect(table.export({ format: "xlsx", scope: "currentView" })).rejects.toMatchObject({
      code: "TABLE_EXPORT_SCOPE_UNSUPPORTED",
      issue: { code: "TABLE_EXPORT_SCOPE_UNSUPPORTED" }
    });

    const first = await table.export({ format: "xlsx", scope: "completeDataset" });
    const second = await table.export({ format: "xlsx", scope: "completeDataset" });
    expect(first).toMatchObject({
      bytes: expect.any(Uint8Array),
      mediaType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      fileName: "People.xlsx"
    });
    expect(first.bytes.slice(0, 2)).toEqual(new Uint8Array([0x50, 0x4b]));
    expect(first.bytes).toEqual(second.bytes);

    parent.destroy();
  });

  it("keeps Blob, document, and object URLs outside the session export boundary", async () => {
    const parent = createWorkbookSession({ workbook: workbookFixture() });
    const table = parent.table("table-people");
    const originalBlob = Object.getOwnPropertyDescriptor(globalThis, "Blob");
    const originalDocument = Object.getOwnPropertyDescriptor(globalThis, "document");
    const blobConstructed = vi.fn();
    const createObjectUrl = vi.fn(() => { throw new Error("DOM boundary crossed"); });
    const originalCreateObjectUrl = Object.getOwnPropertyDescriptor(URL, "createObjectURL");

    Object.defineProperty(globalThis, "Blob", {
      configurable: true,
      value: class ForbiddenBlob {
        constructor() {
          blobConstructed();
          throw new Error("Blob boundary crossed");
        }
      }
    });
    Object.defineProperty(globalThis, "document", {
      configurable: true,
      get() { throw new Error("document boundary crossed"); }
    });
    Object.defineProperty(URL, "createObjectURL", { configurable: true, value: createObjectUrl });

    try {
      const csv = await table.export({ format: "csv", scope: "completeDataset" });
      const xlsx = await table.export({ format: "xlsx", scope: "completeDataset" });
      expect(csv.bytes).toBeInstanceOf(Uint8Array);
      expect(xlsx.bytes).toBeInstanceOf(Uint8Array);
      expect(blobConstructed).not.toHaveBeenCalled();
      expect(createObjectUrl).not.toHaveBeenCalled();
    } finally {
      restoreGlobal("Blob", originalBlob);
      restoreGlobal("document", originalDocument);
      restoreProperty(URL, "createObjectURL", originalCreateObjectUrl);
      parent.destroy();
    }
  });

  it("propagates table-only XLSX external-dependency rejection", async () => {
    const workbook = workbookFixture({ B2: "=D1" });
    const parent = createWorkbookSession({ workbook });

    await expect(parent.table("table-people").export({
      format: "xlsx",
      scope: "completeDataset"
    })).rejects.toMatchObject({ code: "TABLE_EXPORT_EXTERNAL_DEPENDENCY" });

    parent.destroy();
  });
});

function exportWorkbookFixture(): WorkbookModel {
  const workbook = createBlankWorkbook();
  const columns = [
    ["formula-result", "Formula Result", "text"],
    ["leading-space", "Leading Space", "text"],
    ["minus-text", "Minus Text", "text"],
    ["at-text", "At Text", "text"],
    ["tab-text", "Tab Text", "text"],
    ["cr-text", "CR Text", "text"],
    ["lf-text", "LF Text", "text"],
    ["safe-apostrophe", "Safe Apostrophe", "text"],
    ["comma", "Comma", "text"],
    ["quote", "Quote", "text"],
    ["unicode-space", "Unicode Space", "text"],
    ["number", "Number", "number"],
    ["date", "Date", "date"]
  ] as const;
  const table: StructuredTable = {
    id: "table-export",
    name: "ExportTable",
    sheetId: "sheet-1",
    range: { start: { row: 0, column: 0 }, end: { row: 2, column: columns.length - 1 } },
    headerRow: true,
    totalsRow: false,
    columns: columns.map(([id, name, dataType], sheetColumn) => ({ id, name, dataType, sheetColumn })),
    rowIds: ["row-first", "row-second"],
    filter: {
      kind: "comparison",
      columnId: "number",
      operator: "eq",
      value: { type: "number", value: -42 }
    }
  };
  const cells = Object.fromEntries(columns.map(([, name], column) => [
    `${String.fromCharCode(65 + column)}1`,
    name
  ]));
  Object.assign(cells, {
    A2: '=\"=1+1\"', B2: " +cmd", C2: "-2+3", D2: "@SUM", E2: "\tplain",
    F2: "\rplain", G2: "\nplain", H2: "'=already-safe", I2: "a,b", J2: 'a"b',
    K2: "\u2003=unicode", L2: -42, M2: 46037,
    A3: '=\"done\"', B3: "plain", C3: "text", D3: "safe", E3: "tabless",
    F3: "crless", G3: "lfless", H3: "apostrophe", I3: "comma-less", J3: "quote-less",
    K3: "unicode-less", L3: 7, M3: 46038
  });
  return {
    ...workbook,
    activeSheetId: "sheet-1",
    tables: [table],
    sheets: [{
      ...workbook.sheets[0],
      id: "sheet-1",
      rowCount: 3,
      columnCount: columns.length,
      cells,
      formats: { M2: { numberFormat: "date" }, M3: { numberFormat: "date" } }
    }]
  };
}

function restoreGlobal(name: string, descriptor: PropertyDescriptor | undefined): void {
  if (descriptor) Object.defineProperty(globalThis, name, descriptor);
  else Reflect.deleteProperty(globalThis, name);
}

function restoreProperty(target: object, name: string, descriptor: PropertyDescriptor | undefined): void {
  if (descriptor) Object.defineProperty(target, name, descriptor);
  else Reflect.deleteProperty(target, name);
}

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
