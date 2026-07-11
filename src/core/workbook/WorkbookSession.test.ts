import { describe, expect, it, vi } from "vitest";
import type { StructuredTable, WorkbookModel } from "../../types";
import type { FormulaEngine } from "../../lib/formulaEngine";
import {
  addSheet,
  createBlankWorkbook,
  getCellContent,
  setActiveSheet,
  setCellContent
} from "../../lib/workbook";
import type { CommandEnvelope } from "../commands/types";
import type { IdKind } from "../ids";
import type { WorkbookCommand } from "./commands";
import {
  createWorkbookSession,
  type WorkbookDiagnosticEvent
} from "./WorkbookSession";

const origin = { start: { row: 0, column: 0 }, end: { row: 0, column: 0 } } as const;
const block = { start: { row: 0, column: 0 }, end: { row: 1, column: 1 } } as const;

describe("WorkbookSession", () => {
  it("commits, publishes, undoes, and redoes a structured table with exact IDs", () => {
    let workbook = createBlankWorkbook();
    const sheetId = workbook.activeSheetId;
    for (const [address, value] of [
      ["A1", "Name"], ["B1", "Department"], ["C1", "Salary"],
      ["A2", "Ada"], ["B2", "Finance"], ["C2", 100],
      ["A3", "Grace"], ["B3", "Research"], ["C3", 200],
      ["A4", "Linus"], ["B4", "IT"], ["C4", 150]
    ] as const) {
      workbook = setCellContent(workbook, sheetId, address, value);
    }
    let nextId = 0;
    const session = createWorkbookSession({
      workbook,
      createId(kind) {
        nextId += 1;
        return `${kind}-${nextId}`;
      }
    });
    let notifications = 0;
    session.subscribe(() => notifications += 1);

    expect(session.dispatch({
      type: "table.create",
      sheetId,
      range: { start: { row: 0, column: 0 }, end: { row: 3, column: 2 } },
      name: "Employees",
      headerRow: true,
      totalsRow: false
    })).toEqual({ status: "committed", revision: "1", changed: true });
    expect(notifications).toBe(1);
    expect(session.getSnapshot().workbook.tables[0]).toMatchObject({
      id: "table-4",
      name: "Employees",
      columns: [
        { id: "table-column-1", name: "Name" },
        { id: "table-column-2", name: "Department" },
        { id: "table-column-3", name: "Salary" }
      ],
      rowIds: ["table-row-5", "table-row-6", "table-row-7"]
    });
    const committed = JSON.stringify(session.getSnapshot().workbook);

    expect(session.dispatch({ type: "history.undo" })).toMatchObject({ status: "committed", changed: true });
    expect(session.getSnapshot().workbook.tables).toEqual([]);
    expect(session.dispatch({ type: "history.redo" })).toMatchObject({ status: "committed", changed: true });
    expect(JSON.stringify(session.getSnapshot().workbook)).toBe(committed);
    expect(notifications).toBe(3);
  });

  it("rejects an invalid table edit batch without committing an earlier edit", () => {
    let workbook = createBlankWorkbook();
    const sheetId = workbook.activeSheetId;
    workbook = setCellContent(workbook, sheetId, "A1", "Name");
    workbook = setCellContent(workbook, sheetId, "B1", "Status");
    workbook = setCellContent(workbook, sheetId, "A2", "Ada");
    workbook = setCellContent(workbook, sheetId, "B2", "allowed");
    let nextId = 0;
    const session = createWorkbookSession({
      workbook,
      createId(kind) {
        nextId += 1;
        return `${kind}-${nextId}`;
      }
    });
    expect(session.dispatch({
      type: "table.create", sheetId,
      range: { start: { row: 0, column: 0 }, end: { row: 1, column: 1 } },
      name: "People", headerRow: true, totalsRow: false
    })).toMatchObject({ status: "committed" });
    expect(session.dispatch({
      type: "range.validation.set", sheetId,
      range: { start: { row: 1, column: 1 }, end: { row: 1, column: 1 } },
      rule: { type: "list", values: ["allowed"] }
    })).toMatchObject({ status: "committed" });
    const before = session.getSnapshot();
    const table = before.workbook.tables[0];

    expect(session.dispatch({
      type: "table.editCells",
      tableId: table.id,
      edits: [
        { rowId: table.rowIds[0], columnId: table.columns[0].id, rawText: "Changed" },
        { rowId: table.rowIds[0], columnId: table.columns[1].id, rawText: "forbidden" }
      ]
    })).toMatchObject({ status: "rejected", reason: "validation" });
    expect(session.getSnapshot()).toBe(before);
    expect(getCellContent(session.getSnapshot().workbook, sheetId, "A2")).toBe("Ada");
    expect(getCellContent(session.getSnapshot().workbook, sheetId, "B2")).toBe("allowed");
  });

  it("serializes two synchronous dispatches without losing the first edit", () => {
    const diagnostics: WorkbookDiagnosticEvent[] = [];
    const ids = ["command-1", "command-2"];
    const session = createWorkbookSession({
      workbook: createBlankWorkbook(),
      createCommandId: () => ids.shift()!,
      onDiagnostic: (event) => diagnostics.push(event)
    });
    const sheetId = session.getSnapshot().workbook.activeSheetId;

    expect(session.dispatch({ type: "cell.set", sheetId, address: "A1", input: "1" })).toEqual({
      status: "committed",
      revision: "1",
      changed: true
    });
    expect(session.dispatch({ type: "cell.set", sheetId, address: "A2", input: "2" })).toEqual({
      status: "committed",
      revision: "2",
      changed: true
    });

    expect(getCellContent(session.getSnapshot().workbook, sheetId, "A1")).toBe(1);
    expect(getCellContent(session.getSnapshot().workbook, sheetId, "A2")).toBe(2);
    expect(diagnostics.map((event) => event.commandId)).toEqual(["command-1", "command-2"]);
  });

  it("commits a transaction with one history entry, revision, engine update, notification, and diagnostic", () => {
    const engines: TrackingEngine[] = [];
    const diagnostics: WorkbookDiagnosticEvent[] = [];
    const session = createWorkbookSession({
      workbook: createBlankWorkbook(),
      formulaEngineFactory: createTrackingFormulaEngineFactory(engines),
      createCommandId: () => "transaction-command",
      onDiagnostic: (event) => diagnostics.push(event)
    });
    const sheetId = session.getSnapshot().workbook.activeSheetId;
    let notifications = 0;
    session.subscribe(() => notifications += 1);

    const result = session.dispatch({
      type: "transaction",
      commands: [
        { type: "cell.set", sheetId, address: "A1", input: "10" },
        { type: "cell.set", sheetId, address: "A2", input: "20" },
        { type: "selection.set", selection: block }
      ]
    });

    expect(result).toEqual({ status: "committed", revision: "1", changed: true });
    expect(notifications).toBe(1);
    expect(engines[0].updates).toHaveLength(1);
    expect(engines[0].projected).toBe(session.getSnapshot().workbook);
    expect(diagnostics).toHaveLength(1);
    expect(diagnostics[0]).toMatchObject({
      category: "command",
      commandId: "transaction-command",
      metadata: {
        commandType: "transaction",
        outcome: "committed",
        changed: true,
        sheetCount: 1,
        rangeCount: 1,
        cellCount: 6
      }
    });

    expect(session.dispatch({ type: "history.undo" })).toMatchObject({ status: "committed", changed: true });
    expect(getCellContent(session.getSnapshot().workbook, sheetId, "A1")).toBeNull();
    expect(getCellContent(session.getSnapshot().workbook, sheetId, "A2")).toBeNull();
    expect(session.dispatch({ type: "history.undo" })).toEqual({
      status: "committed",
      revision: "2",
      changed: false
    });
  });

  it("rolls back workbook, selection, persistence, history, formula projection, and publication when a child rejects", () => {
    const engines: TrackingEngine[] = [];
    const diagnostics: WorkbookDiagnosticEvent[] = [];
    const session = createWorkbookSession({
      workbook: createBlankWorkbook(),
      formulaEngineFactory: createTrackingFormulaEngineFactory(engines),
      onDiagnostic: (event) => diagnostics.push(event)
    });
    const initial = session.getSnapshot();
    const sheetId = initial.workbook.activeSheetId;
    let notifications = 0;
    session.subscribe(() => notifications += 1);

    const result = session.dispatch({
      type: "transaction",
      commands: [
        { type: "selection.set", selection: block },
        { type: "persistence.status", status: "saving" },
        {
          type: "range.validation.set",
          sheetId,
          range: origin,
          rule: { type: "list", values: ["allowed"] }
        },
        { type: "cell.set", sheetId, address: "A1", input: "blocked" }
      ]
    });

    expect(result).toEqual({
      status: "rejected",
      reason: "validation",
      issues: [{
        code: "validation.failed",
          message: "Choose one of: allowed",
        sheetId,
        address: "A1"
      }]
    });
    expect(session.getSnapshot()).toBe(initial);
    expect(session.getSnapshot()).toMatchObject({
      selection: origin,
      revision: "0",
      canUndo: false,
      canRedo: false,
      persistence: { status: "idle" }
    });
    expect(engines[0].updates).toHaveLength(0);
    expect(engines[0].projected).toBe(initial.workbook);
    expect(notifications).toBe(0);
    expect(diagnostics).toHaveLength(1);
    expect(diagnostics[0]).toMatchObject({ category: "validation", metadata: { outcome: "rejected" } });
  });

  it("preserves a transaction rejection when scratch projection cleanup throws", () => {
    const diagnostics: WorkbookDiagnosticEvent[] = [];
    let factoryCalls = 0;
    const session = createWorkbookSession({
      workbook: createBlankWorkbook(),
      formulaEngineFactory(workbook) {
        factoryCalls += 1;
        const engine = createRawFormulaEngine(workbook);
        return factoryCalls === 2
          ? {
              ...engine,
              destroy() {
                throw new Error("scratch cleanup failed");
              }
            }
          : engine;
      },
      onDiagnostic: (event) => diagnostics.push(event)
    });
    const initial = session.getSnapshot();
    const sheetId = initial.workbook.activeSheetId;

    let result: ReturnType<typeof session.dispatch> | undefined;
    expect(() => {
      result = session.dispatch({
        type: "transaction",
        commands: [
          {
            type: "range.validation.set",
            sheetId,
            range: origin,
            rule: { type: "list", values: ["allowed"] }
          },
          { type: "cell.set", sheetId, address: "A1", input: "blocked" }
        ]
      });
    }).not.toThrow();

    expect(result).toEqual({
      status: "rejected",
      reason: "validation",
      issues: [{
        code: "validation.failed",
        message: "Choose one of: allowed",
        sheetId,
        address: "A1"
      }]
    });
    expect(session.getSnapshot()).toBe(initial);
    expect(diagnostics).toHaveLength(1);
    expect(diagnostics[0]).toMatchObject({
      category: "validation",
      metadata: { outcome: "rejected", changed: false }
    });
  });

  it("lets later transaction children validate against earlier formula dependencies", () => {
    const session = createWorkbookSession({ workbook: createBlankWorkbook() });
    const sheetId = session.getSnapshot().workbook.activeSheetId;
    const b1 = { start: { row: 0, column: 1 }, end: { row: 0, column: 1 } } as const;

    const result = session.dispatch({
      type: "transaction",
      commands: [
        { type: "range.validation.set", sheetId, range: b1, rule: { type: "number", min: 5 } },
        { type: "cell.set", sheetId, address: "A1", input: "5" },
        { type: "cell.set", sheetId, address: "B1", input: "=A1" }
      ]
    });

    expect(result).toMatchObject({ status: "committed", revision: "1", changed: true });
    expect(session.getCellEvaluation(sheetId, "B1")).toBe(5);
  });

  it("suppresses no-op revisions and notifications without discarding redo history", () => {
    const session = createWorkbookSession({ workbook: createBlankWorkbook() });
    const sheetId = session.getSnapshot().workbook.activeSheetId;
    let notifications = 0;
    session.subscribe(() => notifications += 1);

    session.dispatch({ type: "cell.set", sheetId, address: "A1", input: "1" });
    session.dispatch({ type: "cell.set", sheetId, address: "A2", input: "2" });
    session.dispatch({ type: "history.undo" });
    const beforeNoOp = session.getSnapshot();
    const beforeNotifications = notifications;

    expect(session.dispatch({ type: "cell.set", sheetId, address: "A1", input: "1" })).toEqual({
      status: "committed",
      revision: beforeNoOp.revision,
      changed: false
    });
    expect(session.getSnapshot()).toBe(beforeNoOp);
    expect(session.getSnapshot().canRedo).toBe(true);
    expect(notifications).toBe(beforeNotifications);
  });

  it("suppresses repeated stable-ID metadata commands as no-ops", () => {
    const session = createWorkbookSession({ workbook: createBlankWorkbook() });
    const sheetId = session.getSnapshot().workbook.activeSheetId;
    const commands: WorkbookCommand[] = [
      {
        type: "range.conditionalFormat.add",
        sheetId,
        range: origin,
        rule: {
          id: "conditional-1",
          range: origin,
          condition: { type: "greaterThan", value: "1" },
          format: { backgroundColor: "#ffffff" }
        }
      },
      {
        type: "sheet.filter.set",
        sheetId,
        filter: {
          id: "filter-1",
          range: origin,
          column: 0,
          operator: "equals",
          value: "Open"
        }
      },
      {
        type: "sheet.chart.add",
        sheetId,
        chart: {
          id: "chart-1",
          title: "Chart",
          type: "bar",
          range: origin,
          anchor: origin.start
        }
      }
    ];

    for (const command of commands) {
      expect(session.dispatch(command)).toMatchObject({ status: "committed", changed: true });
      const before = session.getSnapshot();
      expect(session.dispatch(command)).toEqual({
        status: "committed",
        revision: before.revision,
        changed: false
      });
      expect(session.getSnapshot()).toBe(before);
    }
  });

  it("returns an envelope-only revision conflict without mutating any state", () => {
    const diagnostics: WorkbookDiagnosticEvent[] = [];
    const session = createWorkbookSession({
      workbook: createBlankWorkbook(),
      onDiagnostic: (event) => diagnostics.push(event)
    });
    const initial = session.getSnapshot();
    const envelope: CommandEnvelope<WorkbookCommand> = {
      id: "conflicting-command",
      expectedRevision: "99",
      intent: {
        type: "cell.set",
        sheetId: initial.workbook.activeSheetId,
        address: "A1",
        input: "secret-value"
      }
    };

    const result = session.dispatch(envelope);

    expect(result).toEqual({ status: "conflict", revision: "0", current: initial });
    expect(session.getSnapshot()).toBe(initial);
    expect(diagnostics).toHaveLength(1);
    expect(diagnostics[0]).toMatchObject({
      commandId: "conflicting-command",
      metadata: { commandType: "cell.set", outcome: "conflict", changed: false }
    });
    expect(JSON.stringify(diagnostics[0])).not.toContain("secret-value");
  });

  it("rejects invalid worksheet structure commands atomically with actionable issues", () => {
    const diagnostics: WorkbookDiagnosticEvent[] = [];
    let createdIds = 0;
    const session = createWorkbookSession({
      workbook: createBlankWorkbook(),
      createId(kind) {
        createdIds += 1;
        return `${kind}-${createdIds}`;
      },
      onDiagnostic: (event) => diagnostics.push(event)
    });
    const initial = session.getSnapshot();
    const sheetId = initial.workbook.activeSheetId;
    let notifications = 0;
    session.subscribe(() => notifications += 1);

    const result = session.dispatch({ type: "rows.insert", sheetId, index: -1, count: 1 });

    expect(result).toEqual({
      status: "rejected",
      reason: "validation",
      issues: [{
        code: "SHEET_STRUCTURE_INDEX_INVALID",
        message: "Structure index must be a non-negative integer"
      }]
    });
    expect(session.getSnapshot()).toBe(initial);
    expect(notifications).toBe(0);
    expect(createdIds).toBe(0);
    expect(diagnostics).toHaveLength(1);
    expect(diagnostics[0].metadata).toMatchObject({
      commandType: "rows.insert",
      outcome: "rejected",
      changed: false
    });
    expect(diagnostics[0].metadata).not.toHaveProperty("errorType");
  });

  it("undoes and redoes a table-aware insertion without regenerating ids", () => {
    const createId = vi.fn()
      .mockReturnValueOnce("table-column-new-1")
      .mockReturnValueOnce("table-column-new-2");
    const initial = structuredWorkbook();
    const initialBytes = JSON.stringify(initial);
    const session = createWorkbookSession({ workbook: initial, createId });

    const result = session.dispatch({
      type: "transaction",
      commands: [
        {
          type: "columns.insert",
          sheetId: initial.activeSheetId,
          index: 2,
          count: 2,
          expandTableIds: ["table-sales"]
        },
        {
          type: "selection.set",
          selection: { start: { row: 0, column: 2 }, end: { row: 99, column: 3 } }
        }
      ]
    });

    expect(result.status).toBe("committed");
    const inserted = session.getSnapshot().workbook;
    const insertedBytes = JSON.stringify(inserted);
    expect(createId).toHaveBeenCalledTimes(2);

    expect(session.dispatch({ type: "history.undo" })).toMatchObject({ status: "committed", changed: true });
    expect(session.getSnapshot().workbook).toEqual(initial);
    expect(JSON.stringify(session.getSnapshot().workbook)).toBe(initialBytes);
    expect(session.getSnapshot()).toMatchObject({ canUndo: false, canRedo: true });
    expect(createId).toHaveBeenCalledTimes(2);
    expect(session.dispatch({ type: "history.redo" })).toMatchObject({ status: "committed", changed: true });
    expect(session.getSnapshot().workbook).toEqual(inserted);
    expect(JSON.stringify(session.getSnapshot().workbook)).toBe(insertedBytes);
    expect(session.getSnapshot()).toMatchObject({ canUndo: true, canRedo: false });
    expect(createId).toHaveBeenCalledTimes(2);
  });

  it("restores the selection checkpoint together with workbook undo and redo", () => {
    const initial = createBlankWorkbook();
    const initialSelection = {
      start: { row: 0, column: 25 },
      end: { row: 0, column: 25 }
    } as const;
    const appendedSelection = {
      start: { row: 0, column: 26 },
      end: { row: 0, column: 26 }
    } as const;
    const session = createWorkbookSession({ workbook: initial });
    session.dispatch({ type: "selection.set", selection: initialSelection });

    expect(session.dispatch({
      type: "transaction",
      commands: [
        {
          type: "columns.insert",
          sheetId: initial.activeSheetId,
          index: 26,
          count: 1
        },
        { type: "selection.set", selection: appendedSelection }
      ]
    })).toMatchObject({ status: "committed", changed: true });
    expect(session.getSnapshot()).toMatchObject({ selection: appendedSelection });
    expect(session.getSnapshot().workbook.sheets[0].columnCount).toBe(27);

    session.dispatch({ type: "history.undo" });
    expect(session.getSnapshot()).toMatchObject({ selection: initialSelection });
    expect(session.getSnapshot().workbook.sheets[0].columnCount).toBe(26);

    session.dispatch({ type: "history.redo" });
    expect(session.getSnapshot()).toMatchObject({ selection: appendedSelection });
    expect(session.getSnapshot().workbook.sheets[0].columnCount).toBe(27);
  });

  it("keeps a post-undo transaction selection as the next history checkpoint", () => {
    const workbook = createBlankWorkbook();
    const sheetId = workbook.activeSheetId;
    const firstCheckpoint = {
      start: { row: 1, column: 1 },
      end: { row: 1, column: 1 }
    } as const;
    const replacementCheckpoint = {
      start: { row: 2, column: 2 },
      end: { row: 2, column: 2 }
    } as const;
    const session = createWorkbookSession({ workbook });
    expect(session.dispatch({
      type: "transaction",
      commands: [
        { type: "cell.set", sheetId, address: "A1", input: "first" },
        { type: "selection.set", selection: firstCheckpoint }
      ]
    })).toMatchObject({ status: "committed", changed: true });

    expect(session.dispatch({
      type: "transaction",
      commands: [
        { type: "history.undo" },
        { type: "selection.set", selection: replacementCheckpoint }
      ]
    })).toMatchObject({ status: "committed", changed: true });
    expect(session.getSnapshot()).toMatchObject({ selection: replacementCheckpoint });

    expect(session.dispatch({
      type: "cell.set",
      sheetId,
      address: "A2",
      input: "next"
    })).toMatchObject({ status: "committed", changed: true });
    expect(session.dispatch({ type: "history.undo" }))
      .toMatchObject({ status: "committed", changed: true });
    expect(session.getSnapshot()).toMatchObject({ selection: replacementCheckpoint });
    expect(getCellContent(session.getSnapshot().workbook, sheetId, "A2")).toBeNull();
  });

  it("preflights deterministic references to ids generated earlier in the transaction", () => {
    let workbook = createBlankWorkbook();
    const sheetId = workbook.activeSheetId;
    workbook = setCellContent(workbook, sheetId, "A1", "Name");
    workbook = setCellContent(workbook, sheetId, "A2", "Ada");
    const counts: Record<IdKind, number> = {
      table: 0,
      "table-column": 0,
      "table-row": 0
    };
    const createId = vi.fn((kind: IdKind) => `${kind}-${++counts[kind]}`);
    const session = createWorkbookSession({ workbook, createId });
    const style = { theme: "TableStyleLight2", showRowStripes: false } as const;

    expect(session.dispatch({
      type: "transaction",
      commands: [
        {
          type: "table.create",
          sheetId,
          range: { start: { row: 0, column: 0 }, end: { row: 1, column: 0 } },
          name: "People",
          headerRow: true,
          totalsRow: false
        },
        {
          type: "table.setKeyColumn",
          tableId: "table-1",
          columnId: "table-column-1"
        },
        {
          type: "table.editCells",
          tableId: "table-1",
          edits: [{ rowId: "table-row-1", columnId: "table-column-1", rawText: "Grace" }]
        },
        { type: "table.setStyle", tableId: "table-1", style }
      ]
    })).toMatchObject({ status: "committed", changed: true });

    expect(createId).toHaveBeenCalledTimes(3);
    expect(session.getSnapshot().workbook.tables).toMatchObject([
      {
        id: "table-1",
        columns: [{ id: "table-column-1" }],
        rowIds: ["table-row-1"],
        keyColumnId: "table-column-1",
        style
      }
    ]);
    expect(getCellContent(session.getSnapshot().workbook, sheetId, "A2")).toBe("Grace");
  });

  it("rejects an invalid deterministic dependent command before consuming host ids", () => {
    let workbook = createBlankWorkbook();
    const sheetId = workbook.activeSheetId;
    workbook = setCellContent(workbook, sheetId, "A1", "Name");
    workbook = setCellContent(workbook, sheetId, "A2", "Ada");
    const counts: Record<IdKind, number> = {
      table: 0,
      "table-column": 0,
      "table-row": 0
    };
    const createId = vi.fn((kind: IdKind) => `${kind}-${++counts[kind]}`);
    const session = createWorkbookSession({ workbook, createId });
    const before = session.getSnapshot();

    expect(session.dispatch({
      type: "transaction",
      commands: [
        {
          type: "table.create",
          sheetId,
          range: { start: { row: 0, column: 0 }, end: { row: 1, column: 0 } },
          name: "People",
          headerRow: true,
          totalsRow: false
        },
        { type: "table.setStyle", tableId: "table-1", style: { theme: "" } }
      ]
    })).toEqual({
      status: "rejected",
      reason: "validation",
      issues: [{ code: "TABLE_RANGE_BLOCKED", message: "Table style is invalid" }]
    });
    expect(createId).not.toHaveBeenCalled();
    expect(session.getSnapshot()).toBe(before);
  });

  it("semantic-preflights a whole transaction before consuming host ids", () => {
    const createId = vi.fn(() => "table-column-host");
    const session = createWorkbookSession({ workbook: structuredWorkbook(), createId });
    const before = session.getSnapshot();

    const result = session.dispatch({
      type: "transaction",
      commands: [
        {
          type: "columns.insert",
          sheetId: before.workbook.activeSheetId,
          index: 2,
          count: 1,
          expandTableIds: ["table-sales"]
        },
        {
          type: "rows.delete",
          sheetId: before.workbook.activeSheetId,
          index: -1,
          count: 1
        }
      ]
    });

    expect(result).toMatchObject({
      status: "rejected",
      reason: "validation",
      issues: [{ code: "SHEET_STRUCTURE_INDEX_INVALID" }]
    });
    expect(createId).not.toHaveBeenCalled();
    expect(session.getSnapshot()).toBe(before);
  });

  it("rejects invalid generated table ids without publishing any candidate state", () => {
    let workbook = createBlankWorkbook();
    const sheetId = workbook.activeSheetId;
    workbook = setCellContent(workbook, sheetId, "A1", "Name");
    workbook = setCellContent(workbook, sheetId, "B1", "Amount");
    const createId = vi.fn(() => " ");
    const session = createWorkbookSession({ workbook, createId });
    const before = session.getSnapshot();

    expect(session.dispatch({
      type: "table.create",
      sheetId,
      range: { start: { row: 0, column: 0 }, end: { row: 1, column: 1 } },
      name: "GeneratedIds",
      headerRow: true,
      totalsRow: false
    })).toMatchObject({ status: "rejected", reason: "validation" });
    expect(session.getSnapshot()).toBe(before);
  });

  it("rejects generated ids that collide with another structured table", () => {
    let workbook = structuredWorkbook();
    const sheetId = workbook.activeSheetId;
    workbook = setCellContent(workbook, sheetId, "D1", "Category");
    workbook = setCellContent(workbook, sheetId, "E1", "Value");
    const generatedIds = [
      "sales-region",
      "second-value",
      "table-second",
      "second-row"
    ];
    const createId = vi.fn(() => generatedIds.shift() ?? "unexpected-id");
    const session = createWorkbookSession({ workbook, createId });
    const before = session.getSnapshot();

    expect(session.dispatch({
      type: "table.create",
      sheetId,
      range: { start: { row: 0, column: 3 }, end: { row: 1, column: 4 } },
      name: "SecondTable",
      headerRow: true,
      totalsRow: false
    })).toEqual({
      status: "rejected",
      reason: "validation",
      issues: [{
        code: "TABLE_GENERATED_ID_INVALID",
        message: "Generated structured-table IDs must be nonblank and unique"
      }]
    });
    expect(createId).toHaveBeenCalledTimes(4);
    expect(session.getSnapshot()).toBe(before);
    expect(session.getSnapshot()).toMatchObject({
      revision: before.revision,
      selection: before.selection,
      canUndo: before.canUndo,
      canRedo: before.canRedo
    });
  });

  it("routes duplicated structured-table ids through the session allocator", () => {
    const initial = structuredWorkbook();
    let sequence = 0;
    const createId = vi.fn((kind: IdKind) => `${kind}-copy-${++sequence}`);
    const session = createWorkbookSession({ workbook: initial, createId });

    expect(session.dispatch({
      type: "sheet.duplicate",
      sheetId: initial.activeSheetId
    })).toMatchObject({ status: "committed", changed: true });

    expect(createId).toHaveBeenCalledTimes(5);
    expect(session.getSnapshot().workbook.tables[1]).toMatchObject({
      id: "table-copy-3",
      sheetId: session.getSnapshot().workbook.activeSheetId,
      columns: [
        { id: "table-column-copy-1" },
        { id: "table-column-copy-2" }
      ],
      rowIds: ["table-row-copy-4", "table-row-copy-5"]
    });
  });

  it.each([
    {
      label: "blank",
      ids: [" ", "copy-column-2", "copy-table", "copy-row-1", "copy-row-2"]
    },
    {
      label: "duplicate",
      ids: ["copy-shared", "copy-shared", "copy-table", "copy-row-1", "copy-row-2"]
    },
    {
      label: "global collision",
      ids: ["sales-region", "copy-column-2", "copy-table", "copy-row-1", "copy-row-2"]
    }
  ])("rejects $label ids from structured-table sheet duplication", ({ ids }) => {
    const generatedIds = [...ids];
    const createId = vi.fn(() => generatedIds.shift() ?? "unexpected-id");
    const session = createWorkbookSession({ workbook: structuredWorkbook(), createId });
    const before = session.getSnapshot();

    expect(session.dispatch({
      type: "sheet.duplicate",
      sheetId: before.workbook.activeSheetId
    })).toEqual({
      status: "rejected",
      reason: "validation",
      issues: [{
        code: "TABLE_GENERATED_ID_INVALID",
        message: "Generated structured-table IDs must be nonblank and unique"
      }]
    });
    expect(createId).toHaveBeenCalledTimes(5);
    expect(session.getSnapshot()).toBe(before);
    expect(session.getSnapshot()).toMatchObject({
      revision: before.revision,
      selection: before.selection,
      canUndo: before.canUndo,
      canRedo: before.canRedo
    });
  });

  it("keeps selection, revision, history, and ids unchanged when structure rejects", () => {
    const createId = vi.fn(() => "unused-table-column");
    const session = createWorkbookSession({ workbook: structuredWorkbook(), createId });
    const before = session.getSnapshot();

    const result = session.dispatch({
      type: "columns.delete",
      sheetId: before.workbook.activeSheetId,
      index: 0,
      count: 2
    });

    expect(result).toMatchObject({ status: "rejected", reason: "validation" });
    expect(session.getSnapshot()).toMatchObject({
      revision: before.revision,
      selection: before.selection,
      canUndo: before.canUndo,
      canRedo: before.canRedo
    });
    expect(session.getSnapshot().workbook).toBe(before.workbook);
    expect(session.getSnapshot()).toBe(before);
    expect(createId).not.toHaveBeenCalled();
  });

  it("rejects column insertion preflight without allocating ids or changing the snapshot", () => {
    const createId = vi.fn(() => "unused-table-column");
    const session = createWorkbookSession({ workbook: structuredWorkbook(), createId });
    const before = session.getSnapshot();

    const result = session.dispatch({
      type: "columns.insert",
      sheetId: before.workbook.activeSheetId,
      index: 2,
      count: 1,
      expandTableIds: ["table-sales", "table-sales"]
    });

    expect(result).toEqual({
      status: "rejected",
      reason: "validation",
      issues: [{
        code: "TABLE_EXPANSION_CONTEXT_INVALID",
        message: "Column insertion expansion context contains duplicate table ids"
      }]
    });
    expect(session.getSnapshot()).toBe(before);
    expect(createId).not.toHaveBeenCalled();
  });

  it("preserves safe exception detail in diagnostics while keeping rejections generic", () => {
    const diagnostics: WorkbookDiagnosticEvent[] = [];
    const session = createWorkbookSession({
      workbook: createBlankWorkbook(),
      onDiagnostic: (event) => diagnostics.push(event)
    });
    const sheetId = session.getSnapshot().workbook.activeSheetId;

    const result = session.dispatch({
      type: "rows.resize",
      sheetId,
      rows: [-1],
      height: 32
    });

    expect(result).toEqual({
      status: "rejected",
      reason: "unsupported",
      issues: [{ code: "command.invalid", message: "Command could not be applied" }]
    });
    expect(diagnostics).toHaveLength(1);
    expect(diagnostics[0].metadata).toMatchObject({
      outcome: "rejected",
      errorType: "Error",
      errorFingerprint: expect.stringMatching(/^[0-9a-f]{8}$/),
      errorStackFrame: expect.stringContaining("commands.ts")
    });
    expect(JSON.stringify(result)).not.toContain("Index must be");
  });

  it("keeps reducer rejections generic when exception metadata cannot be inspected", () => {
    let workbook = createBlankWorkbook();
    const sheetId = workbook.activeSheetId;
    workbook = setCellContent(workbook, sheetId, "A1", "Name");
    const diagnostics: WorkbookDiagnosticEvent[] = [];
    const failure = new Error("private failure detail");
    Object.defineProperty(failure, "stack", { value: { unavailable: true } });
    const session = createWorkbookSession({
      workbook,
      createId() {
        throw failure;
      },
      onDiagnostic: (event) => diagnostics.push(event)
    });

    const result = session.dispatch({
      type: "table.create",
      sheetId,
      range: origin,
      name: "People",
      headerRow: true,
      totalsRow: false
    });

    expect(result).toEqual({
      status: "rejected",
      reason: "unsupported",
      issues: [{ code: "command.invalid", message: "Command could not be applied" }]
    });
    expect(diagnostics).toHaveLength(1);
    expect(diagnostics[0].metadata).toMatchObject({
      errorType: "Error",
      errorFingerprint: expect.stringMatching(/^[0-9a-f]{8}$/)
    });
    expect(diagnostics[0].metadata).not.toHaveProperty("errorStackFrame");
  });

  it("does not copy arbitrary exception stack text into diagnostics", () => {
    let workbook = createBlankWorkbook();
    const sheetId = workbook.activeSheetId;
    workbook = setCellContent(workbook, sheetId, "A1", "Name");
    const diagnostics: WorkbookDiagnosticEvent[] = [];
    const failure = new Error("private failure detail");
    failure.stack = "Error: private failure detail\n at secret-stack-value";
    const session = createWorkbookSession({
      workbook,
      createId() {
        throw failure;
      },
      onDiagnostic: (event) => diagnostics.push(event)
    });

    const result = session.dispatch({
      type: "table.create",
      sheetId,
      range: origin,
      name: "People",
      headerRow: true,
      totalsRow: false
    });

    expect(result.status).toBe("rejected");
    expect(diagnostics).toHaveLength(1);
    expect(JSON.stringify(diagnostics[0])).not.toContain("secret-stack-value");
    expect(diagnostics[0].metadata).not.toHaveProperty("errorStackFrame");
  });

  it("serializes hostile non-Error causes without replacing the rejection", () => {
    let workbook = createBlankWorkbook();
    const sheetId = workbook.activeSheetId;
    workbook = setCellContent(workbook, sheetId, "A1", "Name");
    const diagnostics: WorkbookDiagnosticEvent[] = [];
    const hostileCause = new Proxy({}, {
      getPrototypeOf() {
        throw new Error("prototype unavailable");
      },
      get() {
        throw new Error("property unavailable");
      }
    });
    const session = createWorkbookSession({
      workbook,
      createId() {
        throw hostileCause;
      },
      onDiagnostic: (event) => diagnostics.push(event)
    });

    const result = session.dispatch({
      type: "table.create",
      sheetId,
      range: origin,
      name: "People",
      headerRow: true,
      totalsRow: false
    });

    expect(result.status).toBe("rejected");
    expect(diagnostics).toHaveLength(1);
    expect(diagnostics[0].metadata).toMatchObject({
      errorType: "object",
      errorFingerprint: expect.stringMatching(/^[0-9a-f]{8}$/)
    });
  });

  it("updates and recalculates the formula engine before publishing the snapshot", () => {
    let workbook = createBlankWorkbook();
    const sheetId = workbook.activeSheetId;
    workbook = setCellContent(workbook, sheetId, "A1", 2);
    workbook = setCellContent(workbook, sheetId, "B1", "=A1*3");
    const session = createWorkbookSession({ workbook });
    const evaluationsAtPublication: unknown[] = [];
    session.subscribe(() => evaluationsAtPublication.push(session.getCellEvaluation(sheetId, "B1")));

    session.dispatch({ type: "cell.set", sheetId, address: "A1", input: "4" });

    expect(evaluationsAtPublication).toEqual([12]);
    expect(session.getCellEvaluation(sheetId, "B1")).toBe(12);
  });

  it("does not publish or retain a candidate when the formula projection fails", () => {
    const diagnostics: WorkbookDiagnosticEvent[] = [];
    const projected: WorkbookModel[] = [];
    const session = createWorkbookSession({
      workbook: createBlankWorkbook(),
      formulaEngineFactory(workbook) {
        let current = workbook;
        return {
          getDisplayValue(sheetId, address) {
            const value = getCellContent(current, sheetId, address);
            return value === null ? "" : String(value);
          },
          getComputedValue(sheetId, address) {
            return getCellContent(current, sheetId, address);
          },
          getRawContent(sheetId, address) {
            return getCellContent(current, sheetId, address);
          },
          update() {
            throw new Error("projection failed");
          },
          rebuild(nextWorkbook) {
            current = nextWorkbook;
            projected.push(nextWorkbook);
          },
          destroy() {}
        };
      },
      onDiagnostic: (event) => diagnostics.push(event)
    });
    const initial = session.getSnapshot();
    const sheetId = initial.workbook.activeSheetId;
    let notifications = 0;
    session.subscribe(() => notifications += 1);

    expect(session.dispatch({ type: "cell.set", sheetId, address: "A1", input: "5" })).toEqual({
      status: "rejected",
      reason: "unsupported",
      issues: [{ code: "command.invalid", message: "Command could not be applied" }]
    });
    expect(session.getSnapshot()).toBe(initial);
    expect(session.getSnapshot()).toMatchObject({ revision: "0", canUndo: false, canRedo: false });
    expect(notifications).toBe(0);
    expect(projected).toEqual([initial.workbook]);
    expect(diagnostics).toHaveLength(1);
    expect(diagnostics[0]).toMatchObject({ category: "command", metadata: { outcome: "rejected" } });
  });

  it("recreates the live projection from the committed workbook when update and rebuild both fail", () => {
    let workbook = createBlankWorkbook();
    const sheetId = workbook.activeSheetId;
    workbook = setCellContent(workbook, sheetId, "A1", 1);
    let factoryCalls = 0;
    let poisonedEngineDestroyed = 0;
    const session = createWorkbookSession({
      workbook,
      formulaEngineFactory(initialWorkbook) {
        factoryCalls += 1;
        if (factoryCalls === 1) {
          return createPoisoningFormulaEngine(initialWorkbook, () => poisonedEngineDestroyed += 1);
        }
        return createRawFormulaEngine(initialWorkbook);
      }
    });
    const initial = session.getSnapshot();

    expect(session.dispatch({ type: "cell.set", sheetId, address: "A1", input: "99" })).toEqual({
      status: "rejected",
      reason: "unsupported",
      issues: [{ code: "command.invalid", message: "Command could not be applied" }]
    });
    expect(session.getSnapshot()).toBe(initial);
    expect(session.getCellEvaluation(sheetId, "A1")).toBe(1);
    expect(factoryCalls).toBe(2);
    expect(poisonedEngineDestroyed).toBe(1);

    expect(session.dispatch({ type: "cell.set", sheetId, address: "A1", input: "2" })).toEqual({
      status: "committed",
      revision: "1",
      changed: true
    });
    expect(session.getCellEvaluation(sheetId, "A1")).toBe(2);
  });

  it("uses a stable error and rejects projection-dependent mutations while recovery is unavailable", () => {
    let workbook = createBlankWorkbook();
    const sheetId = workbook.activeSheetId;
    workbook = setCellContent(workbook, sheetId, "A1", 1);
    const diagnostics: WorkbookDiagnosticEvent[] = [];
    let factoryCalls = 0;
    const session = createWorkbookSession({
      workbook,
      formulaEngineFactory(initialWorkbook) {
        factoryCalls += 1;
        if (factoryCalls === 1) {
          return createPoisoningFormulaEngine(initialWorkbook, () => {
            throw new Error("destroy failed");
          });
        }
        throw new Error("factory unavailable");
      },
      onDiagnostic: (event) => diagnostics.push(event)
    });
    const initial = session.getSnapshot();

    expect(session.dispatch({ type: "cell.set", sheetId, address: "A1", input: "secret-candidate" }))
      .toMatchObject({ status: "rejected", reason: "unsupported" });
    expect(session.getSnapshot()).toBe(initial);
    const firstUnavailable = session.getCellEvaluation(sheetId, "A1");
    expect(firstUnavailable).toEqual({ kind: "error", code: "#PROJECTION!" });
    expect(session.getCellEvaluation(sheetId, "A1")).toBe(firstUnavailable);

    expect(session.dispatch({ type: "cell.set", sheetId, address: "A1", input: "second-secret" })).toEqual({
      status: "rejected",
      reason: "unsupported",
      issues: [{ code: "projection.unavailable", message: "Formula projection is unavailable" }]
    });
    expect(session.getSnapshot()).toBe(initial);
    expect(factoryCalls).toBe(3);
    expect(JSON.stringify(diagnostics)).not.toContain("secret-candidate");
    expect(JSON.stringify(diagnostics)).not.toContain("second-secret");
    expect(() => session.destroy()).not.toThrow();
    expect(() => session.destroy()).not.toThrow();
  });

  it("retries an unavailable projection before the next dependent mutation", () => {
    let workbook = createBlankWorkbook();
    const sheetId = workbook.activeSheetId;
    workbook = setCellContent(workbook, sheetId, "A1", 1);
    let factoryCalls = 0;
    const session = createWorkbookSession({
      workbook,
      formulaEngineFactory(initialWorkbook) {
        factoryCalls += 1;
        if (factoryCalls === 1) {
          return createPoisoningFormulaEngine(initialWorkbook);
        }
        if (factoryCalls === 2) {
          throw new Error("transient factory failure");
        }
        return createRawFormulaEngine(initialWorkbook);
      }
    });

    expect(session.dispatch({ type: "cell.set", sheetId, address: "A1", input: "9" }))
      .toMatchObject({ status: "rejected" });
    expect(session.getCellEvaluation(sheetId, "A1")).toEqual({ kind: "error", code: "#PROJECTION!" });
    expect(session.dispatch({ type: "cell.set", sheetId, address: "A1", input: "3" })).toEqual({
      status: "committed",
      revision: "1",
      changed: true
    });
    expect(session.getCellEvaluation(sheetId, "A1")).toBe(3);
    expect(factoryCalls).toBe(3);
  });

  it("maps only exact zero-or-one freeze counts and rejects all other counts atomically", () => {
    const session = createWorkbookSession({ workbook: createBlankWorkbook() });
    const sheetId = session.getSnapshot().workbook.activeSheetId;

    expect(session.dispatch({ type: "sheet.freeze.set", sheetId, rows: 1, columns: 0 }))
      .toEqual({ status: "committed", revision: "1", changed: true });
    expect(session.getSnapshot().workbook.sheets[0]).toMatchObject({
      freezeTopRow: true,
      freezeFirstColumn: false
    });
    expect(session.dispatch({ type: "sheet.freeze.set", sheetId, rows: 0, columns: 1 }))
      .toEqual({ status: "committed", revision: "2", changed: true });
    expect(session.getSnapshot().workbook.sheets[0]).toMatchObject({
      freezeTopRow: false,
      freezeFirstColumn: true
    });

    for (const [rows, columns] of [
      [-1, 0],
      [0, -1],
      [0.5, 0],
      [0, 0.5],
      [2, 0],
      [0, 2],
      [5, 3]
    ] as const) {
      const before = session.getSnapshot();
      expect(session.dispatch({ type: "sheet.freeze.set", sheetId, rows, columns })).toEqual({
        status: "rejected",
        reason: "unsupported",
        issues: [{ code: "command.invalid", message: "Command could not be applied" }]
      });
      expect(session.getSnapshot()).toBe(before);
    }
  });

  it("updates selection without workbook history and persistence without workbook revision", () => {
    const session = createWorkbookSession({ workbook: createBlankWorkbook() });
    let notifications = 0;
    session.subscribe(() => notifications += 1);

    expect(session.dispatch({ type: "selection.set", selection: block })).toEqual({
      status: "committed",
      revision: "1",
      changed: true
    });
    expect(session.getSnapshot()).toMatchObject({ selection: block, canUndo: false, canRedo: false });

    expect(session.dispatch({ type: "persistence.status", status: "failed", message: "quota" })).toEqual({
      status: "committed",
      revision: "1",
      changed: true
    });
    expect(session.getSnapshot()).toMatchObject({
      revision: "1",
      canUndo: false,
      canRedo: false,
      persistence: { status: "failed", message: "quota" }
    });
    expect(notifications).toBe(2);
    expect(session.dispatch({ type: "persistence.status", status: "failed", message: "quota" })).toEqual({
      status: "committed",
      revision: "1",
      changed: false
    });
    expect(notifications).toBe(2);
  });

  it("publishes sheet activation without adding workbook history", () => {
    const workbook = addSheet(createBlankWorkbook(), "Second");
    const session = createWorkbookSession({ workbook: setActiveSheet(workbook, "sheet-1") });

    expect(session.dispatch({ type: "sheet.activate", sheetId: "sheet-2" })).toMatchObject({
      status: "committed",
      changed: true,
      revision: "1"
    });
    expect(session.getSnapshot()).toMatchObject({ canUndo: false, canRedo: false });
    expect(session.getSnapshot().workbook.activeSheetId).toBe("sheet-2");

    session.dispatch({ type: "cell.set", sheetId: "sheet-2", address: "A1", input: "durable" });
    session.dispatch({ type: "sheet.activate", sheetId: "sheet-1" });
    expect(session.dispatch({ type: "history.undo" })).toMatchObject({ status: "committed", changed: true });
    expect(getCellContent(session.getSnapshot().workbook, "sheet-2", "A1")).toBeNull();
    session.destroy();
  });

  it("enforces read-only permissions without publishing the rejected candidate", () => {
    const session = createWorkbookSession({ workbook: createBlankWorkbook() });
    const sheetId = session.getSnapshot().workbook.activeSheetId;
    session.dispatch({ type: "range.readOnly.set", sheetId, range: origin, readOnly: true });
    const before = session.getSnapshot();

    expect(session.dispatch({ type: "cell.set", sheetId, address: "A1", input: "42" })).toEqual({
      status: "rejected",
      reason: "permission",
      issues: [{
        code: "permission.readOnly",
        message: "Cell is read-only",
        sheetId,
        address: "A1"
      }]
    });
    expect(session.getSnapshot()).toBe(before);
  });

  it("rejects malformed commands without throwing or publishing partial state", () => {
    const session = createWorkbookSession({ workbook: createBlankWorkbook() });
    const initial = session.getSnapshot();

    expect(session.dispatch({
      type: "cell.set",
      sheetId: initial.workbook.activeSheetId,
      address: "not-a-cell",
      input: "value"
    })).toEqual({
      status: "rejected",
      reason: "unsupported",
      issues: [{ code: "command.invalid", message: "Command could not be applied" }]
    });
    expect(session.getSnapshot()).toBe(initial);
  });

  it("uses injected IDs deterministically and default IDs uniquely across sessions", () => {
    const injected: WorkbookDiagnosticEvent[] = [];
    const deterministic = createWorkbookSession({
      workbook: createBlankWorkbook(),
      createCommandId: () => "deterministic-id",
      onDiagnostic: (event) => injected.push(event)
    });
    deterministic.dispatch({ type: "history.undo" });
    expect(injected[0].commandId).toBe("deterministic-id");

    const firstEvents: WorkbookDiagnosticEvent[] = [];
    const secondEvents: WorkbookDiagnosticEvent[] = [];
    const first = createWorkbookSession({
      workbook: createBlankWorkbook(),
      onDiagnostic: (event) => firstEvents.push(event)
    });
    const second = createWorkbookSession({
      workbook: createBlankWorkbook(),
      onDiagnostic: (event) => secondEvents.push(event)
    });
    first.dispatch({ type: "history.undo" });
    second.dispatch({ type: "history.undo" });

    expect(firstEvents[0].commandId).toBeTruthy();
    expect(secondEvents[0].commandId).toBeTruthy();
    expect(firstEvents[0].commandId).not.toBe(secondEvents[0].commandId);
  });

  it("emits value-free diagnostics for transactions and persistence failures", () => {
    const events: WorkbookDiagnosticEvent[] = [];
    const session = createWorkbookSession({
      workbook: createBlankWorkbook(),
      createCommandId: () => "safe-command-id",
      now: sequenceNow(10, 16),
      onDiagnostic: (event) => events.push(event)
    });
    const sheetId = session.getSnapshot().workbook.activeSheetId;
    const secret = "ultra-secret-value";

    session.dispatch({
      type: "transaction",
      commands: [
        { type: "cell.set", sheetId, address: "A1", input: secret },
        { type: "cell.comment.set", sheetId, address: "A1", comment: `${secret}-comment` },
        { type: "cell.hyperlink.set", sheetId, address: "A1", hyperlink: `https://${secret}.test` },
        {
          type: "range.validation.set",
          sheetId,
          range: origin,
          rule: { type: "list", values: [secret] }
        },
        { type: "persistence.status", status: "failed", message: `${secret}-storage` }
      ]
    });

    expect(events).toHaveLength(1);
    expect(events[0]).toMatchObject({
      category: "persistence",
      commandId: "safe-command-id",
      durationMs: 6,
      metadata: { commandType: "transaction", outcome: "committed", changed: true }
    });
    expect(JSON.stringify(events[0])).not.toContain(secret);
    expect(Object.values(events[0].metadata).every((value) =>
      value === null || ["string", "number", "boolean"].includes(typeof value)
    )).toBe(true);
  });

  it("isolates host callback failures and diagnostic subscribe/unsubscribe", () => {
    const received: WorkbookDiagnosticEvent[] = [];
    const session = createWorkbookSession({
      workbook: createBlankWorkbook(),
      onDiagnostic: () => { throw new Error("host diagnostic failed"); }
    });
    const sheetId = session.getSnapshot().workbook.activeSheetId;
    let regularNotifications = 0;
    session.subscribe(() => { throw new Error("host subscriber failed"); });
    session.subscribe(() => regularNotifications += 1);
    session.subscribeDiagnostics(() => { throw new Error("diagnostic subscriber failed"); });
    const unsubscribe = session.subscribeDiagnostics((event) => received.push(event));

    expect(() => session.dispatch({ type: "cell.set", sheetId, address: "A1", input: "1" })).not.toThrow();
    expect(regularNotifications).toBe(1);
    expect(received).toHaveLength(1);
    unsubscribe();
    session.dispatch({ type: "cell.set", sheetId, address: "A2", input: "2" });
    expect(received).toHaveLength(1);
  });

  it("replaces workbooks with preserved or reset bounded history", () => {
    const session = createWorkbookSession({
      workbook: createBlankWorkbook(),
      history: { maxEntries: 1, maxWeight: 10_000 }
    });
    const initial = session.getSnapshot().workbook;
    const sheetId = initial.activeSheetId;
    const imported = setCellContent(initial, sheetId, "A1", 7);

    expect(session.replaceWorkbook(imported, { history: "preserve", origin: "import" })).toMatchObject({
      status: "committed",
      changed: true
    });
    expect(session.getSnapshot().canUndo).toBe(true);
    session.dispatch({ type: "history.undo" });
    expect(session.getSnapshot().workbook).toBe(initial);

    expect(session.dispatch({ type: "workbook.replace", workbook: imported, history: "reset" })).toMatchObject({
      status: "committed",
      changed: true
    });
    expect(session.getSnapshot()).toMatchObject({ canUndo: false, canRedo: false });
    expect(session.dispatch({ type: "history.undo" })).toMatchObject({ changed: false });
  });

  it("destroys the engine and clears all subscriptions", () => {
    const engines: TrackingEngine[] = [];
    let notifications = 0;
    let diagnosticNotifications = 0;
    let optionDiagnostics = 0;
    const session = createWorkbookSession({
      workbook: createBlankWorkbook(),
      formulaEngineFactory: createTrackingFormulaEngineFactory(engines),
      onDiagnostic: () => optionDiagnostics += 1
    });
    session.subscribe(() => notifications += 1);
    session.subscribeDiagnostics(() => diagnosticNotifications += 1);

    session.destroy();
    expect(engines[0].destroyed).toBe(1);
    expect(session.dispatch({ type: "history.undo" })).toEqual({
      status: "rejected",
      reason: "unsupported",
      issues: [{ code: "session.destroyed", message: "Workbook session is destroyed" }]
    });
    expect(notifications).toBe(0);
    expect(diagnosticNotifications).toBe(0);
    expect(optionDiagnostics).toBe(0);
    expect(() => session.destroy()).not.toThrow();
    expect(engines[0].destroyed).toBe(1);
  });
});

type TrackingEngine = {
  projected: WorkbookModel;
  updates: WorkbookModel[];
  destroyed: number;
};

function structuredWorkbook(): WorkbookModel {
  const workbook = createBlankWorkbook();
  const table: StructuredTable = {
    id: "table-sales",
    name: "SalesTable",
    sheetId: workbook.activeSheetId,
    range: { start: { row: 0, column: 0 }, end: { row: 2, column: 1 } },
    headerRow: true,
    totalsRow: false,
    columns: [
      { id: "sales-region", name: "Region", sheetColumn: 0, dataType: "text" },
      { id: "sales-amount", name: "Amount", sheetColumn: 1, dataType: "number" }
    ],
    rowIds: ["sales-row-1", "sales-row-2"],
    keyColumnId: "sales-region",
    sort: [{ columnId: "sales-amount", direction: "desc" }]
  };
  return {
    ...workbook,
    sheets: [{
      ...workbook.sheets[0],
      cells: {
        ...workbook.sheets[0].cells,
        A1: "Region",
        B1: "Amount",
        A2: "West",
        B2: 10,
        A3: "East",
        B3: 20
      }
    }],
    tables: [table]
  };
}

function createTrackingFormulaEngineFactory(records: TrackingEngine[]): (workbook: WorkbookModel) => FormulaEngine {
  return (workbook) => {
    const record: TrackingEngine = { projected: workbook, updates: [], destroyed: 0 };
    records.push(record);
    return {
      getDisplayValue(sheetId, address) {
        const value = getCellContent(record.projected, sheetId, address);
        return value === null ? "" : String(value);
      },
      getComputedValue(sheetId, address) {
        return getCellContent(record.projected, sheetId, address);
      },
      getRawContent(sheetId, address) {
        return getCellContent(record.projected, sheetId, address);
      },
      update(nextWorkbook) {
        record.projected = nextWorkbook;
        record.updates.push(nextWorkbook);
      },
      rebuild(nextWorkbook) {
        record.projected = nextWorkbook;
      },
      destroy() {
        record.destroyed += 1;
      }
    };
  };
}

function sequenceNow(...values: number[]): () => number {
  return () => values.shift() ?? 0;
}

function createRawFormulaEngine(workbook: WorkbookModel): FormulaEngine {
  let projected = workbook;
  return {
    getDisplayValue(sheetId, address) {
      const value = getCellContent(projected, sheetId, address);
      return value === null ? "" : String(value);
    },
    getComputedValue(sheetId, address) {
      return getCellContent(projected, sheetId, address);
    },
    getRawContent(sheetId, address) {
      return getCellContent(projected, sheetId, address);
    },
    update(nextWorkbook) {
      projected = nextWorkbook;
    },
    rebuild(nextWorkbook) {
      projected = nextWorkbook;
    },
    destroy() {}
  };
}

function createPoisoningFormulaEngine(
  workbook: WorkbookModel,
  onDestroy: () => void = () => {}
): FormulaEngine {
  let projected = workbook;
  return {
    getDisplayValue(sheetId, address) {
      const value = getCellContent(projected, sheetId, address);
      return value === null ? "" : String(value);
    },
    getComputedValue(sheetId, address) {
      return getCellContent(projected, sheetId, address);
    },
    getRawContent(sheetId, address) {
      return getCellContent(projected, sheetId, address);
    },
    update(nextWorkbook) {
      projected = nextWorkbook;
      throw new Error("projection update failed after mutation");
    },
    rebuild() {
      throw new Error("projection rebuild failed");
    },
    destroy: onDestroy
  };
}
