import { describe, expect, it } from "vitest";
import type { WorkbookModel } from "../../types";
import type { FormulaEngine } from "../../lib/formulaEngine";
import {
  createBlankWorkbook,
  getCellContent,
  setCellContent
} from "../../lib/workbook";
import type { CommandEnvelope } from "../commands/types";
import type { WorkbookCommand } from "./commands";
import {
  createWorkbookSession,
  type WorkbookDiagnosticEvent
} from "./WorkbookSession";

const origin = { start: { row: 0, column: 0 }, end: { row: 0, column: 0 } } as const;
const block = { start: { row: 0, column: 0 }, end: { row: 1, column: 1 } } as const;

describe("WorkbookSession", () => {
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
        message: "Cell value does not satisfy validation",
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
