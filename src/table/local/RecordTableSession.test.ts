import { describe, expect, it, vi } from "vitest";
import { createColumnHelper } from "../core/columnHelper";
import { createCommandIdFactory } from "../core/commandId";
import type { TableViewState } from "../core/types";
import { createLocalRecordTableSession, type LocalRecordTableSessionOptions } from "./RecordTableSession";
import { createTableMetadataKey } from "./tableMetadata";

type Employee = { id: string; name: string; salary: number; active: boolean };

function createColumns() {
  const helper = createColumnHelper<Employee>();
  return [
    helper.accessor("name", { id: "name", header: "Name", dataType: "text" }),
    helper.accessor("salary", {
      id: "salary",
      header: "Salary",
      dataType: "number",
      validate: ({ parsed }) => parsed < 0
        ? [{ code: "negative-salary", message: "Salary must be non-negative", columnId: "salary" }]
        : []
    }),
    helper.accessor("active", { id: "active", header: "Active", dataType: "boolean" })
  ] as const;
}

function employee(id = "e1", name = "Ada", salary = 100): Employee {
  return { id, name, salary, active: true };
}

function deterministicOptions(
  overrides: Partial<LocalRecordTableSessionOptions<Employee>> = {}
): LocalRecordTableSessionOptions<Employee> {
  return {
    source: { kind: "local", rows: [employee()], getRowId: (row) => row.id },
    columns: createColumns(),
    commandIdFactory: createCommandIdFactory(() => "local-test"),
    ...overrides
  };
}

describe("createLocalRecordTableSession", () => {
  it("parses a typed controlled edit and emits one replayable updater", async () => {
    const original: readonly Employee[] = [employee()];
    const onRowsChange = vi.fn();
    const session = createLocalRecordTableSession({
      source: { kind: "local", rows: original, getRowId: (row) => row.id, onRowsChange },
      columns: createColumns()
    });

    const result = await session.dispatch({
      type: "edit-cells",
      edits: [{ rowId: "e1", columnId: "salary", rawText: "125" }]
    });
    const updater = onRowsChange.mock.calls[0][0] as (rows: readonly Employee[]) => readonly Employee[];

    expect(result).toMatchObject({ status: "committed", changed: true });
    expect(updater(original)).toEqual([{ id: "e1", name: "Ada", salary: 125, active: true }]);
    expect(onRowsChange).toHaveBeenCalledTimes(1);
    expect(session.getSnapshot().getCell("e1", "salary").storedValue).toBe(125);
  });

  it("rejects an invalid batch without partial rows, metadata, history, or publication", async () => {
    const rows: readonly Employee[] = [employee(), { ...employee("e2", "Grace", 200), active: false }];
    const onRowsChange = vi.fn();
    const listener = vi.fn();
    const session = createLocalRecordTableSession({
      source: { kind: "local", rows, getRowId: (row) => row.id, onRowsChange },
      columns: createColumns()
    });
    session.subscribe(listener);

    const result = await session.dispatch({
      type: "edit-cells",
      edits: [
        { rowId: "e1", columnId: "name", rawText: "Changed" },
        { rowId: "e2", columnId: "salary", rawText: "-1" }
      ]
    });

    expect(result).toEqual({
      status: "rejected",
      reason: "validation",
      issues: [{ code: "negative-salary", message: "Salary must be non-negative", columnId: "salary" }]
    });
    expect(onRowsChange).not.toHaveBeenCalled();
    expect(listener).not.toHaveBeenCalled();
    expect(session.getSnapshot().getCell("e1", "name").storedValue).toBe("Ada");
    expect(session.getSnapshot().canUndo).toBe(false);
  });

  it("publishes source-neutral operation state and empty local async state", () => {
    const session = createLocalRecordTableSession(deterministicOptions());
    const snapshot = session.getSnapshot();
    expect(snapshot.operationStates.sort).toEqual({ enabled: true, scopeLabel: "Complete dataset" });
    expect(snapshot.pendingOperations).toEqual([]);
    expect(snapshot.conflicts).toEqual([]);
  });

  it("uses callback presence as the controlled row authority contract", async () => {
    const rows = [employee()];
    const onRowsChange = vi.fn();
    const session = createLocalRecordTableSession(deterministicOptions({
      source: { kind: "local", rows, getRowId: (row) => row.id, onRowsChange }
    }));
    await session.dispatch({ type: "edit-cells", edits: [{ rowId: "e1", columnId: "name", rawText: "Grace" }] });
    expect(rows[0].name).toBe("Ada");
    expect(session.getSnapshot().getCell("e1", "name").storedValue).toBe("Grace");
    expect(onRowsChange).toHaveBeenCalledTimes(1);
  });

  it("replays controlled cell edits by stable id without overwriting concurrent fields", async () => {
    const original = [employee()];
    const onRowsChange = vi.fn();
    const session = createLocalRecordTableSession(deterministicOptions({
      source: { kind: "local", rows: original, getRowId: (row) => row.id, onRowsChange }
    }));
    await session.dispatch({ type: "edit-cells", edits: [{ rowId: "e1", columnId: "name", rawText: "Grace" }] });
    const updater = onRowsChange.mock.calls[0][0] as (rows: readonly Employee[]) => readonly Employee[];
    expect(updater([{ ...employee(), salary: 999 }])).toEqual([{ ...employee("e1", "Grace"), salary: 999 }]);
  });

  it("keeps an uncontrolled edit across rerenders with new rows array identities", async () => {
    const session = createLocalRecordTableSession(deterministicOptions());
    await session.dispatch({ type: "edit-cells", edits: [{ rowId: "e1", columnId: "name", rawText: "Grace" }] });
    session.updateOptions(deterministicOptions({
      source: { kind: "local", rows: [{ ...employee() }], getRowId: (row) => row.id }
    }));
    expect(session.getSnapshot().getCell("e1", "name").storedValue).toBe("Grace");
  });

  it("resets uncontrolled rows and history only when resetKey changes", async () => {
    const session = createLocalRecordTableSession(deterministicOptions({
      source: { kind: "local", rows: [employee()], getRowId: (row) => row.id, resetKey: 1 }
    }));
    await session.dispatch({ type: "edit-cells", edits: [{ rowId: "e1", columnId: "name", rawText: "Grace" }] });
    expect(session.getSnapshot().canUndo).toBe(true);
    session.updateOptions(deterministicOptions({
      source: { kind: "local", rows: [employee("e1", "Katherine")], getRowId: (row) => row.id, resetKey: 2 }
    }));
    expect(session.getSnapshot().getCell("e1", "name").storedValue).toBe("Katherine");
    expect(session.getSnapshot().canUndo).toBe(false);
  });

  it("accepts a changed controlled rows reference synchronously", () => {
    const listener = vi.fn();
    const onRowsChange = vi.fn();
    const session = createLocalRecordTableSession(deterministicOptions({
      source: { kind: "local", rows: [employee()], getRowId: (row) => row.id, onRowsChange }
    }));
    session.subscribe(listener);
    session.updateOptions(deterministicOptions({
      source: { kind: "local", rows: [employee("e1", "Grace")], getRowId: (row) => row.id, onRowsChange }
    }));
    expect(session.getSnapshot().getCell("e1", "name").storedValue).toBe("Grace");
    expect(listener).not.toHaveBeenCalled();
  });

  it("resets safely when switching controlled and uncontrolled modes", async () => {
    const onRowsChange = vi.fn();
    const session = createLocalRecordTableSession(deterministicOptions({
      source: { kind: "local", rows: [employee()], getRowId: (row) => row.id, onRowsChange }
    }));
    await session.dispatch({ type: "edit-cells", edits: [{ rowId: "e1", columnId: "name", rawText: "Grace" }] });
    session.updateOptions(deterministicOptions({
      source: { kind: "local", rows: [employee("e1", "Katherine")], getRowId: (row) => row.id }
    }));
    expect(session.getSnapshot().getCell("e1", "name").storedValue).toBe("Katherine");
    expect(session.getSnapshot().canUndo).toBe(false);
  });

  it("updates only the controlled sorting slice through onStateChange", async () => {
    const onStateChange = vi.fn();
    const session = createLocalRecordTableSession(deterministicOptions({ state: { sorting: [] }, onStateChange }));
    const result = await session.dispatch({ type: "set-sorting", sorting: [{ columnId: "name", direction: "asc" }] });
    const updater = onStateChange.mock.calls[0][0] as (state: TableViewState) => TableViewState;
    const next = updater(session.getSnapshot().state);
    expect(result).toMatchObject({ status: "committed", changed: true });
    expect(next.sorting).toEqual([{ columnId: "name", direction: "asc" }]);
    expect(next.filter).toBeNull();
    expect(onStateChange).toHaveBeenCalledTimes(1);
  });

  it("updates cell metadata through one collision-safe document updater", async () => {
    const onDocumentChange = vi.fn();
    const document = { version: 1 as const, cells: {}, calculatedColumns: [], namedStyles: [] };
    const session = createLocalRecordTableSession(deterministicOptions({ document, onDocumentChange }));
    await session.dispatch({
      type: "update-cell-metadata",
      updates: [{ rowId: "e1", columnId: "name", patch: { comment: "note" } }]
    });
    const updater = onDocumentChange.mock.calls[0][0] as (value: typeof document) => typeof document;
    expect(updater(document).cells).toEqual({
      [createTableMetadataKey("e1", "name")]: { comment: "note" }
    });
    expect(session.getSnapshot().getCell("e1", "name").metadata.comment).toBe("note");
  });

  it("recalculates a computed column after an accessor edit", async () => {
    const helper = createColumnHelper<Employee>();
    const session = createLocalRecordTableSession(deterministicOptions({
      columns: [
        ...createColumns(),
        helper.computed<number>({
          id: "doubleSalary",
          header: "Double",
          calculate: ({ getValue }) => Number(getValue("salary")) * 2
        })
      ]
    }));
    expect(session.getSnapshot().getCell("e1", "doubleSalary").evaluatedValue).toBe(200);
    await session.dispatch({ type: "edit-cells", edits: [{ rowId: "e1", columnId: "salary", rawText: "125" }] });
    expect(session.getSnapshot().getCell("e1", "doubleSalary").evaluatedValue).toBe(250);
  });

  it("reports a calculated-column cycle as a cell issue without recursive overflow", () => {
    const helper = createColumnHelper<Employee>();
    const session = createLocalRecordTableSession(deterministicOptions({
      columns: [
        helper.computed<number>({ id: "a", header: "A", calculate: ({ getValue }) => Number(getValue("b")) }),
        helper.computed<number>({ id: "b", header: "B", calculate: ({ getValue }) => Number(getValue("a")) })
      ]
    }));
    expect(session.getSnapshot().getCell("e1", "a")).toMatchObject({ displayValue: "#ERROR!" });
    expect(session.getSnapshot().getCell("e1", "a").issues).toContainEqual(expect.objectContaining({ code: "calculated-column-cycle" }));
  });

  it("rejects formula text when no formula service is configured", async () => {
    const session = createLocalRecordTableSession(deterministicOptions());
    const result = await session.dispatch({
      type: "edit-cells",
      edits: [{ rowId: "e1", columnId: "salary", rawText: "=salary*2" }]
    });
    expect(result).toMatchObject({ status: "rejected", reason: "unsupported" });
    expect(session.getSnapshot().getCell("e1", "salary").storedValue).toBe(100);
  });

  it("uses a configured formula service and validates its evaluated value", async () => {
    const formulaService = {
      evaluate: vi.fn(({ expression }: { expression: string }) => expression === "=negative"
        ? { value: -1, displayValue: "-1" }
        : { value: 250, displayValue: "250" })
    };
    const session = createLocalRecordTableSession(deterministicOptions({ formulaService }));
    const rejected = await session.dispatch({ type: "edit-cells", edits: [{ rowId: "e1", columnId: "salary", rawText: "=negative" }] });
    const committed = await session.dispatch({ type: "edit-cells", edits: [{ rowId: "e1", columnId: "salary", rawText: "=positive" }] });
    expect(rejected).toMatchObject({ status: "rejected", reason: "validation" });
    expect(committed).toMatchObject({ status: "committed", changed: true });
    expect(session.getSnapshot().getCell("e1", "salary")).toMatchObject({
      storedValue: 250,
      evaluatedValue: 250,
      displayValue: "250",
      formula: "=positive"
    });
  });

  it("rejects edits after a permission predicate changes", async () => {
    let permitted = true;
    const helper = createColumnHelper<Employee>();
    const session = createLocalRecordTableSession(deterministicOptions({
      columns: [helper.accessor("name", { id: "name", header: "Name", permitted: () => permitted })]
    }));
    permitted = false;
    const result = await session.dispatch({ type: "edit-cells", edits: [{ rowId: "e1", columnId: "name", rawText: "Grace" }] });
    expect(result).toMatchObject({ status: "rejected", reason: "permission" });
  });

  it("sanitizes accessor failures during dispatch instead of rejecting the promise", async () => {
    const session = createLocalRecordTableSession(deterministicOptions({
      columns: [{
        kind: "accessor",
        id: "name",
        header: "Name",
        accessor: () => { throw new Error("secret row value"); },
        update: (row, value: string) => ({ ...row, name: value })
      }]
    }));
    const result = await session.dispatch({ type: "edit-cells", edits: [{ rowId: "e1", columnId: "name", rawText: "Grace" }] });
    expect(result).toEqual({
      status: "rejected",
      reason: "unsupported",
      issues: [{
        code: "TABLE_EXTENSION_ERROR",
        message: "Host accessor extension failed",
        rowId: "e1",
        columnId: "name"
      }]
    });
  });

  it("suppresses no-op publications and reports changed false", async () => {
    const listener = vi.fn();
    const session = createLocalRecordTableSession(deterministicOptions());
    session.subscribe(listener);
    const result = await session.dispatch({ type: "edit-cells", edits: [{ rowId: "e1", columnId: "name", rawText: "Ada" }] });
    expect(result).toMatchObject({ status: "committed", changed: false });
    expect(listener).not.toHaveBeenCalled();
  });

  it("exports the current view and complete local dataset as platform-neutral artifacts with explicit scope", async () => {
    const session = createLocalRecordTableSession(deterministicOptions({
      source: {
        kind: "local",
        rows: [employee(), employee("e2", "Grace", 200)],
        getRowId: (row) => row.id
      },
      defaultState: { pagination: { kind: "offset", offset: 0, limit: 1 } }
    }));
    const current = await session.export({ format: "csv", scope: "currentView", fileName: "current" });
    const complete = await session.export({ format: "csv", scope: "completeDataset", fileName: "all.csv" });
    const xlsx = await session.export({ format: "xlsx", scope: "completeDataset", fileName: "all" });
    expect(new TextDecoder().decode(current.bytes)).toBe("Name,Salary,Active\r\nAda,100,TRUE\r\n");
    expect(new TextDecoder().decode(complete.bytes)).toBe("Name,Salary,Active\r\nAda,100,TRUE\r\nGrace,200,TRUE\r\n");
    expect(current).toMatchObject({ mediaType: "text/csv;charset=utf-8", fileName: "current.csv" });
    expect(xlsx.mediaType).toBe("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    expect(xlsx.fileName).toBe("all.xlsx");
    expect(xlsx.bytes.byteLength).toBeGreaterThan(0);
  });

  it("rejects remote conflict-resolution intents without changing local state", async () => {
    const listener = vi.fn();
    const session = createLocalRecordTableSession(deterministicOptions());
    session.subscribe(listener);
    const before = session.getSnapshot();
    const reload = await session.dispatch({ type: "reload-authoritative", operationId: "op", rowId: "e1" });
    const retry = await session.dispatch({ type: "retry-with-revision", operationId: "op", rowId: "e1", expectedRevision: "2" });
    expect(reload).toMatchObject({ status: "rejected", reason: "unsupported" });
    expect(retry).toMatchObject({ status: "rejected", reason: "unsupported" });
    expect(session.getSnapshot()).toBe(before);
    expect(listener).not.toHaveBeenCalled();
  });

  it("does not publish after destroy", async () => {
    const listener = vi.fn();
    const session = createLocalRecordTableSession(deterministicOptions());
    session.subscribe(listener);
    session.destroy();
    const result = await session.dispatch({ type: "edit-cells", edits: [{ rowId: "e1", columnId: "name", rawText: "Grace" }] });
    expect(result).toMatchObject({ status: "rejected", reason: "unsupported" });
    expect(listener).not.toHaveBeenCalled();
  });

  it("keeps two sessions completely isolated", async () => {
    const first = createLocalRecordTableSession(deterministicOptions());
    const second = createLocalRecordTableSession(deterministicOptions());
    await first.dispatch({ type: "edit-cells", edits: [{ rowId: "e1", columnId: "name", rawText: "Grace" }] });
    expect(first.getSnapshot().getCell("e1", "name").storedValue).toBe("Grace");
    expect(second.getSnapshot().getCell("e1", "name").storedValue).toBe("Ada");
  });

  it("uses collision-resistant per-session command ids in change contexts and diagnostics", async () => {
    const onRowsChange = vi.fn();
    const onDiagnostic = vi.fn();
    const session = createLocalRecordTableSession(deterministicOptions({
      source: { kind: "local", rows: [employee()], getRowId: (row) => row.id, onRowsChange },
      onDiagnostic
    }));
    await session.dispatch({ type: "edit-cells", edits: [{ rowId: "e1", columnId: "name", rawText: "Grace" }] });
    expect(onRowsChange.mock.calls[0][1]).toMatchObject({ commandId: "local-test:1", reason: "edit-cells", revision: "1" });
    expect(onDiagnostic).toHaveBeenCalledWith(expect.objectContaining({
      commandId: "local-test:1",
      metadata: expect.objectContaining({ commandType: "edit-cells", changed: true })
    }));
    expect(JSON.stringify(onDiagnostic.mock.calls)).not.toContain("Grace");
  });

  it("normalizes grouping and rejects pagination conflicts atomically", async () => {
    const session = createLocalRecordTableSession(deterministicOptions({
      defaultState: { pagination: { kind: "offset", offset: 0, limit: 25 } }
    }));
    await session.dispatch({ type: "set-grouping", grouping: [{ columnId: "active" }] });
    expect(session.getSnapshot().state.pagination).toEqual({ kind: "none" });
    const rejected = await session.dispatch({ type: "set-pagination", pagination: { kind: "offset", offset: 0, limit: 5 } });
    expect(rejected).toEqual({
      status: "rejected",
      reason: "validation",
      issues: [{ code: "TABLE_GROUPING_PAGINATION_CONFLICT", message: "Grouping requires pagination kind none" }]
    });
  });

  it("preserves a last valid projection when controlled state is invalid", () => {
    const onStateChange = vi.fn();
    const session = createLocalRecordTableSession(deterministicOptions({ state: { grouping: [] }, onStateChange }));
    session.updateOptions(deterministicOptions({
      state: {
        grouping: [{ columnId: "active" }],
        pagination: { kind: "offset", offset: 0, limit: 5 }
      },
      onStateChange
    }));
    expect(session.getSnapshot().status.phase).toBe("error");
    expect(session.getSnapshot().issues).toContainEqual(expect.objectContaining({ code: "TABLE_GROUPING_PAGINATION_CONFLICT" }));
    expect(session.getSnapshot().operationStates.pagination.enabled).toBe(false);
    expect(onStateChange).toHaveBeenCalledTimes(1);
  });

  it("isolates subscriber failures and continues publishing", async () => {
    const first = vi.fn(() => { throw new Error("secret"); });
    const second = vi.fn();
    const session = createLocalRecordTableSession(deterministicOptions());
    session.subscribe(first);
    session.subscribe(second);
    await session.dispatch({ type: "edit-cells", edits: [{ rowId: "e1", columnId: "name", rawText: "Grace" }] });
    expect(first).toHaveBeenCalledTimes(1);
    expect(second).toHaveBeenCalledTimes(1);
  });

  it("rejects invalid controlled ownership combinations", () => {
    const document = { version: 1 as const, cells: {}, calculatedColumns: [], namedStyles: [] };
    expect(() => createLocalRecordTableSession(deterministicOptions({ document, defaultDocument: document })))
      .toThrow("document and defaultDocument are mutually exclusive");
    expect(() => createLocalRecordTableSession(deterministicOptions({ document })))
      .toThrow("Controlled document requires onDocumentChange");
    expect(() => createLocalRecordTableSession(deterministicOptions({ state: { sorting: [] } })))
      .toThrow("Controlled state requires onStateChange");
  });
});
