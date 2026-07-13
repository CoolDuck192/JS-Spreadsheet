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

  it("stores temporal edits as ISO strings across display, filter, sort, and grouping", async () => {
    type TemporalRow = { id: string; day: string; instant: string };
    const helper = createColumnHelper<TemporalRow>();
    const rows: readonly TemporalRow[] = [
      { id: "edited", day: "2026-01-15", instant: "2026-01-15T12:00:00.000Z" },
      { id: "later", day: "2026-04-01", instant: "2026-04-01T12:00:00.000Z" }
    ];
    const onRowsChange = vi.fn();
    const session = createLocalRecordTableSession({
      source: { kind: "local", rows, getRowId: (row) => row.id, onRowsChange },
      columns: [
        helper.accessor("day", { id: "day", header: "Day", dataType: "date" }),
        helper.accessor("instant", { id: "instant", header: "Instant", dataType: "datetime" })
      ]
    });

    const result = await session.dispatch({
      type: "edit-cells",
      edits: [
        { rowId: "edited", columnId: "day", rawText: "2026-03-05" },
        { rowId: "edited", columnId: "instant", rawText: "2026-03-05T07:30:00-05:00" }
      ]
    });

    expect(result).toMatchObject({ status: "committed", changed: true });
    const updater = onRowsChange.mock.calls[0][0] as (source: readonly TemporalRow[]) => readonly TemporalRow[];
    expect(updater(rows)[0]).toEqual({
      id: "edited",
      day: "2026-03-05",
      instant: "2026-03-05T12:30:00.000Z"
    });
    expect(session.getSnapshot().getCell("edited", "day")).toMatchObject({
      storedValue: "2026-03-05",
      evaluatedValue: "2026-03-05",
      displayValue: "2026-03-05"
    });
    expect(session.getSnapshot().getCell("edited", "instant")).toMatchObject({
      storedValue: "2026-03-05T12:30:00.000Z",
      displayValue: "2026-03-05T12:30:00.000Z"
    });

    await session.dispatch({
      type: "set-filter",
      filter: {
        kind: "comparison",
        columnId: "day",
        operator: "eq",
        value: { type: "date", value: "2026-03-05" }
      }
    });
    expect(session.getSnapshot().rows.filter((row) => row.kind === "data").map((row) => row.id))
      .toEqual(["edited"]);

    await session.dispatch({
      type: "set-filter",
      filter: {
        kind: "comparison",
        columnId: "instant",
        operator: "eq",
        value: { type: "datetime", value: "2026-03-05T12:30" }
      }
    });
    expect(session.getSnapshot().rows.filter((row) => row.kind === "data").map((row) => row.id))
      .toEqual(["edited"]);

    await session.dispatch({ type: "set-filter", filter: null });
    await session.dispatch({
      type: "set-sorting",
      sorting: [{ columnId: "instant", direction: "asc" }]
    });
    expect(session.getSnapshot().rows.filter((row) => row.kind === "data").map((row) => row.id))
      .toEqual(["edited", "later"]);

    await session.dispatch({ type: "set-grouping", grouping: [{ columnId: "day" }] });
    const editedGroup = session.getSnapshot().rows.find((row) =>
      row.kind === "group" && row.key.type === "date" && row.key.value === "2026-03-05"
    );
    expect(editedGroup).toBeDefined();
    expect(session.getSnapshot().getCell(editedGroup!.id, "day").displayValue).toBe("2026-03-05");
  });

  it("rejects Excel's phantom leap-day serial for ISO-backed date columns", async () => {
    type TemporalRow = { id: string; day: string };
    const helper = createColumnHelper<TemporalRow>();
    const session = createLocalRecordTableSession({
      source: {
        kind: "local",
        rows: [{ id: "row-1", day: "1900-02-28" }],
        getRowId: (row) => row.id
      },
      columns: [helper.accessor("day", { id: "day", header: "Day", dataType: "date" })]
    });

    const result = await session.dispatch({
      type: "edit-cells",
      edits: [{ rowId: "row-1", columnId: "day", rawText: "1900-02-29" }]
    });

    expect(result).toMatchObject({
      status: "rejected",
      reason: "validation",
      issues: [{ code: "TABLE_VALUE_TYPE", rowId: "row-1", columnId: "day" }]
    });
    expect(session.getSnapshot().getCell("row-1", "day").storedValue).toBe("1900-02-28");
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

  it("validates an entire edit batch before invoking any row updater", async () => {
    const updateName = vi.fn((row: Employee, value: string) => ({ ...row, name: value }));
    const helper = createColumnHelper<Employee>();
    const session = createLocalRecordTableSession(deterministicOptions({
      columns: [
        {
          kind: "accessor",
          id: "name",
          header: "Name",
          dataType: "text",
          accessor: (row: Employee) => row.name,
          update: updateName
        },
        helper.accessor("salary", {
          id: "salary",
          header: "Salary",
          dataType: "number",
          validate: ({ parsed }) => parsed < 0 ? [{ code: "negative", message: "No negatives" }] : []
        })
      ]
    }));
    const result = await session.dispatch({
      type: "edit-cells",
      edits: [
        { rowId: "e1", columnId: "name", rawText: "Grace" },
        { rowId: "e1", columnId: "salary", rawText: "-1" }
      ]
    });
    expect(result).toMatchObject({ status: "rejected", reason: "validation" });
    expect(updateName).not.toHaveBeenCalled();
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

  it("invalidates a memoized snapshot when formula capabilities or features change", () => {
    const source = { kind: "local" as const, rows: [employee()], getRowId: (row: Employee) => row.id };
    const columns = createColumns();
    const session = createLocalRecordTableSession({ source, columns });
    const withoutFormula = session.getSnapshot();
    const formulaService = { evaluate: () => ({ value: 1, displayValue: "1" }) };

    session.updateOptions({ source, columns, formulaService });
    const withFormula = session.getSnapshot();
    expect(withFormula).not.toBe(withoutFormula);
    expect(withFormula.operationStates.formula.enabled).toBe(true);

    session.updateOptions({ source, columns, formulaService, features: { formula: false } });
    expect(session.getSnapshot()).not.toBe(withFormula);
    expect(session.getSnapshot().operationStates.formula).toMatchObject({ enabled: false });
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
    const otherKey = createTableMetadataKey("e1", "salary");
    const document = {
      version: 1 as const,
      cells: { [otherKey]: { comment: "original" } },
      calculatedColumns: [],
      namedStyles: []
    };
    const session = createLocalRecordTableSession(deterministicOptions({ document, onDocumentChange }));
    await session.dispatch({
      type: "update-cell-metadata",
      updates: [{ rowId: "e1", columnId: "name", patch: { comment: "note" } }]
    });
    const updater = onDocumentChange.mock.calls[0][0] as (value: typeof document) => typeof document;
    expect(updater({ ...document, cells: { [otherKey]: { comment: "concurrent" } } }).cells).toEqual({
      [createTableMetadataKey("e1", "name")]: { comment: "note" },
      [otherKey]: { comment: "concurrent" }
    });
    expect(session.getSnapshot().getCell("e1", "name").metadata.comment).toBe("note");
  });

  it("applies repeated metadata targets with deterministic last-write-wins semantics", async () => {
    const session = createLocalRecordTableSession(deterministicOptions());

    await expect(session.dispatch({
      type: "update-cell-metadata",
      updates: [
        {
          rowId: "e1", columnId: "name",
          patch: { comment: "first", format: { bold: true } }
        },
        {
          rowId: "e1", columnId: "name",
          patch: { comment: "last", validation: { kind: "textLength", min: 1 } }
        },
        {
          rowId: "e1", columnId: "name",
          patch: { format: { italic: true }, formula: "=A1", readOnly: true }
        }
      ]
    })).resolves.toMatchObject({ status: "committed", changed: true });

    expect(session.getSnapshot().getCell("e1", "name").metadata).toEqual({
      comment: "last",
      format: { italic: true },
      validation: { kind: "textLength", min: 1 },
      formula: "=A1",
      readOnly: true
    });
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

  it("rejects formulas disabled by host configuration even when a service exists", async () => {
    const session = createLocalRecordTableSession(deterministicOptions({
      formulaService: { evaluate: () => ({ value: 250, displayValue: "250" }) },
      features: { formula: false }
    }));
    const result = await session.dispatch({
      type: "edit-cells",
      edits: [{ rowId: "e1", columnId: "salary", rawText: "=positive" }]
    });
    expect(result).toMatchObject({ status: "rejected", reason: "unsupported" });
    expect(session.getSnapshot().getCell("e1", "salary").storedValue).toBe(100);
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

  it("accepts changed column predicates when accessor identities stay stable", async () => {
    const source = { kind: "local" as const, rows: [employee()], getRowId: (row: Employee) => row.id };
    const accessor = (row: Employee) => row.name;
    const update = (row: Employee, value: string) => ({ ...row, name: value });
    const columns = [{ kind: "accessor" as const, id: "name", header: "Name", accessor, update, permitted: () => true }];
    const session = createLocalRecordTableSession({ source, columns });
    session.getSnapshot();
    session.updateOptions({
      source,
      columns: [{ ...columns[0], permitted: () => false }]
    });
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

  it("rejects export when the host disables the export feature", async () => {
    const session = createLocalRecordTableSession(deterministicOptions({
      features: { export: false }
    }));

    expect(session.getSnapshot().operationStates.export).toMatchObject({ enabled: false });
    await expect(session.export({ format: "csv", scope: "completeDataset" }))
      .rejects.toMatchObject({ code: "TABLE_CAPABILITY_UNSUPPORTED" });
  });

  it("includes collapsed descendants in a complete-dataset export", async () => {
    type TreeEmployee = Employee & { children?: readonly TreeEmployee[] };
    const helper = createColumnHelper<TreeEmployee>();
    const rows: readonly TreeEmployee[] = [
      { ...employee("parent", "Parent"), children: [{ ...employee("child", "Child") }] }
    ];
    const source = {
      kind: "local" as const,
      rows,
      getRowId: (row: TreeEmployee) => row.id,
      getSubRows: (row: TreeEmployee) => row.children
    };
    const session = createLocalRecordTableSession<TreeEmployee>({
      source,
      columns: [helper.accessor("name", { id: "name", header: "Name" })]
    });
    const current = await session.export({ format: "csv", scope: "currentView" });
    const complete = await session.export({ format: "csv", scope: "completeDataset" });
    expect(new TextDecoder().decode(current.bytes)).toBe("Name\r\nParent\r\n");
    expect(new TextDecoder().decode(complete.bytes)).toBe("Name\r\nParent\r\nChild\r\n");
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

  it.each([
    [
      "cursor pagination",
      deterministicOptions(),
      { type: "set-pagination", pagination: { kind: "cursor", limit: 25 } }
    ],
    [
      "infinite pagination",
      deterministicOptions(),
      { type: "set-pagination", pagination: { kind: "infinite", limit: 25 } }
    ],
    [
      "tree grouping",
      deterministicOptions({ source: {
        kind: "local", rows: [employee()], getRowId: (row) => row.id, getSubRows: () => []
      } }),
      { type: "set-grouping", grouping: [{ columnId: "name" }] }
    ],
    [
      "tree pagination",
      deterministicOptions({ source: {
        kind: "local", rows: [employee()], getRowId: (row) => row.id, getSubRows: () => []
      } }),
      { type: "set-pagination", pagination: { kind: "offset", offset: 0, limit: 25 } }
    ],
    [
      "sorting by an unknown column",
      deterministicOptions(),
      { type: "set-sorting", sorting: [{ columnId: "missing", direction: "asc" }] }
    ],
    [
      "filtering by an unknown column",
      deterministicOptions(),
      {
        type: "set-filter",
        filter: {
          kind: "comparison", columnId: "missing", operator: "eq",
          value: { type: "string", value: "Ada" }
        }
      }
    ],
    [
      "grouping by an unknown column",
      deterministicOptions(),
      { type: "set-grouping", grouping: [{ columnId: "missing" }] }
    ],
    [
      "aggregating an unknown column",
      deterministicOptions(),
      { type: "set-aggregates", aggregates: [{ id: "missing", columnId: "missing", function: "count" }] }
    ]
  ] as const)("rejects %s before it can brick the local session", async (_label, options, intent) => {
    const session = createLocalRecordTableSession(options);
    const before = session.getSnapshot();

    const result = await session.dispatch(intent);

    expect(result).toMatchObject({ status: "rejected", reason: "validation" });
    expect(session.getSnapshot()).toBe(before);
  });

  it.each([
    [
      "default sorting on an unknown column",
      { defaultState: { sorting: [{ columnId: "missing", direction: "asc" as const }] } }
    ],
    [
      "controlled sorting on an unknown column",
      {
        state: { sorting: [{ columnId: "missing", direction: "asc" as const }] },
        onStateChange: vi.fn()
      }
    ],
    [
      "cursor pagination",
      { defaultState: { pagination: { kind: "cursor" as const, limit: 25 } } }
    ],
    [
      "grouping with pagination",
      {
        defaultState: {
          grouping: [{ columnId: "active" }],
          pagination: { kind: "offset" as const, offset: 0, limit: 25 }
        }
      }
    ]
  ] as const)("surfaces invalid initial %s as a recoverable error snapshot", async (_label, overrides) => {
    const session = createLocalRecordTableSession(deterministicOptions(overrides));

    expect(() => session.getSnapshot()).not.toThrow();
    expect(session.getSnapshot()).toMatchObject({
      status: { phase: "error" },
      issues: [expect.objectContaining({ code: expect.stringMatching(/^TABLE_/) })]
    });

    await expect(session.dispatch({
      type: "set-sorting",
      sorting: [{ columnId: "name", direction: "asc" }]
    })).resolves.toMatchObject({ status: "committed" });
    expect(session.getSnapshot().status.phase).toBe("ready");
    expect(session.getSnapshot().state.sorting).toEqual([{ columnId: "name", direction: "asc" }]);
  });

  it("invalidates a cached initial-state error when unchanged options confirm the sanitized state", () => {
    const stableOptions = deterministicOptions({
      defaultState: {
        grouping: [{ columnId: "active" }],
        pagination: { kind: "offset", offset: 0, limit: 25 }
      }
    });
    const session = createLocalRecordTableSession(stableOptions);
    const invalidSnapshot = session.getSnapshot();
    expect(invalidSnapshot.status.phase).toBe("error");

    session.updateOptions(stableOptions);

    expect(session.getSnapshot()).not.toBe(invalidSnapshot);
    expect(session.getSnapshot()).toMatchObject({
      status: { phase: "ready" },
      state: { grouping: [{ columnId: "active" }], pagination: { kind: "none" } },
      issues: []
    });
  });

  it("prunes query state that references columns removed by updateOptions", () => {
    const source = { kind: "local" as const, rows: [employee()], getRowId: (row: Employee) => row.id };
    const initialColumns = createColumns();
    const session = createLocalRecordTableSession({
      source,
      columns: initialColumns,
      defaultState: {
        sorting: [{ columnId: "name", direction: "asc" }],
        filter: {
          kind: "comparison", columnId: "name", operator: "eq",
          value: { type: "string", value: "Ada" }
        },
        grouping: [{ columnId: "name" }],
        aggregates: [{ id: "name-count", columnId: "name", function: "count" }]
      }
    });

    session.getSnapshot();
    session.updateOptions({ source, columns: [initialColumns[1], initialColumns[2]] });

    expect(session.getSnapshot().state).toMatchObject({
      sorting: [],
      filter: null,
      grouping: [],
      aggregates: []
    });
  });

  it("renders aggregate values under their source column when aggregate ids differ", async () => {
    const session = createLocalRecordTableSession(deterministicOptions({
      source: {
        kind: "local",
        rows: [employee("e1", "Ada", 100), employee("e2", "Grace", 200)],
        getRowId: (row) => row.id
      }
    }));
    await session.dispatch({ type: "set-grouping", grouping: [{ columnId: "active" }] });
    await session.dispatch({
      type: "set-aggregates",
      aggregates: [{ id: "salary-sum", columnId: "salary", function: "sum" }]
    });

    const group = session.getSnapshot().rows.find((row) => row.kind === "group");
    expect(group).toBeDefined();
    expect(session.getSnapshot().getCell(group!.id, "salary").displayValue).toBe("300");
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

  it("inserts supplied local records before a stable row id", async () => {
    const session = createLocalRecordTableSession(deterministicOptions({
      source: { kind: "local", rows: [employee("e1", "Ada"), employee("e3", "Linus")], getRowId: (row) => row.id }
    }));
    const result = await session.dispatch({
      type: "insert-rows",
      rows: [employee("e2", "Grace")],
      beforeRowId: "e3"
    });
    expect(result).toMatchObject({ status: "committed", changed: true });
    expect(session.getSnapshot().rows.filter((row) => row.kind === "data").map((row) => row.id))
      .toEqual(["e1", "e2", "e3"]);
  });

  it("inserts supplied local records after a stable row id", async () => {
    const session = createLocalRecordTableSession(deterministicOptions({
      source: { kind: "local", rows: [employee("e1", "Ada"), employee("e3", "Linus")], getRowId: (row) => row.id }
    }));
    await session.dispatch({ type: "insert-rows", rows: [employee("e2", "Grace")], afterRowId: "e1" });
    expect(session.getSnapshot().rows.filter((row) => row.kind === "data").map((row) => row.id))
      .toEqual(["e1", "e2", "e3"]);
  });

  it("appends supplied local records when no anchor is present", async () => {
    const session = createLocalRecordTableSession(deterministicOptions());
    await session.dispatch({ type: "insert-rows", rows: [employee("e2", "Grace"), employee("e3", "Linus")] });
    expect(session.getSnapshot().rows.filter((row) => row.kind === "data").map((row) => row.id))
      .toEqual(["e1", "e2", "e3"]);
  });

  it("rejects count-only insertion because local records require host values", async () => {
    const session = createLocalRecordTableSession(deterministicOptions());
    const before = session.getSnapshot();
    const result = await session.dispatch({ type: "insert-rows", count: 2 });
    expect(result).toMatchObject({ status: "rejected", reason: "unsupported" });
    expect(session.getSnapshot()).toBe(before);
  });

  it("rejects duplicate inserted row ids atomically", async () => {
    const session = createLocalRecordTableSession(deterministicOptions());
    const listener = vi.fn();
    session.subscribe(listener);
    const result = await session.dispatch({
      type: "insert-rows",
      rows: [employee("e2", "Grace"), employee("e2", "Duplicate")]
    });
    expect(result).toMatchObject({ status: "rejected", reason: "validation" });
    expect(session.getSnapshot().rows.filter((row) => row.kind === "data").map((row) => row.id)).toEqual(["e1"]);
    expect(session.getSnapshot().canUndo).toBe(false);
    expect(listener).not.toHaveBeenCalled();
  });

  it("rejects ambiguous runtime row anchors atomically", async () => {
    const session = createLocalRecordTableSession(deterministicOptions());
    const result = await session.dispatch({
      type: "insert-rows",
      rows: [employee("e2", "Grace")],
      beforeRowId: "e1",
      afterRowId: "e1"
    } as unknown as Parameters<typeof session.dispatch>[0]);
    expect(result).toMatchObject({ status: "rejected", reason: "validation" });
    expect(session.getSnapshot().rows.filter((row) => row.kind === "data").map((row) => row.id)).toEqual(["e1"]);
  });

  it("emits controlled structural updaters that preserve unrelated host rows", async () => {
    const onRowsChange = vi.fn();
    const original = [employee("e1", "Ada")];
    const session = createLocalRecordTableSession(deterministicOptions({
      source: { kind: "local", rows: original, getRowId: (row) => row.id, onRowsChange }
    }));
    await session.dispatch({ type: "insert-rows", rows: [employee("e2", "Grace")], afterRowId: "e1" });
    const insert = onRowsChange.mock.calls[0][0] as (rows: readonly Employee[]) => readonly Employee[];
    expect(insert([employee("external", "External"), { ...employee(), salary: 999 }]))
      .toEqual([employee("external", "External"), { ...employee(), salary: 999 }, employee("e2", "Grace")]);

    await session.dispatch({ type: "delete-rows", rowIds: ["e1"] });
    const remove = onRowsChange.mock.calls[1][0] as (rows: readonly Employee[]) => readonly Employee[];
    expect(remove([employee("external", "External"), employee(), employee("e2", "Grace")]))
      .toEqual([employee("external", "External"), employee("e2", "Grace")]);
  });

  it("deletes stable row ids regardless of current sort order", async () => {
    const session = createLocalRecordTableSession(deterministicOptions({
      source: {
        kind: "local",
        rows: [employee("e1", "Ada"), employee("e2", "Grace"), employee("e3", "Linus")],
        getRowId: (row) => row.id
      },
      defaultState: { sorting: [{ columnId: "name", direction: "desc" }] }
    }));
    await session.dispatch({ type: "delete-rows", rowIds: ["e2"] });
    await session.dispatch({ type: "set-sorting", sorting: [] });
    expect(session.getSnapshot().rows.filter((row) => row.kind === "data").map((row) => row.id))
      .toEqual(["e1", "e3"]);
  });

  it("prunes deleted row ids from cell, row, and expansion selection state", async () => {
    const onStateChange = vi.fn();
    const session = createLocalRecordTableSession(deterministicOptions({
      source: {
        kind: "local",
        rows: [employee("e1", "Ada"), employee("e2", "Grace"), employee("e3", "Linus")],
        getRowId: (row) => row.id
      },
      state: {
        selection: {
          anchor: { rowId: "e2", columnId: "name" },
          focus: { rowId: "e2", columnId: "name" }
        },
        selectedRowIds: ["e1", "e2", "e3"],
        expandedRowIds: ["e2", "e3"]
      },
      onStateChange
    }));
    const beforeState = session.getSnapshot().state;

    expect(await session.dispatch({ type: "delete-rows", rowIds: ["e2"] }))
      .toMatchObject({ status: "committed" });
    const expectedState = {
      selection: null,
      selectedRowIds: ["e1", "e3"],
      expandedRowIds: ["e3"]
    };
    expect(session.getSnapshot().state).toMatchObject(expectedState);
    expect(onStateChange).toHaveBeenCalledTimes(1);
    expect(onStateChange.mock.calls[0][0](beforeState)).toMatchObject(expectedState);

    await session.dispatch({ type: "insert-rows", rows: [employee("e2", "New row")] });
    expect(session.getSnapshot().state.selectedRowIds).toEqual(["e1", "e3"]);
    expect(session.getSnapshot().state.expandedRowIds).toEqual(["e3"]);
    expect(await session.dispatch({ type: "delete-rows", rowIds: ["e1", "e3"] }))
      .toMatchObject({ status: "committed" });
  });

  it("prunes inserted-row selection through undo and redo cycles", async () => {
    const session = createLocalRecordTableSession(deterministicOptions());
    await session.dispatch({ type: "insert-rows", rows: [employee("e2", "Grace")] });
    await session.dispatch({
      type: "set-selection",
      selection: {
        anchor: { rowId: "e2", columnId: "name" },
        focus: { rowId: "e2", columnId: "name" }
      }
    });
    await session.dispatch({ type: "set-row-selection", rowIds: ["e2"] });
    await session.dispatch({ type: "set-row-expanded", rowId: "e2", expanded: true });

    await session.undo();
    expect(session.getSnapshot().rows.map((row) => row.id)).toEqual(["e1"]);
    expect(session.getSnapshot().state).toMatchObject({
      selection: null,
      selectedRowIds: [],
      expandedRowIds: []
    });

    await session.redo();
    expect(session.getSnapshot().rows.map((row) => row.id)).toEqual(["e1", "e2"]);
    expect(session.getSnapshot().state).toMatchObject({
      selection: null,
      selectedRowIds: [],
      expandedRowIds: []
    });

    await session.dispatch({ type: "set-row-selection", rowIds: ["e2"] });
    await session.undo();
    expect(session.getSnapshot().rows.map((row) => row.id)).toEqual(["e1"]);
    expect(session.getSnapshot().state.selectedRowIds).toEqual([]);
  });

  it("undoes and redoes values, row order, and metadata together", async () => {
    const metadataKey = createTableMetadataKey("e2", "name");
    const session = createLocalRecordTableSession(deterministicOptions({
      source: {
        kind: "local",
        rows: [employee("e1", "Ada"), employee("e2", "Grace", 200), employee("e3", "Linus", 300)],
        getRowId: (row) => row.id
      },
      defaultDocument: {
        version: 1,
        cells: { [metadataKey]: { comment: "retain me" } },
        calculatedColumns: [],
        namedStyles: []
      }
    }));

    await session.dispatch({ type: "delete-rows", rowIds: ["e2"] });
    expect(session.getSnapshot().getRowIndex("e2")).toBe(-1);
    const undone = await session.undo();
    expect(undone).toMatchObject({ status: "committed", changed: true });
    expect(session.getSnapshot().rows.filter((row) => row.kind === "data").map((row) => row.id))
      .toEqual(["e1", "e2", "e3"]);
    expect(session.getSnapshot().getCell("e2", "name")).toMatchObject({
      storedValue: "Grace",
      metadata: { comment: "retain me" }
    });

    await session.redo();
    expect(session.getSnapshot().getRowIndex("e2")).toBe(-1);
    await session.undo();
    expect(session.getSnapshot().getCell("e2", "salary").storedValue).toBe(200);
  });

  it("clears redo after a new durable command", async () => {
    const session = createLocalRecordTableSession(deterministicOptions());
    await session.dispatch({ type: "edit-cells", edits: [{ rowId: "e1", columnId: "name", rawText: "Grace" }] });
    await session.undo();
    expect(session.getSnapshot().canRedo).toBe(true);
    await session.dispatch({
      type: "update-cell-metadata",
      updates: [{ rowId: "e1", columnId: "name", patch: { comment: "new durable command" } }]
    });
    expect(session.getSnapshot().canRedo).toBe(false);
    expect(await session.redo()).toMatchObject({ status: "committed", changed: false });
  });

  it("never retains more than the configured history limit", async () => {
    const session = createLocalRecordTableSession(deterministicOptions({ historyLimit: 2 }));
    for (const name of ["Grace", "Katherine", "Dorothy"]) {
      await session.dispatch({ type: "edit-cells", edits: [{ rowId: "e1", columnId: "name", rawText: name }] });
    }
    expect(await session.undo()).toMatchObject({ status: "committed", changed: true });
    expect(await session.undo()).toMatchObject({ status: "committed", changed: true });
    expect(await session.undo()).toMatchObject({ status: "committed", changed: false });
    expect(session.getSnapshot().getCell("e1", "name").storedValue).toBe("Grace");
  });

  it("replays history onto the latest controlled record without restoring unrelated fields", async () => {
    const onRowsChange = vi.fn();
    const options = deterministicOptions({
      source: { kind: "local", rows: [employee()], getRowId: (row) => row.id, onRowsChange }
    });
    const session = createLocalRecordTableSession(options);
    await session.dispatch({ type: "edit-cells", edits: [{ rowId: "e1", columnId: "name", rawText: "Grace" }] });
    const edit = onRowsChange.mock.calls[0][0] as (rows: readonly Employee[]) => readonly Employee[];
    const hostRows = edit([{ ...employee(), salary: 999 }]);
    session.updateOptions({ ...options, source: { ...options.source, rows: hostRows } });

    await session.undo();
    expect(session.getSnapshot().getCell("e1", "name").storedValue).toBe("Ada");
    expect(session.getSnapshot().getCell("e1", "salary").storedValue).toBe(999);
  });
});
