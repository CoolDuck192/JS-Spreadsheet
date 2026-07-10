import { describe, expect, it, vi } from "vitest";
import type { TableCapabilities } from "../core/capabilities";
import type { QueryRequest, QueryResult, QueryRow } from "../core/query";
import type { ColumnDef, TableCellMetadata, TableViewState } from "../core/types";
import { createTestRemoteSource, defaultRemoteCapabilities } from "./testUtils";
import { createRemoteTableSession } from "./RemoteTableSession";
import type { RemoteMutation, RemoteMutationResult } from "./types";

type Employee = {
  id: string;
  name: string;
  department: string;
  salary: number;
  active: boolean;
  startDate: string;
};

const employees: readonly Employee[] = [{
  id: "employee-1",
  name: "Ada",
  department: "Finance",
  salary: 100,
  active: true,
  startDate: "2026-01-02"
}];

const columns: readonly ColumnDef<Employee>[] = [
  accessorColumn("name", "Name"),
  accessorColumn("department", "Department"),
  accessorColumn("salary", "Salary", "number"),
  accessorColumn("active", "Active", "boolean"),
  accessorColumn("startDate", "Start date", "date")
];

describe("RemoteTableSession", () => {
  it("is inert until start and serializes uncontrolled sort/filter state exactly", async () => {
    const query = vi.fn(async (request: QueryRequest) => offsetResult("r1", employees, request));
    const source = createTestRemoteSource<Employee>({ query });
    const session = createRemoteTableSession({ source, columns });

    expect(query).not.toHaveBeenCalled();
    session.start();
    await waitUntilReady(session);
    expect(query).toHaveBeenCalledTimes(1);

    await session.dispatch({
      type: "set-sorting",
      sorting: [{ columnId: "salary", direction: "desc" }]
    });
    await waitForCalls(query, 2);
    await session.dispatch({
      type: "set-filter",
      filter: {
        kind: "comparison",
        columnId: "department",
        operator: "eq",
        value: { type: "string", value: "Finance" }
      }
    });
    await waitForCalls(query, 3);

    expect(query).toHaveBeenLastCalledWith(
      expect.objectContaining({
        sorting: [{ columnId: "salary", direction: "desc" }],
        filter: {
          kind: "comparison",
          columnId: "department",
          operator: "eq",
          value: { type: "string", value: "Finance" }
        }
      }),
      expect.objectContaining({ generation: 3, operationId: expect.any(String) })
    );
    expect(session.getSnapshot().operationStates.sort.scopeLabel).toBe("Complete dataset");

    session.destroy();
  });

  it("gates loaded-row capabilities when complete-dataset scope is required", () => {
    const capabilities = { ...loadedRowCapabilities(), formula: "loadedRows" as const };
    const source = createTestRemoteSource<Employee>({ capabilities, undoMode: "none" });
    const session = createRemoteTableSession({
      source,
      columns,
      features: {
        sort: { requiredScope: "completeDataset" },
        formula: { requiredScope: "completeDataset" }
      }
    });

    expect(session.getSnapshot().operationStates.sort).toMatchObject({
      enabled: false,
      scopeLabel: "Loaded rows",
      reason: "This table only supports sorting loaded rows."
    });
    expect(session.getSnapshot().operationStates.formula).toMatchObject({
      enabled: false,
      scopeLabel: "Loaded rows",
      reason: "This table only supports formulas for loaded rows."
    });

    session.destroy();
  });

  it("projects lazy server cells, typed accessors, groups, and aggregate keys", async () => {
    const aggregateRequest = { id: "sum-salary", columnId: "salary", function: "sum" as const };
    const items: readonly QueryRow<Employee>[] = [
      { kind: "data", id: employees[0].id, original: employees[0], depth: 1, parentId: "group-finance" },
      {
        kind: "group",
        id: "group-finance",
        depth: 0,
        columnId: "department",
        key: { type: "string", value: "Finance" },
        count: 1,
        aggregates: { "sum-salary": 100 }
      },
      { kind: "aggregate", id: "total", depth: 0, aggregates: { "sum-salary": 100 } }
    ];
    const capabilities = {
      ...defaultRemoteCapabilities(),
      group: { executor: "server", scope: "completeDataset" },
      aggregate: { executor: "server", scope: "completeDataset" }
    } satisfies TableCapabilities;
    const source = createTestRemoteSource<Employee>({
      capabilities,
      readCell(row, columnId) {
        if (columnId === "salary") {
          return {
            storedValue: "=BASE*2",
            evaluatedValue: row.salary,
            displayValue: "$100.00",
            formula: "=BASE*2",
            metadata: { comment: "server value" },
            rowVersion: "row-v1",
            issues: [{ code: "SERVER_NOTE", message: "Reviewed", columnId }]
          };
        }
        const value = row[columnId as keyof Employee];
        return { storedValue: value, evaluatedValue: value };
      },
      query: async (request) => ({
        items,
        revision: "r1",
        completeness: "loadedRows",
        pageInfo: {
          kind: "offset",
          offset: 0,
          limit: 50,
          total: { kind: "unknown" },
          hasMore: false
        }
      })
    });
    const session = createRemoteTableSession({
      source,
      columns,
      defaultState: { aggregates: [aggregateRequest] }
    });
    session.start();
    await waitUntilReady(session);
    const snapshot = session.getSnapshot();

    expect(snapshot.totalRowCount).toEqual({ kind: "unknown" });
    expect(snapshot.getCell("employee-1", "salary")).toMatchObject({
      storedValue: "=BASE*2",
      evaluatedValue: 100,
      displayValue: "$100.00",
      formula: "=BASE*2",
      metadata: { comment: "server value" },
      issues: [{ code: "SERVER_NOTE", message: "Reviewed", columnId: "salary", rowId: "employee-1" }]
    });
    expect(snapshot.getCell("employee-1", "active").displayValue).toBe("TRUE");
    expect(snapshot.getCell("employee-1", "startDate").evaluatedValue).toBe("2026-01-02");
    expect(snapshot.getCell("group-finance", "department").evaluatedValue).toBe("Finance");
    expect(snapshot.getCell("group-finance", "salary").evaluatedValue).toBe(100);
    expect(snapshot.getCell("total", "salary").evaluatedValue).toBe(100);

    session.destroy();
  });

  it("falls back lazily to accessor and computed columns when readCell is absent", async () => {
    const computedColumns: readonly ColumnDef<Employee>[] = [
      ...columns,
      {
        kind: "computed",
        id: "summary",
        header: "Summary",
        calculate: (context) => `${context.getValue("name")}:${context.getValue("salary")}`
      }
    ];
    const source = createTestRemoteSource<Employee>({
      capabilities: loadedRowCapabilities(),
      mutationMode: "none",
      undoMode: "none",
      readCell: undefined,
      query: async (request) => offsetResult("r1", employees, request)
    });
    const session = createRemoteTableSession({ source, columns: computedColumns });
    session.start();
    await waitUntilReady(session);

    expect(session.getSnapshot().getCell("employee-1", "name").displayValue).toBe("Ada");
    expect(session.getSnapshot().getCell("employee-1", "active").displayValue).toBe("TRUE");
    expect(session.getSnapshot().getCell("employee-1", "summary").evaluatedValue).toBe("Ada:100");

    session.destroy();
  });

  it("emits controlled query state without querying until committed options arrive", async () => {
    const query = vi.fn(async (request: QueryRequest) => offsetResult("r1", employees, request));
    const onStateChange = vi.fn();
    const source = createTestRemoteSource<Employee>({ query });
    const initialSorting = [{ columnId: "name", direction: "asc" as const }];
    const session = createRemoteTableSession({
      source,
      columns,
      state: { sorting: initialSorting },
      onStateChange
    });
    session.start();
    await waitUntilReady(session);

    await session.dispatch({
      type: "set-sorting",
      sorting: [{ columnId: "salary", direction: "desc" }]
    });
    expect(onStateChange).toHaveBeenCalledTimes(1);
    expect(query).toHaveBeenCalledTimes(1);

    const updater = onStateChange.mock.calls[0][0] as (state: TableViewState) => TableViewState;
    const committed = updater(session.getSnapshot().state);
    session.updateOptions({ source, columns, state: { sorting: committed.sorting }, onStateChange });
    expect(query).toHaveBeenCalledTimes(1);
    session.start();
    await waitForCalls(query, 2);
    expect(query.mock.calls[1][0].sorting).toEqual([{ columnId: "salary", direction: "desc" }]);

    session.destroy();
  });

  it("keeps layout local, enforces grouping pagination, and rejects Task 4 mutations", async () => {
    const query = vi.fn(async (request: QueryRequest) => offsetResult("r1", employees, request));
    const session = createRemoteTableSession({
      source: createTestRemoteSource({
        query,
        capabilities: loadedRowCapabilities(),
        mutationMode: "none",
        undoMode: "none",
        readCell: undefined
      }),
      columns
    });
    session.start();
    await waitUntilReady(session);

    expect(await session.dispatch({
      type: "set-column-pinning",
      columnId: "name",
      pin: "left"
    })).toMatchObject({ status: "committed" });
    expect(session.getSnapshot().state.columnPinning.left).toEqual(["name"]);
    expect(query).toHaveBeenCalledTimes(1);

    expect(await session.dispatch({
      type: "set-grouping",
      grouping: [{ columnId: "department" }]
    })).toMatchObject({ status: "rejected", reason: "unsupported" });
    expect(await session.dispatch({
      type: "edit-cells",
      edits: [{ rowId: "employee-1", columnId: "name", rawText: "Grace" }]
    })).toMatchObject({ status: "rejected", reason: "unsupported" });

    session.destroy();
  });

  it("prevalidates a typed batch, publishes bounded overlays, and reconciles an authoritative row", async () => {
    let serverRow: Employee = employees[0];
    let serverRevision = "1";
    const acknowledgement = createDeferredMutation<Employee>();
    const mutate = vi.fn(async (_batch: readonly RemoteMutation[]) => acknowledgement.promise);
    const query = vi.fn(async (request: QueryRequest) => offsetResult(serverRevision, [serverRow], request));
    const source = createTestRemoteSource<Employee>({
      query,
      mutate,
      compareRevisions: numericRevisionComparator
    });
    const session = createRemoteTableSession({ source, columns });
    session.start();
    await waitUntilReady(session);

    expect(await session.dispatch({
      type: "edit-cells",
      edits: [
        { rowId: "employee-1", columnId: "salary", rawText: "120" },
        { rowId: "employee-1", columnId: "active", rawText: "not-a-boolean" }
      ]
    })).toMatchObject({ status: "rejected", reason: "validation" });
    expect(mutate).not.toHaveBeenCalled();

    const pending = await session.dispatch({
      type: "edit-cells",
      edits: [
        { rowId: "employee-1", columnId: "salary", rawText: "120" },
        { rowId: "employee-1", columnId: "active", rawText: "FALSE" },
        { rowId: "employee-1", columnId: "startDate", rawText: "2026-02-03" }
      ]
    });
    expect(pending).toMatchObject({ status: "pending" });
    expect(session.getSnapshot().getCell("employee-1", "salary").storedValue).toBe(120);
    expect(session.getSnapshot().getCell("employee-1", "active").storedValue).toBe(false);
    expect(typeof session.getSnapshot().getCell("employee-1", "startDate").storedValue).toBe("number");
    expect(session.getSnapshot().rows[0]).toMatchObject({
      kind: "data",
      original: { salary: 100, active: true, startDate: "2026-01-02" }
    });
    expect(session.getSnapshot().pendingOperations).toHaveLength(1);

    const sent = mutate.mock.calls[0][0];
    expect(sent.map((mutation) => mutation.kind === "cell-value" ? mutation.parsedValue : null)).toEqual([
      120,
      false,
      expect.any(Number)
    ]);
    serverRow = {
      ...serverRow,
      salary: 120,
      active: false,
      startDate: "2026-02-03",
      name: "Ada (recalculated)"
    };
    serverRevision = "2";
    acknowledgement.resolve(sent.map((mutation) => ({
      clientMutationId: mutation.clientMutationId,
      status: "committed" as const,
      revision: "2",
      row: serverRow,
      rowVersion: "row-2"
    })));

    await vi.waitFor(() => expect(session.getSnapshot().pendingOperations).toHaveLength(0));
    await vi.waitFor(() => expect(session.getSnapshot().status.phase).toBe("ready"));
    expect(session.getSnapshot().getCell("employee-1", "salary").evaluatedValue).toBe(120);
    expect(session.getSnapshot().getCell("employee-1", "name").displayValue).toBe("Ada (recalculated)");

    session.destroy();
  });

  it("sends complete metadata replacements and blocks formulas when formula capability is absent", async () => {
    let metadata: TableCellMetadata = { comment: "original" };
    const acknowledgement = createDeferredMutation<Employee>();
    const capabilities = {
      ...defaultRemoteCapabilities(),
      pagination: false,
      formula: "none" as const,
      undo: false,
      subscription: false
    } satisfies TableCapabilities;
    const mutate = vi.fn(async (batch) => {
      const mutation = batch[0];
      if (mutation.kind === "cell-metadata") metadata = mutation.metadata;
      return acknowledgement.promise;
    });
    const source = createTestRemoteSource<Employee>({
      capabilities,
      paginationMode: "none",
      undoMode: "none",
      query: async () => unpaginatedResult("1", employees),
      readCell(row, columnId) {
        const value = row[columnId as keyof Employee];
        return { storedValue: value, evaluatedValue: value, metadata };
      },
      mutate,
      compareRevisions: numericRevisionComparator
    });
    const session = createRemoteTableSession({ source, columns });
    session.start();
    await waitUntilReady(session);

    expect(await session.dispatch({
      type: "edit-cells",
      edits: [{ rowId: "employee-1", columnId: "name", rawText: "=A1" }]
    })).toMatchObject({ status: "rejected", reason: "unsupported" });
    expect(mutate).not.toHaveBeenCalled();

    expect(await session.dispatch({
      type: "update-cell-metadata",
      updates: [{
        rowId: "employee-1",
        columnId: "name",
        patch: { comment: undefined, format: { bold: true } }
      }]
    })).toMatchObject({ status: "pending" });
    const sent = mutate.mock.calls[0][0][0];
    expect(sent).toMatchObject({
      kind: "cell-metadata",
      metadata: { format: { bold: true } }
    });
    expect(sent).not.toHaveProperty("metadata.comment");
    expect(session.getSnapshot().getCell("employee-1", "name").metadata)
      .toEqual({ format: { bold: true } });

    acknowledgement.resolve([{
      clientMutationId: sent.clientMutationId,
      status: "committed",
      revision: "2",
      row: employees[0]
    }]);
    await vi.waitFor(() => expect(session.getSnapshot().pendingOperations).toHaveLength(0));
    expect(session.getSnapshot().getCell("employee-1", "name").metadata)
      .toEqual({ format: { bold: true } });

    session.destroy();
  });

  it("preserves the last valid projection for invalid controlled grouping/pagination", () => {
    const query = vi.fn(async (request: QueryRequest) => offsetResult("r1", employees, request));
    const source = createTestRemoteSource<Employee>({ query });
    const session = createRemoteTableSession({ source, columns });
    let publications = 0;
    session.subscribe(() => { publications += 1; });
    session.updateOptions({
      source,
      columns,
      state: {
        grouping: [{ columnId: "department" }],
        pagination: { kind: "offset", offset: 0, limit: 50 }
      }
    });
    session.start();

    expect(query).not.toHaveBeenCalled();
    expect(publications).toBe(1);
    expect(session.getSnapshot()).toMatchObject({
      status: { phase: "error", message: "Grouping requires pagination kind none" },
      issues: [{ code: "TABLE_GROUPING_PAGINATION_CONFLICT" }]
    });

    session.destroy();
  });

  it("forwards exact export scope/query/revision and copies a valid artifact", async () => {
    const bytes = new Uint8Array([1, 2, 3]);
    const exportRemote = vi.fn(async () => ({
      bytes,
      mediaType: "text/csv",
      fileName: "employees.csv"
    }));
    const capabilities = {
      ...defaultRemoteCapabilities(),
      export: { executor: "server", scope: "completeDataset" }
    } satisfies TableCapabilities;
    const source = createTestRemoteSource<Employee>({
      capabilities,
      export: exportRemote,
      query: async (request) => offsetResult("r1", employees, request)
    });
    const session = createRemoteTableSession({ source, columns });
    session.start();
    await waitUntilReady(session);

    const artifact = await session.export({ format: "csv", scope: "completeDataset" });
    expect(exportRemote).toHaveBeenCalledWith(
      expect.objectContaining({
        format: "csv",
        scope: "completeDataset",
        query: expect.objectContaining({ pagination: { kind: "offset", offset: 0, limit: 50 } }),
        revision: "r1"
      }),
      expect.objectContaining({ operationId: expect.any(String), signal: expect.any(Object) })
    );
    expect(artifact).toEqual({ bytes: new Uint8Array([1, 2, 3]), mediaType: "text/csv", fileName: "employees.csv" });
    expect(artifact.bytes).not.toBe(bytes);
    bytes[0] = 9;
    expect(artifact.bytes[0]).toBe(1);

    session.destroy();
  });

  it("rejects malformed export artifacts and unsupported complete-dataset scope", async () => {
    const exportRemote = vi.fn(async () => ({
      bytes: new Uint8Array([1]), mediaType: "", fileName: "wrong.txt"
    }));
    const capabilities = {
      ...loadedRowCapabilities(),
      export: { executor: "server", scope: "loadedRows" }
    } satisfies TableCapabilities;
    const source = createTestRemoteSource<Employee>({
      capabilities,
      export: exportRemote,
      query: async (request) => offsetResult("r1", employees, request),
      undoMode: "none"
    });
    const session = createRemoteTableSession({ source, columns });
    session.start();
    await waitUntilReady(session);

    await expect(session.export({ format: "csv", scope: "completeDataset" }))
      .rejects.toMatchObject({ code: "TABLE_CAPABILITY_UNSUPPORTED" });
    expect(exportRemote).not.toHaveBeenCalled();
    await expect(session.export({ format: "csv", scope: "currentView" }))
      .rejects.toMatchObject({ code: "REMOTE_EXPORT_PROTOCOL_ERROR" });

    session.destroy();
  });

  it("aborts an in-flight export and rejects its late result after destroy", async () => {
    let resolveExport!: (artifact: { bytes: Uint8Array; mediaType: string; fileName: string }) => void;
    let exportSignal: { aborted: boolean } | undefined;
    const exportRemote = vi.fn((_request, context) => {
      exportSignal = context.signal;
      return new Promise<{ bytes: Uint8Array; mediaType: string; fileName: string }>((resolve) => {
        resolveExport = resolve;
      });
    });
    const source = createTestRemoteSource<Employee>({
      capabilities: {
        ...defaultRemoteCapabilities(),
        export: { executor: "server", scope: "completeDataset" }
      },
      export: exportRemote,
      query: async (request) => offsetResult("r1", employees, request)
    });
    const session = createRemoteTableSession({ source, columns });
    session.start();
    await waitUntilReady(session);
    const exported = session.export({ format: "csv", scope: "completeDataset" });

    session.destroy();
    expect(exportSignal?.aborted).toBe(true);
    resolveExport({ bytes: new Uint8Array([1]), mediaType: "text/csv", fileName: "late.csv" });
    await expect(exported).rejects.toMatchObject({ code: "REMOTE_SESSION_DESTROYED" });
  });

  it("rejects duplicate columns before any source work", () => {
    const source = createTestRemoteSource<Employee>();
    expect(() => createRemoteTableSession({ source, columns: [columns[0], columns[0]] }))
      .toThrow(/Duplicate column id/);
  });

  it("stops and restarts safely, then becomes permanently inert after destroy", async () => {
    const query = vi.fn(async (request: QueryRequest) => offsetResult("r1", employees, request));
    const session = createRemoteTableSession({ source: createTestRemoteSource({ query }), columns });
    session.start();
    await waitForCalls(query, 1);
    session.stop();
    session.start();
    await waitForCalls(query, 2);
    session.destroy();
    session.start();
    expect(query).toHaveBeenCalledTimes(2);
    expect(session.getDiagnostics().destroyed).toBe(true);
  });
});

function accessorColumn(
  id: keyof Employee,
  header: string,
  dataType: "text" | "number" | "boolean" | "date" = "text"
): ColumnDef<Employee> {
  return {
    id,
    header,
    dataType,
    accessor: (row) => row[id],
    update: (row, value) => ({ ...row, [id]: value }) as Employee,
    sortable: true,
    filterable: true
  };
}

function offsetResult(
  revision: string,
  rows: readonly Employee[],
  request: QueryRequest
): QueryResult<Employee> {
  const pagination = request.pagination.kind === "offset"
    ? request.pagination
    : { kind: "offset" as const, offset: 0, limit: 50 };
  return {
    items: rows.map((row) => ({ kind: "data" as const, id: row.id, original: row, depth: 0 })),
    revision,
    completeness: "loadedRows",
    pageInfo: {
      ...pagination,
      total: { kind: "known", value: rows.length },
      hasMore: false
    }
  };
}

function unpaginatedResult(
  revision: string,
  rows: readonly Employee[]
): QueryResult<Employee> {
  return {
    items: rows.map((row) => ({ kind: "data" as const, id: row.id, original: row, depth: 0 })),
    revision,
    completeness: "completeDataset",
    pageInfo: { kind: "none", total: { kind: "known", value: rows.length } }
  };
}

function createDeferredMutation<TRow>() {
  let resolve!: (results: readonly RemoteMutationResult<TRow>[]) => void;
  const promise = new Promise<readonly RemoteMutationResult<TRow>[]>((resolvePromise) => {
    resolve = resolvePromise;
  });
  return { promise, resolve };
}

function numericRevisionComparator(candidate: string, current: string) {
  if (candidate === current) return "equal" as const;
  return Number(candidate) > Number(current) ? "newer" as const : "older" as const;
}

function loadedRowCapabilities(): TableCapabilities {
  const loaded = { executor: "server" as const, scope: "loadedRows" as const };
  return {
    sort: { ...loaded }, filter: { ...loaded }, group: false, aggregate: false,
    pagination: { ...loaded, modes: ["offset"] },
    edit: false, bulkEdit: false, metadata: false, validation: false,
    formula: "none", subscription: false, undo: false, export: false
  };
}

async function waitUntilReady(session: ReturnType<typeof createRemoteTableSession<Employee>>): Promise<void> {
  await vi.waitFor(() => expect(session.getSnapshot().status.phase).toBe("ready"));
}

async function waitForCalls(mock: ReturnType<typeof vi.fn>, count: number): Promise<void> {
  await vi.waitFor(() => expect(mock).toHaveBeenCalledTimes(count));
}
