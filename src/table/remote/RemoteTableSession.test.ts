import { describe, expect, it, vi } from "vitest";
import type { TableCapabilities } from "../core/capabilities";
import type { QueryRequest, QueryResult, QueryRow } from "../core/query";
import type { ColumnDef, TableCellMetadata, TableViewState } from "../core/types";
import { createTestRemoteSource, defaultRemoteCapabilities } from "./testUtils";
import { createRemoteTableSession } from "./RemoteTableSession";
import type { RemoteMutation, RemoteMutationResult, RemoteSourceEvent } from "./types";

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

  it("bounds accumulated pages, exposes the gap, and emits one sanitized warning", async () => {
    const secondEmployee = { ...employees[0], id: "employee-2", name: "Grace" };
    const thirdEmployee = { ...employees[0], id: "employee-3", name: "Linus" };
    const responses = [
      infiniteResult("1", employees, "next"),
      infiniteResult("2", [secondEmployee], "last"),
      infiniteResult("3", [thirdEmployee])
    ];
    const query = vi.fn(async () => responses.shift()!);
    const onDiagnostic = vi.fn();
    const capabilities = {
      ...defaultRemoteCapabilities(),
      pagination: {
        executor: "server" as const,
        scope: "completeDataset" as const,
        modes: ["infinite" as const]
      },
      subscription: false
    } satisfies TableCapabilities;
    const source = createTestRemoteSource<Employee>({
      capabilities,
      paginationMode: "infinite",
      compareRevisions: numericRevisionComparator,
      query
    });
    const session = createRemoteTableSession({
      source,
      columns,
      cache: { maxPages: 1 },
      commandIdFactory: (() => {
        let command = 0;
        return () => `command-${++command}`;
      })(),
      onDiagnostic
    });
    session.start();
    await waitUntilReady(session);

    await session.dispatch({
      type: "set-pagination",
      pagination: { kind: "infinite", after: "next", limit: 50 }
    });
    await vi.waitFor(() => expect(session.getSnapshot().revision).toContain("2:"));
    await session.dispatch({
      type: "set-pagination",
      pagination: { kind: "infinite", after: "last", limit: 50 }
    });
    await vi.waitFor(() => expect(session.getSnapshot().revision).toContain("3:"));

    expect(session.getSnapshot().rows.map((row) => row.id)).toEqual(["employee-3"]);
    expect(session.getSnapshot().pageGaps).toEqual([{
      kind: "evicted-pages", at: 0, omittedPages: 2, omittedItems: 2
    }]);
    expect(session.getDiagnostics()).toMatchObject({
      cachedPages: 1,
      cachedGapPages: 2,
      cachedGapItems: 2,
      hasPageGaps: true
    });
    const warnings = onDiagnostic.mock.calls
      .map(([event]) => event)
      .filter((event) => event.metadata.code === "REMOTE_CACHE_WINDOW_GAP");
    expect(warnings).toEqual([expect.objectContaining({
      category: "remote",
      commandId: "command-3",
      metadata: {
        code: "REMOTE_CACHE_WINDOW_GAP",
        outcome: "warning",
        mode: "infinite",
        maxPages: 1,
        cachedPages: 1,
        omittedPages: 1,
        omittedItems: 1
      }
    })]);
    expect(JSON.stringify(warnings)).not.toContain("employee-");
    expect(JSON.stringify(warnings)).not.toContain("next");
    expect(JSON.stringify(warnings)).not.toContain("last");

    session.destroy();
  });

  it("applies cache maxPages changes to a live session", async () => {
    const secondEmployee = { ...employees[0], id: "employee-2", name: "Grace" };
    const responses = [
      infiniteResult("1", employees, "next"),
      infiniteResult("2", employees, "next"),
      infiniteResult("3", [secondEmployee])
    ];
    const query = vi.fn(async () => responses.shift()!);
    const capabilities = {
      ...defaultRemoteCapabilities(),
      pagination: {
        executor: "server" as const,
        scope: "completeDataset" as const,
        modes: ["infinite" as const]
      },
      subscription: false
    } satisfies TableCapabilities;
    const source = createTestRemoteSource<Employee>({
      capabilities,
      paginationMode: "infinite",
      compareRevisions: numericRevisionComparator,
      query
    });
    const session = createRemoteTableSession({ source, columns, cache: { maxPages: 2 } });
    session.start();
    await waitUntilReady(session);

    session.updateOptions({ source, columns, cache: { maxPages: 1 } });
    session.start();
    await waitForCalls(query, 2);
    await vi.waitFor(() => expect(session.getSnapshot().revision).toContain("2:"));
    await session.dispatch({
      type: "set-pagination",
      pagination: { kind: "infinite", after: "next", limit: 50 }
    });
    await vi.waitFor(() => expect(session.getSnapshot().revision).toContain("3:"));

    expect(session.getSnapshot().rows.map((row) => row.id)).toEqual(["employee-2"]);
    expect(session.getSnapshot().pageGaps).toEqual([{
      kind: "evicted-pages", at: 0, omittedPages: 1, omittedItems: 1
    }]);
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

  it("rejects grouping without wedging a paginated group-capable source", async () => {
    const query = vi.fn(async (request: QueryRequest) => offsetResult("r1", employees, request));
    const capabilities = {
      ...defaultRemoteCapabilities(),
      group: { executor: "server", scope: "completeDataset" }
    } satisfies TableCapabilities;
    const source = createTestRemoteSource<Employee>({ query, capabilities });
    const session = createRemoteTableSession({ source, columns });
    session.start();
    await waitUntilReady(session);

    expect(session.getSnapshot().operationStates.group).toMatchObject({ enabled: false });
    expect(await session.dispatch({
      type: "set-grouping",
      grouping: [{ columnId: "department" }]
    })).toMatchObject({ status: "rejected", reason: "unsupported" });
    expect(session.getSnapshot()).toMatchObject({
      status: { phase: "ready" },
      state: {
        grouping: [],
        pagination: { kind: "offset", offset: 0, limit: 50 }
      }
    });

    expect(await session.dispatch({
      type: "set-pagination",
      pagination: { kind: "offset", offset: 0, limit: 25 }
    })).toMatchObject({ status: "committed" });
    await waitForCalls(query, 2);
    session.stop();
    session.start();
    await waitForCalls(query, 3);
    expect(session.getSnapshot().status.phase).toBe("ready");

    session.destroy();
  });

  it.each(["offset", "cursor", "infinite"] as const)(
    "rejects invalid default grouping before a %s query and recovers when grouping is cleared",
    async (mode) => {
      const query = vi.fn(async (request: QueryRequest) => resultForRequest("1", employees, request));
      const source = createTestRemoteSource<Employee>({
        capabilities: capabilitiesForPagination(mode),
        paginationMode: mode,
        query
      });
      const session = createRemoteTableSession({
        source,
        columns,
        defaultState: {
          grouping: [{ columnId: "department" }],
          pagination: { kind: "none" }
        }
      });

      expect(session.getSnapshot()).toMatchObject({
        status: { phase: "error", message: "Grouping is unavailable for paginated remote sources" },
        issues: [{ code: "TABLE_GROUPING_PAGINATION_CONFLICT" }]
      });
      session.start();
      expect(query).not.toHaveBeenCalled();

      await expect(session.dispatch({ type: "set-grouping", grouping: [] })).resolves.toMatchObject({
        status: "committed"
      });
      await waitForCalls(query, 1);
      await waitUntilReady(session);
      expect(session.getSnapshot().state).toMatchObject({
        grouping: [],
        pagination: defaultPaginationForMode(mode)
      });

      session.stop();
      session.start();
      await waitForCalls(query, 2);
      await waitUntilReady(session);
      session.destroy();
    }
  );

  it.each(["offset", "cursor", "infinite"] as const)(
    "rejects invalid controlled grouping during %s option sync and accepts a valid recovery",
    async (mode) => {
      const query = vi.fn(async (request: QueryRequest) => resultForRequest("1", employees, request));
      const source = createTestRemoteSource<Employee>({
        capabilities: capabilitiesForPagination(mode),
        paginationMode: mode,
        query
      });
      const session = createRemoteTableSession({ source, columns });
      session.start();
      await waitUntilReady(session);

      session.updateOptions({
        source,
        columns,
        state: {
          grouping: [{ columnId: "department" }],
          pagination: { kind: "none" }
        }
      });
      expect(session.getSnapshot()).toMatchObject({
        status: { phase: "error", message: "Grouping is unavailable for paginated remote sources" },
        issues: [{ code: "TABLE_GROUPING_PAGINATION_CONFLICT" }]
      });
      expect(query).toHaveBeenCalledTimes(1);

      session.updateOptions({
        source,
        columns,
        state: { grouping: [], pagination: defaultPaginationForMode(mode) }
      });
      session.start();
      await waitForCalls(query, 2);
      await waitUntilReady(session);
      expect(session.getSnapshot().state.grouping).toEqual([]);
      session.destroy();
    }
  );

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
    expect(session.getSnapshot().canUndo).toBe(false);
    expect(session.getDiagnostics().journalEntries).toBe(0);

    session.destroy();
  });

  it("coalesces repeated metadata targets with deterministic last-write-wins semantics", async () => {
    const acknowledgement = createDeferredMutation<Employee>();
    const mutate = vi.fn(async (_batch: readonly RemoteMutation[]) => acknowledgement.promise);
    const capabilities = {
      ...defaultRemoteCapabilities(),
      pagination: false,
      subscription: false,
      undo: false
    } satisfies TableCapabilities;
    const source = createTestRemoteSource<Employee>({
      capabilities,
      paginationMode: "none",
      undoMode: "none",
      query: async () => unpaginatedResult("1", employees),
      readCell(row, columnId) {
        const value = row[columnId as keyof Employee];
        return {
          storedValue: value,
          evaluatedValue: value,
          metadata: { readOnly: false },
          rowVersion: "row-1"
        };
      },
      mutate,
      compareRevisions: numericRevisionComparator
    });
    const session = createRemoteTableSession({ source, columns });
    session.start();
    await waitUntilReady(session);

    expect(await session.dispatch({
      type: "update-cell-metadata",
      updates: [
        {
          rowId: "employee-1", columnId: "name",
          patch: { comment: "first", format: { bold: true } }
        },
        {
          rowId: "employee-1", columnId: "name",
          patch: { comment: "last", validation: { kind: "textLength", min: 1 } }
        },
        {
          rowId: "employee-1", columnId: "name",
          patch: { format: { italic: true }, formula: "=A1", readOnly: true }
        }
      ]
    })).toMatchObject({ status: "pending" });

    const batch = mutate.mock.calls[0][0];
    expect(batch).toHaveLength(1);
    expect(batch[0]).toMatchObject({
      kind: "cell-metadata",
      rowId: "employee-1",
      columnId: "name",
      rowVersion: "row-1",
      metadata: {
        readOnly: true,
        comment: "last",
        format: { italic: true },
        validation: { kind: "textLength", min: 1 },
        formula: "=A1"
      }
    });
    acknowledgement.resolve([{
      clientMutationId: batch[0].clientMutationId,
      status: "committed",
      revision: "2",
      row: employees[0],
      rowVersion: "row-2"
    }]);
    await vi.waitFor(() => expect(session.getSnapshot().pendingOperations).toHaveLength(0));
    session.destroy();
  });

  it("reconciles an uncertain refresh and never lets it mask a later committed edit", async () => {
    let serverRow = { ...employees[0] };
    let revision = "1";
    let rowVersion = "row-1";
    const mutate = vi.fn(async (batch: readonly RemoteMutation[]) => {
      if (mutate.mock.calls.length === 1) throw new Error("connection lost");
      serverRow = { ...serverRow, name: "Marie" };
      revision = "2";
      rowVersion = "row-2";
      return batch.map((mutation) => ({
        clientMutationId: mutation.clientMutationId,
        status: "committed" as const,
        revision,
        row: serverRow,
        rowVersion
      }));
    });
    const capabilities = {
      ...defaultRemoteCapabilities(),
      pagination: false,
      subscription: false,
      undo: false
    } satisfies TableCapabilities;
    const source = createTestRemoteSource<Employee>({
      capabilities,
      paginationMode: "none",
      undoMode: "none",
      query: async () => unpaginatedResult(revision, [serverRow]),
      mutate,
      compareRevisions: numericRevisionComparator,
      readCell(row, columnId) {
        const value = row[columnId as keyof Employee];
        return { storedValue: value, evaluatedValue: value, rowVersion };
      }
    });
    const session = createRemoteTableSession({ source, columns });
    session.start();
    await waitUntilReady(session);

    expect(await session.dispatch({
      type: "edit-cells",
      edits: [{ rowId: "employee-1", columnId: "name", rawText: "Grace" }]
    })).toMatchObject({ status: "pending" });
    await vi.waitFor(() => expect(session.getSnapshot().pendingOperations).toHaveLength(1));

    await session.refresh();
    expect(session.getSnapshot().pendingOperations).toEqual([]);
    expect(session.getSnapshot().conflicts).toEqual([expect.objectContaining({
      attemptedValue: "Grace",
      authoritativeValue: "Ada",
      revision: "1"
    })]);

    expect(await session.dispatch({
      type: "edit-cells",
      edits: [{ rowId: "employee-1", columnId: "name", rawText: "Marie" }]
    })).toMatchObject({ status: "pending" });
    await vi.waitFor(() => expect(session.getSnapshot().pendingOperations).toHaveLength(0));

    expect(session.getSnapshot().getCell("employee-1", "name").storedValue).toBe("Marie");
    expect(session.getSnapshot().conflicts).toEqual([]);
    expect(session.getDiagnostics().pendingMutations).toBe(0);

    session.destroy();
  });

  it("reconciles an uncertain edit from accessor-backed query rows when readCell is absent", async () => {
    let serverRow = { ...employees[0] };
    let revision = "1";
    const mutate = vi.fn(async () => {
      serverRow = { ...serverRow, name: "Grace" };
      revision = "2";
      throw new Error("connection lost after mutation committed");
    });
    const source = createTestRemoteSource<Employee>({
      capabilities: {
        ...defaultRemoteCapabilities(),
        pagination: false,
        metadata: false,
        formula: "none",
        subscription: false,
        undo: false
      },
      paginationMode: "none",
      undoMode: "none",
      readCell: undefined,
      query: async () => unpaginatedResult(revision, [serverRow]),
      mutate,
      compareRevisions: numericRevisionComparator
    });
    const session = createRemoteTableSession({ source, columns });
    session.start();
    await waitUntilReady(session);

    expect(await session.dispatch({
      type: "edit-cells",
      edits: [{ rowId: "employee-1", columnId: "name", rawText: "Grace" }]
    })).toMatchObject({ status: "pending" });
    await vi.waitFor(() => expect(session.getSnapshot().pendingOperations).toHaveLength(1));

    await session.refresh();

    expect(session.getSnapshot().pendingOperations).toEqual([]);
    expect(session.getSnapshot().conflicts).toEqual([]);
    expect(session.getSnapshot().getCell("employee-1", "name").storedValue).toBe("Grace");
    expect(session.getDiagnostics()).toMatchObject({ pendingMutations: 0, journalEntries: 0 });

    session.destroy();
  });

  it("clears pending journal state when a complete refresh proves an uncertain row was deleted", async () => {
    let rows: readonly Employee[] = employees;
    let revision = "1";
    const source = createTestRemoteSource<Employee>({
      capabilities: unpaginatedUndoCapabilities(),
      paginationMode: "none",
      undoMode: "compensating",
      query: async () => unpaginatedResult(revision, rows),
      mutate: async () => { throw new Error("connection lost"); },
      compareRevisions: numericRevisionComparator
    });
    const session = createRemoteTableSession({ source, columns });
    session.start();
    await waitUntilReady(session);

    expect(await session.dispatch({
      type: "edit-cells",
      edits: [{ rowId: "employee-1", columnId: "name", rawText: "Grace" }]
    })).toMatchObject({ status: "pending" });
    await vi.waitFor(() => expect(session.getSnapshot().canUndo).toBe(true));

    rows = [];
    revision = "2";
    await session.refresh();

    expect(session.getSnapshot().pendingOperations).toEqual([]);
    expect(session.getSnapshot().conflicts).toEqual([]);
    expect(session.getSnapshot().canUndo).toBe(false);
    expect(session.getDiagnostics()).toMatchObject({ pendingMutations: 0, journalEntries: 0 });

    session.destroy();
  });

  it("cancels a pending batch, retains tombstone reconciliation, and lets refresh win over a late acknowledgement", async () => {
    const acknowledgement = createDeferredMutation<Employee>();
    const refresh = createDeferredResult<QueryResult<Employee>>();
    let mutationSignal: { aborted: boolean } | undefined;
    const query = vi.fn(async (request: QueryRequest) => query.mock.calls.length === 1
      ? offsetResult("1", employees, request)
      : refresh.promise);
    const mutate = vi.fn(async (_batch: readonly RemoteMutation[], context: { signal: { aborted: boolean } }) => {
      mutationSignal = context.signal;
      return acknowledgement.promise;
    });
    let command = 0;
    const session = createRemoteTableSession({
      source: createTestRemoteSource({ query, mutate, compareRevisions: numericRevisionComparator }),
      columns,
      commandIdFactory: () => `command-${++command}`
    });
    session.start();
    await waitUntilReady(session);

    expect(await session.dispatch({
      type: "edit-cells",
      edits: [
        { rowId: "employee-1", columnId: "name", rawText: "Grace" },
        { rowId: "employee-1", columnId: "salary", rawText: "120" }
      ]
    })).toEqual({ status: "pending", operationId: "command-2" });
    expect(session.getSnapshot().canUndo).toBe(true);
    expect(session.getSnapshot().getCell("employee-1", "name").storedValue).toBe("Grace");
    expect(session.getSnapshot().getCell("employee-1", "salary").storedValue).toBe(120);

    expect(await session.undo()).toEqual({ status: "pending", operationId: "command-3" });
    expect(mutationSignal?.aborted).toBe(true);
    expect(session.getSnapshot().pendingOperations).toHaveLength(0);
    expect(session.getSnapshot().canUndo).toBe(false);
    expect(session.getDiagnostics().journalEntries).toBe(0);
    await waitForCalls(query, 2);

    const sent = mutate.mock.calls[0][0];
    acknowledgement.resolve(sent.map((mutation) => ({
      clientMutationId: mutation.clientMutationId,
      status: "committed" as const,
      revision: "2",
      row: { ...employees[0], name: "Grace", salary: 120 },
      rowVersion: "row-2"
    })));
    await Promise.resolve();
    expect(session.getSnapshot().canUndo).toBe(false);
    expect(session.getDiagnostics().journalEntries).toBe(0);

    refresh.resolve(offsetResult(
      "3",
      [{ ...employees[0], name: "Refreshed", salary: 130 }],
      query.mock.calls[1][0]
    ));
    await waitUntilReady(session);
    expect(session.getSnapshot().getCell("employee-1", "name").storedValue).toBe("Refreshed");
    expect(session.getSnapshot().getCell("employee-1", "salary").storedValue).toBe(130);
    expect(session.getSnapshot().pendingOperations).toHaveLength(0);

    session.destroy();
  });

  it("does not let an abort-ignoring acknowledgement overwrite a completed undo refresh", async () => {
    const acknowledgement = createDeferredMutation<Employee>();
    const refresh = createDeferredResult<QueryResult<Employee>>();
    const compareRevisions = vi.fn(numericRevisionComparator);
    const query = vi.fn(async (request: QueryRequest) => query.mock.calls.length === 1
      ? offsetResult("1", employees, request)
      : refresh.promise);
    const mutate = vi.fn(async (_batch: readonly RemoteMutation[]) => acknowledgement.promise);
    const session = createRemoteTableSession({
      source: createTestRemoteSource({ query, mutate, compareRevisions }),
      columns
    });
    session.start();
    await waitUntilReady(session);

    expect(await session.dispatch({
      type: "edit-cells",
      edits: [{ rowId: "employee-1", columnId: "salary", rawText: "120" }]
    })).toMatchObject({ status: "pending" });
    expect(await session.undo()).toMatchObject({ status: "pending" });
    await waitForCalls(query, 2);

    refresh.resolve(offsetResult("3", [{ ...employees[0], salary: 130 }], query.mock.calls[1][0]));
    await waitUntilReady(session);
    expect(session.getSnapshot().getCell("employee-1", "salary").storedValue).toBe(130);

    const sent = mutate.mock.calls[0][0];
    acknowledgement.resolve(sent.map((mutation) => ({
      clientMutationId: mutation.clientMutationId,
      status: "committed" as const,
      revision: "4",
      row: { ...employees[0], salary: 120 },
      rowVersion: "row-4"
    })));
    await vi.waitFor(() => expect(compareRevisions).toHaveBeenCalledWith("4", "3"));
    expect(session.getSnapshot().getCell("employee-1", "salary").storedValue).toBe(130);
    expect(session.getSnapshot().canUndo).toBe(false);

    session.destroy();
  });

  it("does not treat an older undo refresh as authoritative over a late acknowledgement", async () => {
    const acknowledgement = createDeferredMutation<Employee>();
    const compareRevisions = vi.fn(numericRevisionComparator);
    let queryRevision = "2";
    let queryRows: readonly Employee[] = employees;
    const query = vi.fn(async (request: QueryRequest) =>
      offsetResult(queryRevision, queryRows, request));
    const mutate = vi.fn(async (_batch: readonly RemoteMutation[]) => acknowledgement.promise);
    const session = createRemoteTableSession({
      source: createTestRemoteSource({ query, mutate, compareRevisions }),
      columns
    });
    session.start();
    await waitUntilReady(session);

    expect(await session.dispatch({
      type: "edit-cells",
      edits: [{ rowId: "employee-1", columnId: "salary", rawText: "120" }]
    })).toMatchObject({ status: "pending" });
    queryRevision = "1";
    expect(await session.undo()).toMatchObject({ status: "pending" });
    await vi.waitFor(() => expect(session.getSnapshot()).toMatchObject({
      status: { phase: "error" },
      issues: [{ code: "REMOTE_INVALIDATED" }]
    }));

    const sent = mutate.mock.calls[0][0];
    acknowledgement.resolve(sent.map((mutation) => ({
      clientMutationId: mutation.clientMutationId,
      status: "committed" as const,
      revision: "3",
      row: { ...employees[0], salary: 120 },
      rowVersion: "row-3"
    })));
    await vi.waitFor(() => expect(compareRevisions).toHaveBeenCalledWith("3", "2"));

    queryRevision = "2";
    await session.refresh();
    expect(session.getSnapshot()).toMatchObject({
      status: { phase: "error" },
      issues: [{ code: "REMOTE_INVALIDATED" }]
    });

    queryRevision = "4";
    queryRows = [{ ...employees[0], salary: 100 }];
    await session.refresh();
    expect(session.getSnapshot().status.phase).toBe("ready");
    expect(session.getSnapshot().getCell("employee-1", "salary").storedValue).toBe(100);

    session.destroy();
  });

  it("returns a typed rejection when a custom command ID is reused after acknowledgement", async () => {
    let serverRow = employees[0];
    let serverRevision = "1";
    const mutate = vi.fn(async (batch: readonly RemoteMutation[]) => {
      serverRow = { ...serverRow, name: "Grace" };
      serverRevision = "2";
      return batch.map((mutation) => ({
        clientMutationId: mutation.clientMutationId,
        status: "committed" as const,
        revision: serverRevision,
        row: serverRow,
        rowVersion: "row-2"
      }));
    });
    const session = createRemoteTableSession({
      source: createTestRemoteSource({
        query: async (request) => offsetResult(serverRevision, [serverRow], request),
        mutate,
        compareRevisions: numericRevisionComparator
      }),
      columns,
      commandIdFactory: () => "duplicate-command"
    });
    session.start();
    await waitUntilReady(session);

    expect(await session.dispatch({
      type: "edit-cells",
      edits: [{ rowId: "employee-1", columnId: "name", rawText: "Grace" }]
    })).toEqual({ status: "pending", operationId: "duplicate-command" });
    await vi.waitFor(() => expect(session.getSnapshot().pendingOperations).toHaveLength(0));
    await waitUntilReady(session);
    expect(session.getDiagnostics().journalEntries).toBe(1);

    await expect(session.dispatch({
      type: "edit-cells",
      edits: [{ rowId: "employee-1", columnId: "name", rawText: "Marie" }]
    })).resolves.toMatchObject({
      status: "rejected",
      reason: "validation",
      issues: [{ code: "REMOTE_MUTATION_PROTOCOL_ERROR" }]
    });
    expect(mutate).toHaveBeenCalledTimes(1);

    session.destroy();
  });

  it("undoes an acknowledged value batch with original values and latest row versions", async () => {
    let serverRow = employees[0];
    let serverRevision = "1";
    let rowVersion = "row-1";
    const compensation = createDeferredMutation<Employee>();
    const mutate = vi.fn(async (batch: readonly RemoteMutation[]) => {
      if (mutate.mock.calls.length === 1) {
        serverRow = { ...serverRow, name: "Grace", salary: 120 };
        serverRevision = "3";
        rowVersion = "row-3";
        return batch.map((mutation, index) => ({
          clientMutationId: mutation.clientMutationId,
          status: "committed" as const,
          revision: String(index + 2),
          row: serverRow,
          rowVersion: `row-${index + 2}`
        }));
      }
      return compensation.promise;
    });
    let command = 0;
    const source = createTestRemoteSource<Employee>({
      capabilities: unpaginatedUndoCapabilities(),
      paginationMode: "none",
      query: async () => unpaginatedResult(serverRevision, [serverRow]),
      mutate,
      compareRevisions: numericRevisionComparator,
      readCell(row, columnId) {
        const value = row[columnId as keyof Employee];
        return { storedValue: value, evaluatedValue: value, rowVersion };
      }
    });
    const session = createRemoteTableSession({
      source,
      columns,
      commandIdFactory: () => `command-${++command}`
    });
    session.start();
    await waitUntilReady(session);

    expect(await session.dispatch({
      type: "edit-cells",
      edits: [
        { rowId: "employee-1", columnId: "name", rawText: "Grace" },
        { rowId: "employee-1", columnId: "salary", rawText: "120" }
      ]
    })).toEqual({ status: "pending", operationId: "command-2" });
    await vi.waitFor(() => expect(session.getSnapshot().canUndo).toBe(true));
    expect(session.getDiagnostics().journalEntries).toBe(1);

    serverRevision = "4";
    rowVersion = "row-4";
    await session.refresh();

    expect(await session.undo()).toEqual({ status: "pending", operationId: "command-4" });
    expect(mutate).toHaveBeenCalledTimes(2);
    const compensatingBatch = mutate.mock.calls[1][0];
    expect(compensatingBatch).toEqual([
      expect.objectContaining({
        kind: "cell-value",
        rowId: "employee-1",
        columnId: "name",
        baseRevision: "3",
        rowVersion: "row-3",
        rawText: "Ada",
        parsedValue: "Ada"
      }),
      expect.objectContaining({
        kind: "cell-value",
        rowId: "employee-1",
        columnId: "salary",
        baseRevision: "3",
        rowVersion: "row-3",
        rawText: "100",
        parsedValue: 100
      })
    ]);
    expect(session.getDiagnostics().journalEntries).toBe(1);
    expect(session.getSnapshot().canUndo).toBe(false);

    serverRow = employees[0];
    serverRevision = "5";
    rowVersion = "row-5";
    compensation.resolve(compensatingBatch.map((mutation) => ({
      clientMutationId: mutation.clientMutationId,
      status: "committed" as const,
      revision: serverRevision,
      row: serverRow,
      rowVersion
    })));
    await vi.waitFor(() => expect(session.getDiagnostics().journalEntries).toBe(0));
    expect(session.getSnapshot().getCell("employee-1", "name").storedValue).toBe("Ada");
    expect(session.getSnapshot().getCell("employee-1", "salary").storedValue).toBe(100);

    session.destroy();
  });

  it("undoes acknowledged metadata with complete original metadata objects", async () => {
    const originals: Record<string, TableCellMetadata> = {
      name: { comment: "name-original", format: { italic: true } },
      department: { comment: "department-original", readOnly: true }
    };
    let metadata = structuredClone(originals);
    let serverRevision = "1";
    let rowVersion = "row-1";
    const compensation = createDeferredMutation<Employee>();
    const mutate = vi.fn(async (batch: readonly RemoteMutation[]) => {
      if (mutate.mock.calls.length === 1) {
        for (const mutation of batch) {
          if (mutation.kind === "cell-metadata") metadata[mutation.columnId] = mutation.metadata;
        }
        serverRevision = "2";
        rowVersion = "row-2";
        return batch.map((mutation) => ({
          clientMutationId: mutation.clientMutationId,
          status: "committed" as const,
          revision: serverRevision,
          row: employees[0],
          rowVersion
        }));
      }
      return compensation.promise;
    });
    const source = createTestRemoteSource<Employee>({
      capabilities: unpaginatedUndoCapabilities(),
      paginationMode: "none",
      query: async () => unpaginatedResult(serverRevision, employees),
      mutate,
      compareRevisions: numericRevisionComparator,
      readCell(row, columnId) {
        const value = row[columnId as keyof Employee];
        return {
          storedValue: value,
          evaluatedValue: value,
          metadata: structuredClone(metadata[columnId] ?? {}),
          rowVersion
        };
      }
    });
    const session = createRemoteTableSession({ source, columns });
    session.start();
    await waitUntilReady(session);

    expect(await session.dispatch({
      type: "update-cell-metadata",
      updates: [
        { rowId: "employee-1", columnId: "name", patch: { comment: "changed", format: { bold: true } } },
        { rowId: "employee-1", columnId: "department", patch: { comment: "changed-too", readOnly: false } }
      ]
    })).toMatchObject({ status: "pending" });
    await vi.waitFor(() => expect(session.getSnapshot().canUndo).toBe(true));

    serverRevision = "3";
    rowVersion = "row-3";
    await session.refresh();

    expect(await session.undo()).toMatchObject({ status: "pending" });
    const compensatingBatch = mutate.mock.calls[1][0];
    expect(compensatingBatch).toEqual([
      expect.objectContaining({
        kind: "cell-metadata",
        columnId: "name",
        baseRevision: "2",
        rowVersion: "row-2",
        metadata: originals.name
      }),
      expect.objectContaining({
        kind: "cell-metadata",
        columnId: "department",
        baseRevision: "2",
        rowVersion: "row-2",
        metadata: originals.department
      })
    ]);
    metadata = structuredClone(originals);
    serverRevision = "3";
    rowVersion = "row-3";
    compensation.resolve(compensatingBatch.map((mutation) => ({
      clientMutationId: mutation.clientMutationId,
      status: "committed" as const,
      revision: serverRevision,
      row: employees[0],
      rowVersion
    })));
    await vi.waitFor(() => expect(session.getDiagnostics().journalEntries).toBe(0));
    expect(session.getSnapshot().getCell("employee-1", "name").metadata).toEqual(originals.name);
    expect(session.getSnapshot().getCell("employee-1", "department").metadata).toEqual(originals.department);

    session.destroy();
  });

  it("completes the original journal entry when authority confirms an uncertain compensation", async () => {
    let serverRow = { ...employees[0] };
    let revision = "1";
    let rowVersion = "row-1";
    const mutate = vi.fn(async (batch: readonly RemoteMutation[]) => {
      if (mutate.mock.calls.length === 1) {
        serverRow = { ...serverRow, salary: 120 };
        revision = "2";
        rowVersion = "row-2";
        return batch.map((mutation) => ({
          clientMutationId: mutation.clientMutationId,
          status: "committed" as const,
          revision,
          row: serverRow,
          rowVersion
        }));
      }
      serverRow = { ...serverRow, salary: 100 };
      revision = "3";
      rowVersion = "row-3";
      throw new Error("connection lost after compensation committed");
    });
    const source = createTestRemoteSource<Employee>({
      capabilities: unpaginatedUndoCapabilities(),
      paginationMode: "none",
      query: async () => unpaginatedResult(revision, [serverRow]),
      mutate,
      compareRevisions: numericRevisionComparator,
      readCell(row, columnId) {
        const value = row[columnId as keyof Employee];
        return { storedValue: value, evaluatedValue: value, rowVersion };
      }
    });
    const session = createRemoteTableSession({ source, columns });
    session.start();
    await waitUntilReady(session);

    expect(await session.dispatch({
      type: "edit-cells",
      edits: [{ rowId: "employee-1", columnId: "salary", rawText: "120" }]
    })).toMatchObject({ status: "pending" });
    await vi.waitFor(() => expect(session.getDiagnostics().journalEntries).toBe(1));

    expect(await session.undo()).toMatchObject({ status: "pending" });
    await vi.waitFor(() => expect(session.getSnapshot().pendingOperations).toHaveLength(1));
    expect(session.getSnapshot().canUndo).toBe(false);
    await session.refresh();

    expect(session.getSnapshot().getCell("employee-1", "salary").storedValue).toBe(100);
    expect(session.getSnapshot().pendingOperations).toEqual([]);
    expect(session.getDiagnostics().journalEntries).toBe(0);
    expect(session.getSnapshot().canUndo).toBe(false);

    session.destroy();
  });

  it.each(["rejected", "conflict", "uncertain"] as const)(
    "keeps %s compensation retryable and never advertises remote redo",
    async (failure) => {
      let serverRow = employees[0];
      let revision = "1";
      const mutate = vi.fn(async (
        batch: readonly RemoteMutation[]
      ): Promise<readonly RemoteMutationResult<Employee>[]> => {
        if (mutate.mock.calls.length === 1) {
          serverRow = { ...serverRow, salary: 120 };
          revision = "2";
          return batch.map((mutation) => ({
            clientMutationId: mutation.clientMutationId,
            status: "committed" as const,
            revision,
            row: serverRow,
            rowVersion: "row-2"
          }));
        }
        if (failure === "uncertain") throw new Error("network outcome unknown");
        if (failure === "conflict" && mutate.mock.calls.length === 2) {
          return batch.map((mutation) => ({
            clientMutationId: mutation.clientMutationId,
            status: "conflict" as const,
            revision: "3",
            current: { ...serverRow, salary: 125 },
            rowVersion: "row-3"
          }));
        }
        if (failure === "conflict") {
          serverRow = employees[0];
          revision = "4";
          return batch.map((mutation) => ({
            clientMutationId: mutation.clientMutationId,
            status: "committed" as const,
            revision,
            row: serverRow,
            rowVersion: "row-4"
          }));
        }
        return batch.map((mutation) => ({
          clientMutationId: mutation.clientMutationId,
          status: "rejected" as const,
          revision,
          issues: [{ code: "UNDO_REJECTED", message: "Retry later" }]
        }));
      });
      const source = createTestRemoteSource<Employee>({
        capabilities: unpaginatedUndoCapabilities(),
        paginationMode: "none",
        query: async () => unpaginatedResult(revision, [serverRow]),
        mutate,
        compareRevisions: numericRevisionComparator,
        readCell(row, columnId) {
          const value = row[columnId as keyof Employee];
          return { storedValue: value, evaluatedValue: value, rowVersion: revision === "1" ? "row-1" : "row-2" };
        }
      });
      const session = createRemoteTableSession({ source, columns });
      session.start();
      await waitUntilReady(session);

      expect(await session.undo()).toMatchObject({ status: "rejected", reason: "unsupported" });
      expect(await session.redo()).toMatchObject({ status: "rejected", reason: "unsupported" });
      expect(await session.dispatch({
        type: "edit-cells",
        edits: [{ rowId: "employee-1", columnId: "salary", rawText: "120" }]
      })).toMatchObject({ status: "pending" });
      await vi.waitFor(() => expect(session.getSnapshot().canUndo).toBe(true));

      expect(await session.undo()).toMatchObject({ status: "pending" });
      if (failure === "uncertain") {
        await vi.waitFor(() => expect(session.getSnapshot().pendingOperations).toHaveLength(1));
        expect(session.getSnapshot().canUndo).toBe(false);
      } else {
        await vi.waitFor(() => expect(session.getSnapshot().canUndo).toBe(true));
      }
      expect(session.getDiagnostics().journalEntries).toBe(1);
      if (failure === "conflict") {
        const conflict = session.getSnapshot().conflicts[0];
        expect(conflict).toBeDefined();
        expect(await session.dispatch({
          type: "retry-with-revision",
          operationId: conflict.operationId,
          rowId: conflict.rowId,
          expectedRevision: conflict.revision
        })).toMatchObject({ status: "pending" });
        expect(mutate.mock.calls[2][0][0]).toMatchObject({
          baseRevision: "3",
          rowVersion: "row-3"
        });
        await vi.waitFor(() => expect(session.getDiagnostics().journalEntries).toBe(0));
        expect(session.getSnapshot().canUndo).toBe(false);
      }
      if (failure === "uncertain") expect(session.getSnapshot().pendingOperations).toHaveLength(1);
      expect(session.getSnapshot().canRedo).toBe(false);
      expect(await session.redo()).toMatchObject({ status: "rejected", reason: "unsupported" });

      session.destroy();
    }
  );

  it("starts and stops the declared remote subscription with the session lifecycle", async () => {
    let listener: ((event: RemoteSourceEvent<Employee>) => void) | undefined;
    const unsubscribe = vi.fn();
    const subscribe = vi.fn((next: (event: RemoteSourceEvent<Employee>) => void) => {
      listener = next;
      return unsubscribe;
    });
    const capabilities = {
      ...defaultRemoteCapabilities(),
      pagination: false,
      subscription: true,
      undo: false
    } satisfies TableCapabilities;
    const source = createTestRemoteSource<Employee>({
      capabilities,
      paginationMode: "none",
      undoMode: "none",
      query: async () => unpaginatedResult("1", employees),
      subscribe,
      compareRevisions: numericRevisionComparator
    });
    const session = createRemoteTableSession({ source, columns });

    session.start();
    await waitUntilReady(session);
    expect(subscribe).toHaveBeenCalledTimes(1);
    listener?.({
      kind: "rows-upserted",
      revision: "2",
      rows: [{ ...employees[0], name: "Grace" }]
    });
    await vi.waitFor(() => {
      expect(session.getSnapshot().getCell("employee-1", "name").displayValue).toBe("Grace");
    });

    session.stop();
    expect(unsubscribe).toHaveBeenCalledTimes(1);
    session.destroy();
  });

  it("reloads the authoritative row and retries a retained conflict only at its current revision", async () => {
    const reloadHarness = await createConflictHarness();
    const reloadConflict = reloadHarness.session.getSnapshot().conflicts[0];
    expect(reloadConflict).toMatchObject({
      operationId: "command-2",
      rowId: "employee-1",
      columnId: "salary",
      attemptedValue: 120,
      authoritativeValue: 115,
      revision: "2"
    });

    expect(await reloadHarness.session.dispatch({
      type: "reload-authoritative",
      operationId: reloadConflict.operationId,
      rowId: reloadConflict.rowId
    })).toMatchObject({ status: "pending" });
    expect(reloadHarness.session.getSnapshot().conflicts).toHaveLength(0);
    await waitForCalls(reloadHarness.query, 2);
    reloadHarness.session.destroy();

    const retryHarness = await createConflictHarness();
    const retryConflict = retryHarness.session.getSnapshot().conflicts[0];
    expect(await retryHarness.session.dispatch({
      type: "retry-with-revision",
      operationId: retryConflict.operationId,
      rowId: retryConflict.rowId,
      expectedRevision: retryConflict.revision
    })).toMatchObject({ status: "pending", operationId: "command-3" });
    expect(retryHarness.mutate).toHaveBeenCalledTimes(2);
    const first = retryHarness.mutate.mock.calls[0][0][0];
    const retried = retryHarness.mutate.mock.calls[1][0][0];
    expect(retried).toMatchObject({
      kind: "cell-value",
      rowId: "employee-1",
      columnId: "salary",
      rawText: "120",
      parsedValue: 120,
      baseRevision: "2"
    });
    expect(retried.clientMutationId).not.toBe(first.clientMutationId);
    expect(retryHarness.session.getSnapshot().conflicts).toHaveLength(0);
    retryHarness.session.destroy();
  });

  it("rejects stale conflict reconciliation without losing the retained conflict", async () => {
    const harness = await createConflictHarness();
    const conflict = harness.session.getSnapshot().conflicts[0];
    const calls = harness.mutate.mock.calls.length;
    const before = harness.session.getSnapshot().conflicts;

    for (const intent of [
      { type: "reload-authoritative", operationId: "", rowId: conflict.rowId } as const,
      { type: "reload-authoritative", operationId: conflict.operationId, rowId: "missing" } as const,
      {
        type: "retry-with-revision",
        operationId: conflict.operationId,
        rowId: conflict.rowId,
        expectedRevision: "1"
      } as const
    ]) {
      expect(await harness.session.dispatch(intent)).toMatchObject({
        status: "rejected",
        reason: "validation",
        issues: [{ code: "REMOTE_CONFLICT_NOT_CURRENT" }]
      });
    }
    expect(harness.mutate).toHaveBeenCalledTimes(calls);
    expect(harness.session.getSnapshot().conflicts).toEqual(before);
    harness.session.destroy();
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
      status: { phase: "error", message: "Grouping is unavailable for paginated remote sources" },
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

function infiniteResult(
  revision: string,
  rows: readonly Employee[],
  nextCursor?: string
): QueryResult<Employee> {
  return {
    items: rows.map((row) => ({ kind: "data" as const, id: row.id, original: row, depth: 0 })),
    revision,
    completeness: "loadedRows",
    pageInfo: {
      kind: "infinite",
      ...(nextCursor === undefined ? {} : { nextCursor }),
      loadedCount: rows.length,
      total: { kind: "unknown" }
    }
  };
}

function resultForRequest(
  revision: string,
  rows: readonly Employee[],
  request: QueryRequest
): QueryResult<Employee> {
  switch (request.pagination.kind) {
    case "none":
      return unpaginatedResult(revision, rows);
    case "offset":
      return offsetResult(revision, rows, request);
    case "cursor":
      return {
        items: rows.map((row) => ({ kind: "data" as const, id: row.id, original: row, depth: 0 })),
        revision,
        completeness: "loadedRows",
        pageInfo: { kind: "cursor", total: { kind: "unknown" } }
      };
    case "infinite":
      return infiniteResult(revision, rows);
  }
}

function capabilitiesForPagination(
  mode: "offset" | "cursor" | "infinite"
): TableCapabilities {
  const complete = { executor: "server" as const, scope: "completeDataset" as const };
  return {
    ...defaultRemoteCapabilities(),
    group: { ...complete },
    pagination: { ...complete, modes: [mode] },
    subscription: false
  };
}

function defaultPaginationForMode(
  mode: "offset" | "cursor" | "infinite"
): TableViewState["pagination"] {
  switch (mode) {
    case "offset": return { kind: "offset", offset: 0, limit: 50 };
    case "cursor": return { kind: "cursor", limit: 50 };
    case "infinite": return { kind: "infinite", limit: 50 };
  }
}

function createDeferredMutation<TRow>() {
  let resolve!: (results: readonly RemoteMutationResult<TRow>[]) => void;
  const promise = new Promise<readonly RemoteMutationResult<TRow>[]>((resolvePromise) => {
    resolve = resolvePromise;
  });
  return { promise, resolve };
}

function createDeferredResult<T>() {
  let resolve!: (value: T) => void;
  const promise = new Promise<T>((resolvePromise) => {
    resolve = resolvePromise;
  });
  return { promise, resolve };
}

async function createConflictHarness() {
  let queryRevision = "1";
  let queryRow = employees[0];
  const query = vi.fn(async () => unpaginatedResult(queryRevision, [queryRow]));
  const mutate = vi.fn(async (batch: readonly RemoteMutation[]) => {
    if (mutate.mock.calls.length === 1) {
      queryRevision = "2";
      queryRow = { ...employees[0], salary: 115 };
      return batch.map((mutation) => ({
        clientMutationId: mutation.clientMutationId,
        status: "conflict" as const,
        revision: "2",
        current: queryRow
      }));
    }
    return new Promise<readonly RemoteMutationResult<Employee>[]>(() => undefined);
  });
  const capabilities = {
    ...defaultRemoteCapabilities(),
    pagination: false,
    subscription: false,
    undo: false
  } satisfies TableCapabilities;
  let command = 0;
  const source = createTestRemoteSource<Employee>({
    capabilities,
    paginationMode: "none",
    undoMode: "none",
    query,
    mutate,
    compareRevisions: numericRevisionComparator
  });
  const session = createRemoteTableSession({
    source,
    columns,
    commandIdFactory: () => `command-${++command}`
  });
  session.start();
  await waitUntilReady(session);
  expect(await session.dispatch({
    type: "edit-cells",
    edits: [{ rowId: "employee-1", columnId: "salary", rawText: "120" }]
  })).toMatchObject({ status: "pending", operationId: "command-2" });
  await vi.waitFor(() => expect(session.getSnapshot().conflicts).toHaveLength(1));
  return { session, query, mutate };
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

function unpaginatedUndoCapabilities(): TableCapabilities {
  return {
    ...defaultRemoteCapabilities(),
    pagination: false,
    subscription: false
  };
}

async function waitUntilReady(session: ReturnType<typeof createRemoteTableSession<Employee>>): Promise<void> {
  await vi.waitFor(() => expect(session.getSnapshot().status.phase).toBe("ready"));
}

async function waitForCalls(mock: ReturnType<typeof vi.fn>, count: number): Promise<void> {
  await vi.waitFor(() => expect(mock).toHaveBeenCalledTimes(count));
}
