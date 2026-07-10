# Remote DataTable Source Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add a race-safe, host-authoritative remote source for `DataTable` with explicit capabilities, query scope, optimistic overlays, mutation acknowledgements, conflicts, subscriptions, and version-checked compensating undo.

**Architecture:** `RemoteTableSession` implements the `TableSession` contract from the local-table plan while keeping its page cache disposable. `RemoteQueryController` owns abort and generation logic; `RemoteMutationController` owns optimistic overlays and server results; the host API remains authoritative.

**Tech Stack:** TypeScript 6, React 19, Vitest 4, Testing Library, Playwright, fast-check, pnpm 11.

## Global Constraints

- Requires the completed local-table plan and its `src/table/core` contracts.
- Remote page caches and optimistic overlays are never authoritative row stores.
- Every query is abortable and generation checked.
- Every mutation carries a client mutation ID plus source or row revision.
- Hosts must process repeated `clientMutationId` values idempotently; uncertain retries reuse IDs.
- Loaded-row operations are never silently presented as complete-dataset operations.
- Server validation and conflict results remain authoritative.
- Remote undo is enabled only when the source advertises version-checked compensation.
- All async work stops publishing after session destruction.

---

### Task 1: Define the remote source and mutation result contracts

**Files:**
- Create: `src/table/remote/types.ts`
- Create: `src/table/remote/createRemoteTableSource.ts`
- Test: `src/table/remote/createRemoteTableSource.test.ts`

**Interfaces:**
- Consumes: `QueryRequest`, `QueryResult`, `ExportOptions`, `ExportArtifact`, `TableAbortSignal`, and `TableCapabilities` from `src/table/core`.
- Produces: `RemoteTableSource<TRow>`, `RemoteMutation`, `RemoteMutationResult<TRow>`, `createRemoteTableSource`.

- [ ] **Step 1: Write the failing source contract test**

```ts
import { describe, expect, it, vi } from "vitest";
import { createRemoteTableSource } from "./createRemoteTableSource";

describe("createRemoteTableSource", () => {
  it("requires stable ids and preserves declared dataset scope", async () => {
    const query = vi.fn(async () => ({
      items: [{ kind: "data" as const, id: "employee-1", original: { id: "employee-1", name: "Ada" }, depth: 0 }],
      revision: "r1",
      completeness: "loadedRows" as const,
      pageInfo: { kind: "offset" as const, offset: 0, limit: 50, total: { kind: "known" as const, value: 1 }, hasMore: false }
    }));
    const source = createRemoteTableSource({
      getRowId: (row: { id: string }) => row.id,
      capabilities: {
        sort: { executor: "server", scope: "completeDataset" },
        filter: { executor: "server", scope: "completeDataset" },
        group: false,
        aggregate: false,
        pagination: { executor: "server", scope: "completeDataset", modes: ["offset"] },
        edit: { executor: "server", scope: "completeDataset" },
        bulkEdit: false,
        metadata: false,
        validation: { executor: "server", scope: "completeDataset" },
        undo: false,
        export: false,
        formula: "none",
        subscription: false
      },
      paginationMode: "offset",
      mutationMode: "versioned",
      undoMode: "none",
      compareRevisions: (candidate, current) => candidate === current ? "equal" : "unknown",
      query
    });

    const result = await source.query(
      { sorting: [], filter: null, grouping: [], aggregates: [], pagination: { kind: "offset", offset: 0, limit: 50 } },
      { signal: new AbortController().signal, generation: 1, operationId: "query-test-1" }
    );

    expect(source.kind).toBe("remote");
    const first = result.items[0];
    expect(first.kind).toBe("data");
    if (first.kind !== "data") throw new Error("expected data row");
    expect(source.getRowId(first.original)).toBe("employee-1");
    expect(result.completeness).toBe("loadedRows");
    expect(source.capabilities.sort).toEqual({ executor: "server", scope: "completeDataset" });
  });
});
```

- [ ] **Step 2: Run the test and verify the missing module failure**

Run: `corepack pnpm exec vitest run src/table/remote/createRemoteTableSource.test.ts`

Expected: FAIL because `createRemoteTableSource` does not exist.

- [ ] **Step 3: Add the source and result types**

```ts
export type RemoteMutationBase = {
  clientMutationId: string;
  rowId: string;
  columnId: string;
  baseRevision: string;
  rowVersion?: string;
};

export type RemoteMutation = RemoteMutationBase & (
  | {
      kind: "cell-value";
      rawText: string;
      parsedValue: unknown;
      formula?: string;
    }
  | {
      kind: "cell-metadata";
      metadata: TableCellMetadata;
    }
);

export type RemoteCellData = {
  storedValue: unknown;
  evaluatedValue: unknown;
  displayValue?: string;
  formula?: string;
  metadata?: TableCellMetadata;
  rowVersion?: string;
  issues?: readonly { code: string; message: string; columnId?: string }[];
};

export type RemoteExportRequest = ExportOptions & {
  query: QueryRequest;
  revision: string;
};

export type RemoteMutationResult<TRow> =
  | { clientMutationId: string; status: "committed"; revision: string; row: TRow; rowVersion?: string }
  | { clientMutationId: string; status: "corrected"; revision: string; row: TRow; rowVersion?: string }
  | { clientMutationId: string; status: "rejected"; revision: string; issues: readonly { code: string; message: string; columnId?: string }[] }
  | { clientMutationId: string; status: "conflict"; revision: string; current: TRow; rowVersion?: string };

export type RemoteSourceEvent<TRow> =
  | { kind: "rows-upserted"; revision: string; rows: readonly TRow[] }
  | { kind: "rows-deleted"; revision: string; rowIds: readonly string[] }
  | { kind: "invalidate"; revision?: string };

export type RevisionComparison = "older" | "equal" | "newer" | "unknown";

export type RemoteTableSource<TRow> = {
  kind: "remote";
  getRowId(row: TRow): string;
  capabilities: TableCapabilities;
  paginationMode: "none" | "offset" | "cursor" | "infinite";
  mutationMode: "none" | "versioned";
  undoMode: "none" | "compensating";
  compareRevisions(candidate: string, current: string): RevisionComparison;
  readCell?(row: TRow, columnId: string): RemoteCellData;
  query(request: QueryRequest, context: { signal: TableAbortSignal; generation: number; operationId: string }): Promise<QueryResult<TRow>>;
  mutate?(batch: readonly RemoteMutation[], context: { signal: TableAbortSignal; operationId: string }): Promise<readonly RemoteMutationResult<TRow>[]>;
  export?(request: RemoteExportRequest, context: { signal: TableAbortSignal; operationId: string }): Promise<ExportArtifact>;
  subscribe?(listener: (event: RemoteSourceEvent<TRow>) => void): () => void;
};
```

Implement `createRemoteTableSource` as an identity-preserving typed factory that rejects blank or duplicate `QueryRow.id` values and verifies each data item's ID equals `getRowId(item.original)` when results are normalized. Reject contradictory declarations: pagination mode absent from `capabilities.pagination.modes`; edit/bulk-edit/metadata enabled without `mutate`; metadata capability without `readCell`; `mutationMode: "none"` with write capabilities; server validation or `formula: "server"` without `mutate`; server formulas without `readCell`; export capability without `export`; subscription enabled without `subscribe`; or compensating undo enabled without versioned mutations, mutate, and an undo capability. Formula, grouping, and aggregation declarations remain independent and are returned unchanged. Revision strings are opaque: the session must call `compareRevisions`; an `unknown` comparison invalidates and requeries instead of guessing lexical or numeric order.

Extend the failing tests with every contradictory declaration, an empty ID, a duplicate ID across one result page, an offset/cursor/infinite result whose `pageInfo.kind` does not match `paginationMode`, and a paginated response that falsely claims `completeness: "completeDataset"`.

- [ ] **Step 4: Run the focused test**

Run: `corepack pnpm exec vitest run src/table/remote/createRemoteTableSource.test.ts`

Expected: PASS.

- [ ] **Step 5: Commit the source contract**

```bash
git add src/table/remote/types.ts src/table/remote/createRemoteTableSource.ts src/table/remote/createRemoteTableSource.test.ts
git commit -m "feat: define remote table source contract"
```

### Task 2: Implement abortable, generation-safe remote queries

**Files:**
- Create: `src/table/remote/RemotePageCache.ts`
- Test: `src/table/remote/RemotePageCache.test.ts`
- Create: `src/table/remote/RemoteQueryController.ts`
- Test: `src/table/remote/RemoteQueryController.test.ts`
- Create: `src/table/remote/testUtils.ts`

**Interfaces:**
- Consumes: `RemoteTableSource<TRow>` and `QueryRequest`.
- Produces: bounded `RemotePageCache<TRow>` and `RemoteQueryController<TRow>` with `load`, `refresh`, `getSnapshot`, `subscribe`, and `destroy`.

- [ ] **Step 1: Write cache, stale-response, and abort tests**

```ts
type Row = { id: string; name: string };

function createDeferred<T>() {
  let resolve!: (value: T) => void;
  let reject!: (reason?: unknown) => void;
  const promise = new Promise<T>((resolvePromise, rejectPromise) => {
    resolve = resolvePromise;
    reject = rejectPromise;
  });
  return { promise, resolve, reject };
}

function queryWithFilter(name: string): QueryRequest {
  return {
    sorting: [],
    filter: {
      kind: "comparison",
      columnId: "name",
      operator: "eq",
      value: { type: "string", value: name }
    },
    grouping: [],
    aggregates: [],
    pagination: { kind: "offset", offset: 0, limit: 50 }
  };
}

function result(revision: string, rows: readonly Row[]): QueryResult<Row> {
  return {
    items: rows.map((row) => ({ kind: "data" as const, id: row.id, original: row, depth: 0 })),
    revision,
    completeness: "loadedRows",
    pageInfo: {
      kind: "offset",
      offset: 0,
      limit: 50,
      total: { kind: "known", value: rows.length },
      hasMore: false
    }
  };
}

it("ignores an older response after a newer query wins", async () => {
  const first = createDeferred<QueryResult<Row>>();
  const second = createDeferred<QueryResult<Row>>();
  const source = createTestRemoteSource<Row>({
    query: vi.fn()
      .mockReturnValueOnce(first.promise)
      .mockReturnValueOnce(second.promise)
  });
  const controller = new RemoteQueryController(source);

  const firstLoad = controller.load(queryWithFilter("Ada"), "query-1");
  const secondLoad = controller.load(queryWithFilter("Grace"), "query-2");
  second.resolve(result("r2", [{ id: "2", name: "Grace" }]));
  await secondLoad;
  first.resolve(result("r1", [{ id: "1", name: "Ada" }]));
  await firstLoad;

  expect(controller.getSnapshot().items).toEqual([
    { kind: "data", id: "2", original: { id: "2", name: "Grace" }, depth: 0 }
  ]);
  expect(controller.getSnapshot().revision).toBe("r2");
});

it("aborts the previous request when query state changes", async () => {
  const signals: AbortSignal[] = [];
  const source = createTestRemoteSource<Row>({
    query: vi.fn((_request, context) => {
      signals.push(context.signal);
      return new Promise<QueryResult<Row>>(() => undefined);
    })
  });
  const controller = new RemoteQueryController(source);
  void controller.load(queryWithFilter("first"), "query-1");
  void controller.load(queryWithFilter("second"), "query-2");
  expect(signals[0].aborted).toBe(true);
});
```

Create `src/table/remote/testUtils.ts` in this task with `createTestRemoteSource`, a complete valid default capability object, a default opaque-revision comparator, and override parameters for `query`, `mutate`, and `subscribe`. Add that file to the task commit; later remote tests must reuse it instead of repeating partial sources.

In `RemotePageCache.test.ts`, cover offset, cursor, and infinite page keys; LRU promotion; replacement of an existing page; aggregation of pages in query order; duplicate item IDs across combined pages; and eviction at the configurable default of five pages. A page cache stores only accepted canonical query results, never optimistic overlays. Duplicate visible IDs are a protocol error and invalidate that query cache rather than rendering duplicate React/grid identities. Also assert `load`/`refresh` after destroy reject with `REMOTE_SESSION_DESTROYED`, make no source call, and publish nothing.

- [ ] **Step 2: Run the tests and verify failure**

Run: `corepack pnpm exec vitest run src/table/remote/RemotePageCache.test.ts src/table/remote/RemoteQueryController.test.ts`

Expected: FAIL because the controller is missing.

- [ ] **Step 3: Implement generation and abort ownership**

```ts
export type RemoteQuerySnapshot<TRow> = {
  status: "idle" | "loading" | "ready" | "error";
  items: readonly QueryRow<TRow>[];
  revision: string | null;
  completeness: QueryResult<TRow>["completeness"];
  pageInfo: QueryResult<TRow>["pageInfo"];
  error?: { code: string; message: string; retryable: boolean };
};

export class RemoteTableError extends Error {
  constructor(readonly code: "NO_QUERY" | "REMOTE_SESSION_DESTROYED", message = code) {
    super(message);
    this.name = "RemoteTableError";
  }
}

export class RemoteQueryController<TRow> {
  private generation = 0;
  private abortController: AbortController | null = null;
  private snapshot: RemoteQuerySnapshot<TRow> = {
    status: "idle",
    items: [],
    revision: null,
    completeness: "loadedRows",
    pageInfo: { kind: "none", total: { kind: "known", value: 0 } }
  };
  private listeners = new Set<() => void>();
  private destroyed = false;

  constructor(private readonly source: RemoteTableSource<TRow>) {}

  async load(query: QueryRequest, operationId: string): Promise<void> {
    if (this.destroyed) throw new RemoteTableError("REMOTE_SESSION_DESTROYED");
    const generation = ++this.generation;
    this.abortController?.abort();
    const abortController = new AbortController();
    this.abortController = abortController;
    this.publish({ ...this.snapshot, status: "loading", error: undefined });
    try {
      const result = await this.source.query(query, { signal: abortController.signal, generation, operationId });
      if (generation !== this.generation || abortController.signal.aborted) return;
      if (this.snapshot.revision) {
        const order = this.source.compareRevisions(result.revision, this.snapshot.revision);
        if (order === "older") return;
        if (order === "unknown") {
          this.publish({ ...this.snapshot, status: "error", error: {
            code: "REVISION_ORDER_UNKNOWN",
            message: "The source could not order this response; refresh required.",
            retryable: true
          }});
          return;
        }
      }
      this.publish({ status: "ready", ...result, error: undefined });
    } catch (error) {
      if (generation !== this.generation || abortController.signal.aborted) return;
      this.publish({ ...this.snapshot, status: "error", error: normalizeRemoteError(error) });
    }
  }

  destroy(): void {
    if (this.destroyed) return;
    this.destroyed = true;
    this.generation += 1;
    this.abortController?.abort();
    this.listeners.clear();
  }

  getSnapshot = (): RemoteQuerySnapshot<TRow> => this.snapshot;

  subscribe = (listener: () => void): (() => void) => {
    this.listeners.add(listener);
    return () => this.listeners.delete(listener);
  };

  refresh(operationId: string): Promise<void> {
    return this.load(this.lastQuery, operationId);
  }

  private publish(next: RemoteQuerySnapshot<TRow>): void {
    this.snapshot = next;
    this.listeners.forEach((listener) => listener());
  }
}
```

The concrete implementation stores `lastQuery`, rejects `refresh()` with a typed `NO_QUERY` error before the first load, calls a `normalizeRemoteError(error)` function defined in the same module, and routes accepted pages through `RemotePageCache`. It also exposes internal adapter methods `getCanonicalRow(rowId)`, `applyCanonicalRows(rows, revision)`, `deleteCanonicalRows(rowIds, revision)`, `noteCurrentRevision(revision)`, and `invalidate()` for the mutation/subscription controllers; none are public package exports. Every public method checks `destroyed` before allocating an AbortController or calling the host. Accept the first response after generation checking. For every later response, call `source.compareRevisions(result.revision, currentRevision)` even when no subscription arrived: accept `newer`/`equal`, ignore `older`, and invalidate with a retryable error on `unknown`. The source adapter must map opaque ETags to an orderable server version or implement a comparator from authoritative version metadata; the runtime never guesses. Offset replacement replaces the visible page; cursor/infinite append combines pages in stable page order. Validate that request pagination matches `source.paginationMode`; invalid variants return `unsupported` without calling the host.

- [ ] **Step 4: Run query tests**

Run: `corepack pnpm exec vitest run src/table/remote/RemotePageCache.test.ts src/table/remote/RemoteQueryController.test.ts`

Expected: PASS, including stale-response and abort cases.

- [ ] **Step 5: Commit query coordination**

```bash
git add src/table/remote/RemotePageCache.ts src/table/remote/RemotePageCache.test.ts src/table/remote/RemoteQueryController.ts src/table/remote/RemoteQueryController.test.ts src/table/remote/testUtils.ts
git commit -m "feat: coordinate abortable remote table queries"
```

### Task 3: Build the remote session read path and capability gating

**Files:**
- Create: `src/table/remote/RemoteTableSession.ts`
- Test: `src/table/remote/RemoteTableSession.test.ts`
- Test: `src/table/remote/RemoteTableSession.contract.test.ts`
- Create: `src/table/remote/index.ts`
- Modify: `src/table/core/capabilities.ts`
- Modify: `src/react/useTableSession.ts`
- Create: `src/react/useTableSession.remote.test.tsx`

**Interfaces:**
- Consumes: `TableSession`, `TableViewSnapshot`, `RemoteQueryController`.
- Produces: `RemoteTableSessionOptions`, `createRemoteTableSession(options)`, and a `TableSession<TRow>` snapshot consumable by `DataTable`.

- [ ] **Step 1: Write failing session tests**

Test that changing controlled sort/filter state generates the exact serializable query AST. A complete-dataset source exposes `scopeLabel: "Complete dataset"`; a separate loaded-row fixture exposes `scopeLabel: "Loaded rows"` and disables configuration that requires complete-dataset scope.

Also cover duplicate columns, lazy accessor-backed cells, `source.readCell` formula/evaluation metadata, group/aggregate cells, typed number/boolean/date snapshots, controlled and uncontrolled query-state slices, unsupported formula scope, one publication per accepted query transition, and export scope/protocol validation. Mutation parsing, optimistic transitions, authoritative rejection, conflict resolution, and undo are deliberately owned by Tasks 4–6 and are not part of this green gate.

In `RemoteTableSession.contract.test.ts`, import `defineTableSessionContract` from `src/table/core/session.contract.ts` and register a deterministic read-only remote-mock harness. Supply feature operations for server sort/filter/group/aggregate/pagination/export only. Set edit, bulk edit, metadata, validation, and undo capabilities to false in this Task 3 harness; omit `conflictResolution` so the shared contract proves both source-neutral conflict intents reject as unsupported. Assert absent features reject without a host call and advertised read/query/export features do not return `unsupported`. Tasks 4, 5, and 6 extend this same harness only after their owning implementations turn green.

```ts
expect(source.query).toHaveBeenCalledWith(
  expect.objectContaining({
    sorting: [{ columnId: "salary", direction: "desc" }],
    filter: {
      kind: "comparison",
      columnId: "department",
      operator: "eq",
      value: { type: "string", value: "Finance" }
    }
  }),
  expect.objectContaining({ generation: 1, operationId: expect.any(String) })
);
expect(session.getSnapshot().operationStates.sort?.scopeLabel).toBe("Complete dataset");
expect(loadedRowsSession.getSnapshot().operationStates.sort).toMatchObject({
  enabled: false,
  scopeLabel: "Loaded rows",
  reason: "This table only supports sorting loaded rows."
});
```

- [ ] **Step 2: Run and verify failure**

Run: `corepack pnpm exec vitest run src/table/remote/RemoteTableSession.test.ts`

Expected: FAIL because the remote session does not exist.

- [ ] **Step 3: Implement read-session composition**

Define the exact public constructor contract:

```ts
export type RemoteTableSessionOptions<
  TRow,
  TColumn extends ColumnDef<TRow> = ColumnDef<TRow>
> = {
  source: RemoteTableSource<TRow>;
  columns: readonly TColumn[];
  state?: Partial<TableViewState>;
  defaultState?: Partial<TableViewState>;
  onStateChange?: (updater: TableStateUpdater, context: ChangeContext) => void;
  features?: TableFeatureConfiguration;
  cache?: { maxPages?: number };
  mutationLimits?: { maxPendingOperations?: number; maxPendingCells?: number };
  commandIdFactory?: CommandIdFactory;
  onDiagnostic?: (event: TableDiagnosticEvent) => void;
};

export type RemoteTableDiagnostics = {
  queryGeneration: number;
  cachedPages: number;
  cachedItems: number;
  pendingMutations: number;
  conflicts: number;
  journalEntries: number;
  destroyed: boolean;
};

export interface RemoteTableSession<
  TRow,
  TColumn extends ColumnDef<TRow> = ColumnDef<TRow>
> extends TableSession<TRow, TColumn> {
  updateOptions(options: RemoteTableSessionOptions<TRow, TColumn>): void;
  start(): void;
  stop(): void;
  getDiagnostics(): RemoteTableDiagnostics;
}

export function createRemoteTableSession<
  TRow,
  TColumn extends ColumnDef<TRow> = ColumnDef<TRow>
>(options: RemoteTableSessionOptions<TRow, TColumn>): RemoteTableSession<TRow, TColumn>;
```

`getDiagnostics()` returns counters only—never row values, auth material, or host request objects. Tests and the final pressure gate use it to assert the cache, overlay, conflict, journal, and destruction bounds. Construct one command-ID factory per session when absent. Every query, mutation, subscription reconciliation, export, undo, refresh, and retry uses an operation ID consistently in source context, pending state, results, and diagnostics. `onDiagnostic` receives shared `{ category: "remote", commandId, durationMs, metadata }` events for cache eviction, protocol violations, opaque revision invalidation, aborts, and connectivity transitions. Metadata is limited to code, generation, counts, scope, and outcome—never rows, values, raw edits, formulas, auth, URLs, headers, or host errors. Exceptions thrown by that callback are caught and never alter session state.

`createRemoteTableSession` is inert: its constructor and `updateOptions` normalize/store options only and never query, subscribe, publish, or invoke host callbacks. `start()` idempotently begins the initial query and subscription after a committed React effect; `stop()` aborts/unsubscribes without destroying reusable configuration; `destroy()` is final. The remote overload of `useTableSession` calls render-pure `updateOptions`, then `start()`/`stop()` in an effect with StrictMode-safe cleanup. Abandoned concurrent renders therefore produce no network activity. Direct non-React callers must call `start()` explicitly. Add tests for abandoned render, StrictMode start/stop/start, source replacement, and zero requests before commit.

Normalize columns once per option identity and reject blank/duplicate IDs. Data-row `getCell` reads the latest optimistic overlay first, then `source.readCell(row, columnId)` when supplied, otherwise the column accessor/calculator; it formats through the column formatter and merges server issues. Group and aggregate `QueryRow` cells read their renderer-neutral key/aggregate fields without fabricating a `TRow`. A cell is editable only when source edit capability, column update/parser, permission, and current conflict state allow it.

Copy query `completeness` into the table snapshot and set `totalRowCount` to the tagged `pageInfo.total` value without coercing unknown totals to loaded length. `rowCount` is the current projected item count. Preserve offset/cursor/infinite page information exactly for controls and ARIA mapping.

Controlled query-state slices emit their updater; uncontrolled query slices update internally and trigger one generation-safe query. Column layout and selection remain view-local and never call the source. Until their owning tasks land, `edit-cells`, `clear-cells`, `update-cell-metadata`, `undo`, `redo`, `reload-authoritative`, and `retry-with-revision` return `rejected/unsupported` before any source call.

Apply the shared grouping/pagination invariant before serialization: starting grouping clears pagination to `none` in the same state update; a pagination intent while grouped rejects; invalid controlled state preserves the last valid projection and publishes a status issue instead of throwing or calling the source.

Implement `RemoteTableSession` by composing the core table state reducer with `RemoteQueryController`. Build the source-neutral `operationStates` snapshot field with `resolveTableOperationStates`. Define an exhaustive local `featureForRemoteIntent(intent)` mapper for source-touching intents; selection and column-layout intents remain local view state and do not require a source capability. Return `unsupported` before any request when the mapped operation state is disabled, including when `features.<feature>.requiredScope === "completeDataset"` conflicts with a loaded-row source.

```ts
const feature = featureForRemoteIntent(intent);
const operation = feature ? operationStates[feature] : { enabled: true };
if (!operation.enabled) {
  return {
    status: "rejected",
    reason: "unsupported",
    issues: [{ code: "TABLE_CAPABILITY_UNSUPPORTED", message: operation.reason ?? "Unsupported operation" }]
  };
}
```

Query intents serialize sorting, filters, grouping, aggregates, and the active pagination variant without falling back to client execution. Server group and aggregate items map to the shared `QueryRow<TRow>` projection; local regrouping or aggregation is forbidden unless the source explicitly advertises `executor: "client"` with `scope: "completeDataset"` and the query result is complete.

Implement `session.export(options)` by validating `operationStates.export`, requested scope, current query, and current revision, then calling `source.export({ ...options, query, revision }, { signal, operationId })`. A complete-dataset export is rejected when export capability is loaded-row scope. Abort on destroy and validate that the result has `bytes instanceof Uint8Array`, a nonblank `mediaType`, and a nonblank extension-correct `fileName`; reject missing or malformed artifacts as `REMOTE_EXPORT_PROTOCOL_ERROR`. Return a copied `Uint8Array` so later host mutation cannot alter the accepted export. Never create a `Blob` in the remote/core layer and never silently serialize only cached pages. Add focused and shared-contract assertions for exact scope/query/revision forwarding, artifact bytes/media/file name, malformed artifact rejection, abort, and no source call when scope is unsupported.

Extend `useTableSession` with overloads for `LocalRecordTableSessionOptions` and `RemoteTableSessionOptions`. It creates the correct owned session from `options.source.kind`, updates options synchronously but render-purely on rerender, and returns the concrete session type. A committed effect starts remote work; a switch between local and remote source kinds creates the new session and destroys the replaced owned session after StrictMode-safe deferred cleanup. `useTableSession.remote.test.tsx` covers inference, no render-phase requests, initial committed query, controlled rerender, source replacement, StrictMode, abandoned render, final unmount, and proves no host-supplied `DataTable session` is destroyed.

- [ ] **Step 4: Run focused and shared contract tests**

Run: `corepack pnpm exec vitest run src/table/remote/RemoteTableSession.test.ts src/table/remote/RemoteTableSession.contract.test.ts src/react/useTableSession.remote.test.tsx`

Expected: PASS; shared cases run only for advertised features.

- [ ] **Step 5: Commit the read session**

```bash
git add src/table/remote/RemoteTableSession.ts src/table/remote/RemoteTableSession.test.ts src/table/remote/RemoteTableSession.contract.test.ts src/table/remote/index.ts src/table/core/capabilities.ts src/react/useTableSession.ts src/react/useTableSession.remote.test.tsx
git commit -m "feat: add capability-aware remote table session"
```

### Task 4: Add optimistic overlays and mutation reconciliation

**Files:**
- Create: `src/table/remote/OptimisticOverlayStore.ts`
- Create: `src/table/remote/RemoteMutationController.ts`
- Test: `src/table/remote/RemoteMutationController.test.ts`
- Modify: `src/table/remote/RemoteTableSession.ts`
- Modify: `src/table/remote/RemoteTableSession.test.ts`
- Modify: `src/table/remote/RemoteTableSession.contract.test.ts`

**Interfaces:**
- Consumes: parsed `edit-cells` intents and versioned `RemoteMutationResult`.
- Produces: overlay snapshots and mutation lifecycle results.

- [ ] **Step 1: Write the four result-path tests**

Cover typed number/boolean/date edits, local pre-validation, committed, corrected, authoritative server rejection, and conflict results. Assert that an invalid batch makes no source call, cached rows are unchanged while an overlay is pending, and each accepted mutation transition publishes once. Cover complete metadata replacement, disabled metadata, server-formula acknowledgement through `readCell`, and formula rejection when source capability is `none`.

Extend the remote shared-contract harness here to advertise and exercise edit, bulk edit, metadata, and validation. Keep undo disabled until Task 6 and continue omitting `conflictResolution` until Task 5; a conflict can now be retained, but its typed resolution commands are not supported until the next owning task.

```ts
expect(session.getSnapshot().getCell("row-1", "salary").storedValue).toBe(120);
expect(queryController.getCanonicalRow("row-1")?.salary).toBe(100);
expect(session.getSnapshot().pendingOperations).toHaveLength(1);
```

- [ ] **Step 2: Run and verify failure**

Run: `corepack pnpm exec vitest run src/table/remote/RemoteMutationController.test.ts src/table/remote/RemoteTableSession.test.ts src/table/remote/RemoteTableSession.contract.test.ts`

Expected: FAIL because optimistic mutation support is missing.

- [ ] **Step 3: Implement overlay-only optimistic state**

```ts
type OptimisticOverlay = {
  clientMutationId: string;
  rowId: string;
  columnId: string;
  value: unknown;
  baseRevision: string;
  status: "pending" | "uncertain" | "conflict";
};

readCellValue(rowId: string, columnId: string): unknown {
  return this.latestOverlay(rowId, columnId)?.value ?? this.readCanonicalCell(rowId, columnId);
}
```

On `edit-cells`, call `parseCellInput(rawText)`, then the column parser and client validator with raw, parsed, and provisional evaluated candidates. Reject the entire batch locally if any cell fails. Formulas reject when capability is `none`; server formulas send both `formula` and parsed value and rely on `source.readCell` after acknowledgement for authoritative evaluation. Server validation remains authoritative even after the local precheck.

On `update-cell-metadata`, require enabled metadata capability and `source.readCell`, apply each intent patch to the current authoritative `RemoteCellData.metadata`, and send the resulting complete metadata object as a `cell-metadata` mutation through the same versioned controller. Complete replacement semantics make clearing optional fields and compensating undo unambiguous. Reject/hide format, validation, comment, read-only, or formula-metadata tools when metadata capability is false; never create a local metadata document for remote rows. A formula typed as a cell value remains an `edit-cells`/`cell-value` mutation and separately requires formula capability.

The mutation controller receives one operation ID from the session and derives collision-safe client mutation IDs as `${operationId}:${editIndex}`; it calls `source.mutate` with the parent operation ID and matches every result by `clientMutationId`. Compare every acknowledgement revision to the current source revision: apply `newer`/`equal`, ignore and invalidate `older`, and create an explicit retryable revision conflict on `unknown`. A committed or corrected result replaces cached canonical copies of that data row before removing the overlay, but that replacement is not treated as a valid projected row model. If the active query has sorting, a filter, grouping, aggregates, or any pagination, invalidate and refetch because membership, order, groups, totals, and page boundaries may have changed; retain the last good projection during the refresh. Only an existing-row update under an unprojected `{ sorting: [], filter: null, grouping: [], aggregates: [], pagination: { kind: "none" } }` query may publish in place. A rejected result removes only its overlay and publishes its issues. A conflict retains the attempted overlay with conflict state and the authoritative row. Missing, duplicate, or unknown result IDs are a `REMOTE_MUTATION_PROTOCOL_ERROR`: remove affected overlays, invalidate the cache, and require refresh. Aborting a pending mutation removes its overlay and ignores any late result. Never mutate canonical cache rows while an overlay is pending.

Add tests for multi-cell and metadata batches, partial out-of-order acknowledgements, duplicate result IDs, abort, destroy, two concurrent edits to one cell, disabled metadata, and a committed response whose server row changes a second derived field. The latest overlay wins visually while each acknowledgement reconciles exactly once; acknowledged metadata is reread from `source.readCell`.

The source contract requires idempotency by `clientMutationId`. If transport fails after a request may have reached the server, keep the overlay as `uncertain`, do not roll it back or create new IDs, and offer retry with the exact same operation/mutation IDs or authoritative reload. Distinguish deliberate AbortError from network-unknown outcomes. Default to at most 100 pending operation batches and 1,000 unique pending cells, configurable downward/upward through `mutationLimits`; reject additional unique work with `REMOTE_MUTATION_BACKPRESSURE`. Coalesce unsent same-cell edits to the newest raw/metadata state and retain at most one queued successor behind an in-flight same-cell mutation. Add a hung/offline 10,000-edit stress test proving overlays, AbortControllers, queued successors, and diagnostics remain bounded, plus a connection-loss-after-server-commit test that resolves through same-ID retry/reload.

- [ ] **Step 4: Run mutation tests**

Run: `corepack pnpm exec vitest run src/table/remote/RemoteMutationController.test.ts src/table/remote/RemoteTableSession.test.ts src/table/remote/RemoteTableSession.contract.test.ts`

Expected: PASS for all result variants and independent overlays.

- [ ] **Step 5: Commit mutation reconciliation**

```bash
git add src/table/remote/OptimisticOverlayStore.ts src/table/remote/RemoteMutationController.ts src/table/remote/RemoteMutationController.test.ts src/table/remote/RemoteTableSession.ts src/table/remote/RemoteTableSession.test.ts src/table/remote/RemoteTableSession.contract.test.ts
git commit -m "feat: reconcile optimistic remote table edits"
```

### Task 5: Reconcile subscriptions, conflicts, and source revisions

**Files:**
- Create: `src/table/remote/RemoteSubscriptionController.ts`
- Test: `src/table/remote/RemoteSubscriptionController.test.ts`
- Modify: `src/table/remote/RemoteTableSession.ts`
- Modify: `src/table/remote/RemoteTableSession.test.ts`
- Modify: `src/table/remote/RemoteTableSession.contract.test.ts`

**Interfaces:**
- Consumes: revision-ordered `RemoteSourceEvent<TRow>`.
- Produces: canonical cache updates, invalidation, rebase, explicit conflict, and typed authoritative reload/version-checked retry behavior.

- [ ] **Step 1: Write race tests**

Cover stale event rejection, newer event application, a subscription update to a row with a pending overlay, and query invalidation. Add session-level tests that retain a conflict with stable `operationId`, `rowId`, and authoritative `revision`, then dispatch the exact shared intents:

```ts
await session.dispatch({
  type: "reload-authoritative",
  operationId: "operation-1",
  rowId: "1"
});
await session.dispatch({
  type: "retry-with-revision",
  operationId: "operation-1",
  rowId: "1",
  expectedRevision: "r3"
});
```

Use a fresh seeded-conflict session per action. Assert reload removes only the matching attempted overlay/conflict and starts an authoritative refresh; retry sends the retained attempted value/metadata with `baseRevision: "r3"` under a new mutation command ID. A blank/mismatched operation ID, row ID, or expected revision rejects without source calls or state loss. Extend `RemoteTableSession.contract.test.ts` with a real `conflictResolution` fixture so the shared contract now proves both intents are implemented rather than `unsupported`.

```ts
controller.accept({
  kind: "rows-upserted",
  revision: "r3",
  rows: [{ id: "1", salary: 115 }]
});
expect(session.getSnapshot().conflicts[0]).toMatchObject({
  rowId: "1",
  columnId: "salary",
  attemptedValue: 120,
  authoritativeValue: 115
});
```

- [ ] **Step 2: Run and verify failure**

Run: `corepack pnpm exec vitest run src/table/remote/RemoteSubscriptionController.test.ts src/table/remote/RemoteTableSession.test.ts src/table/remote/RemoteTableSession.contract.test.ts`

Expected: FAIL because subscription coordination and the two typed conflict-resolution paths are missing.

- [ ] **Step 3: Implement revision ordering and explicit conflict creation**

Call `source.compareRevisions(event.revision, currentRevision)`; never compare revision strings directly. Ignore `older`, make `equal` idempotent, and turn `unknown` into cache invalidation plus an abortable refresh. For a `newer` event, detect conflicts against optimistic cells first and call `queryController.noteCurrentRevision`. If the active query has sorting, a filter, grouping, aggregates, or any pagination, invalidate/refetch instead of patching rows because membership, order, group totals, and page boundaries may change. Under an unprojected non-paginated query, an update to an already-cached row or deletion of an existing row may patch in place; a newly inserted row still invalidates because natural server order is unknown. When an upsert changes an optimistic cell, preserve the attempted overlay and attach the authoritative value as a conflict rather than applying last-arrival-wins. `invalidate` events abort obsolete loads and refresh the last query. Unsubscribe and abort refresh work during `destroy()`.

Handle `reload-authoritative` and `retry-with-revision` in `RemoteTableSession.dispatch` by resolving the retained conflict through the tuple `(operationId, rowId)` before changing anything. Reload abandons only that conflict's attempted overlay, records a reconciliation tombstone so a late acknowledgement cannot recreate it, and performs an authoritative query refresh. Retry additionally requires `intent.expectedRevision === conflict.revision`; it creates a new command/client-mutation ID, reuses the retained attempted value or complete metadata object, and sends the expected revision as the new mutation's `baseRevision`. Never reuse the terminal conflicted mutation ID for this user-authorized new attempt. Reject stale/mismatched intents with a sanitized `REMOTE_CONFLICT_NOT_CURRENT` issue and leave the conflict intact. Both paths remain version checked and never implement last-write-wins.

Add cases where a subscription advances the current revision while a query is in flight; the later generation response is still rejected if its source revision is older. Also test an upsert changing a sort key, filter membership, group key, aggregate input, and page boundary; a paginated delete; safe unprojected existing-row replacement/deletion; new-row invalidation; delete-versus-pending-edit conflicts; duplicate events; unknown revision order; invalidation while offline; reconnect refresh; and no publication after destroy.

- [ ] **Step 4: Run subscription and mutation tests**

Run: `corepack pnpm exec vitest run src/table/remote/RemoteSubscriptionController.test.ts src/table/remote/RemoteMutationController.test.ts src/table/remote/RemoteTableSession.test.ts src/table/remote/RemoteTableSession.contract.test.ts`

Expected: PASS.

- [ ] **Step 5: Commit subscription handling**

```bash
git add src/table/remote/RemoteSubscriptionController.ts src/table/remote/RemoteSubscriptionController.test.ts src/table/remote/RemoteTableSession.ts src/table/remote/RemoteTableSession.test.ts src/table/remote/RemoteTableSession.contract.test.ts
git commit -m "feat: reconcile remote table subscriptions"
```

### Task 6: Implement version-checked remote undo

**Files:**
- Create: `src/table/remote/RemoteOperationJournal.ts`
- Test: `src/table/remote/RemoteOperationJournal.test.ts`
- Modify: `src/table/remote/RemoteTableSession.ts`
- Modify: `src/table/remote/RemoteTableSession.test.ts`
- Modify: `src/table/remote/RemoteTableSession.contract.test.ts`

**Interfaces:**
- Consumes: committed value/metadata mutations with original state and acknowledged revision.
- Produces: pending-cancel behavior and compensating mutations.

- [ ] **Step 1: Write pending and committed undo tests**

Assert pending undo aborts the whole pending command batch, removes its optimistic overlays, retains reconciliation tombstones, and invalidates/refetches because transport abort does not prove server cancellation. Simulate a server that commits and acknowledges after ignoring abort; the late acknowledgement is reconciled and the authoritative refresh wins. Assert committed value and metadata undo send one compensating batch containing every original value/complete metadata object and use the latest acknowledgement revision. Assert `undo` and `redo` return `unsupported` when compensation is absent; first-release remote redo remains disabled because the source contract advertises undo only. Finally enable undo in the remote shared-contract fixture and add its feature operation here; this is the first task whose contract may require advertised undo not to return `unsupported`.

- [ ] **Step 2: Run and verify failure**

Run: `corepack pnpm exec vitest run src/table/remote/RemoteOperationJournal.test.ts src/table/remote/RemoteTableSession.test.ts src/table/remote/RemoteTableSession.contract.test.ts`

Expected: FAIL because the journal and newly advertised undo contract behavior are missing.

- [ ] **Step 3: Implement the bounded operation journal**

Store at most 100 acknowledged command batches. Each entry records its operation ID, acknowledged revision, row versions, and a union of `{ kind: "value", originalValue, committedValue }` or `{ kind: "metadata", originalMetadata, committedMetadata }` per row/column. Pending entries retain their mutation `AbortController`. Pending undo marks a tombstone, aborts, removes visual overlays, and starts an authoritative refresh; a late acknowledgement matching the tombstone may update canonical data but never recreate history/overlays, and the refresh determines final state. A committed undo sends a new version-checked mutation batch of the matching kind, then removes the journal entry only after acknowledgement. Rejection, uncertain network outcome, or conflict remains visible and keeps the entry retryable. Do not advertise `canUndo` unless `capabilities.undo` is enabled and `source.undoMode === "compensating"`; always advertise `canRedo: false` for remote sessions in this release.

- [ ] **Step 4: Run journal and session tests**

Run: `corepack pnpm exec vitest run src/table/remote/RemoteOperationJournal.test.ts src/table/remote/RemoteTableSession.test.ts src/table/remote/RemoteTableSession.contract.test.ts`

Expected: PASS.

- [ ] **Step 5: Commit remote undo**

```bash
git add src/table/remote/RemoteOperationJournal.ts src/table/remote/RemoteOperationJournal.test.ts src/table/remote/RemoteTableSession.ts src/table/remote/RemoteTableSession.test.ts src/table/remote/RemoteTableSession.contract.test.ts
git commit -m "feat: add versioned remote table undo"
```

### Task 7: Surface loading, retry, pending, rejected, and conflict states in React

**Files:**
- Create: `src/react/RemoteTableStatus.tsx`
- Create: `src/react/ConflictPanel.tsx`
- Test: `src/react/DataTable.remote.test.tsx`
- Modify: `src/react/DataTable.tsx`
- Modify: `src/styles/data-table.css`

**Interfaces:**
- Consumes: remote fields on `TableViewSnapshot`.
- Produces: accessible status, retry, and conflict resolution UI.

- [ ] **Step 1: Write accessible remote-state tests**

```tsx
expect(screen.getByRole("status")).toHaveTextContent("Loading rows");
expect(screen.getByRole("button", { name: "Retry loading rows" })).toBeEnabled();
expect(screen.getByRole("alert")).toHaveTextContent("Salary changed on the server");
```

Assert the last good rows remain rendered during retry and conflicts expose reload and version-checked retry actions. Seed a conflict `{ operationId: "operation-1", rowId: "row-1", revision: "r3" }`, click each action in a fresh render, and assert the UI dispatches exactly:

```ts
expect(session.dispatch).toHaveBeenCalledWith({
  type: "reload-authoritative",
  operationId: "operation-1",
  rowId: "row-1"
});
expect(session.dispatch).toHaveBeenCalledWith({
  type: "retry-with-revision",
  operationId: "operation-1",
  rowId: "row-1",
  expectedRevision: "r3"
});
```

Also assert both controls disable while their returned pending command is active and announce rejection if the conflict became stale. Verify offset, cursor, infinite, known-total, and unknown-total controls; server group/aggregate rows; loaded-row scope labels; formula support labels; an offline state; and an unsupported-feature reason. A custom cell renderer exception must stay inside the cell error boundary.

- [ ] **Step 2: Run and verify failure**

Run: `corepack pnpm exec vitest run src/react/DataTable.remote.test.tsx`

Expected: FAIL because remote status components are missing.

- [ ] **Step 3: Implement the state surfaces**

Use `role="status"` with polite live announcements for loading and mutation progress, and `role="alert"` for rejected mutations and conflicts. Query retry dispatches `refresh` and is disabled while the new request is pending. `ConflictPanel` receives the matching `TableConflict` and dispatches the exact typed objects above through `TableSession.dispatch`; it never calls a remote-only method or substitutes the current visible row index. The retry intent always copies `conflict.revision` into `expectedRevision`. Conflict actions are `reload-authoritative` and `retry-with-revision`; there is no last-write-wins button. Retain the approved neutral tokens; use accent color only for active/pending emphasis and never color alone to identify conflict.

- [ ] **Step 4: Run React and axe tests**

Run: `corepack pnpm exec vitest run src/react/DataTable.remote.test.tsx src/react/DataTable.a11y.test.tsx`

Expected: PASS with no axe violations.

- [ ] **Step 5: Commit remote UI states**

```bash
git add src/react/RemoteTableStatus.tsx src/react/ConflictPanel.tsx src/react/DataTable.remote.test.tsx src/react/DataTable.tsx src/styles/data-table.css
git commit -m "feat: surface remote table lifecycle states"
```

### Task 8: Add remote race, property, and browser coverage

**Files:**
- Create: `src/table/remote/RemoteTableSession.property.test.ts`
- Create: `tests/data-table-remote.spec.ts`
- Create: `tests/fixtures/remoteTableApi.ts`
- Modify: `docs/embedding.md`

**Interfaces:**
- Consumes: the public remote source API.
- Produces: adversarial verification and a complete host integration example.

- [ ] **Step 1: Add randomized event-order tests**

Use fast-check to generate offset/cursor/infinite queries, grouped and aggregate responses, query responses, value/metadata mutations and acknowledgements, uncertain retries, subscription events, export, undo, offline/reconnect, backpressure, eviction, and destroy events. Assert stale or unknown-order revisions never replace known-newer snapshots, optimistic values are pending/uncertain/conflicted/reconciled exactly once, cache pages never exceed five, pending cells/operations never exceed configured bounds, page-local data is never labeled complete-dataset, and no listener fires after destroy.

Wrap the generated state machine with:

```ts
fc.assert(remoteEventProperty, {
  seed: 20260709,
  numRuns: 1_000,
  endOnFailure: true
});
```

Print fast-check's replay path on failure.

- [ ] **Step 2: Run the property test with a fixed replay seed**

Run: `corepack pnpm exec vitest run src/table/remote/RemoteTableSession.property.test.ts`

Expected: PASS for at least 1,000 generated event sequences.

- [ ] **Step 3: Add the Playwright remote workflow**

Exercise dataset-wide sorting, filter query serialization, grouping (with pagination cleared), aggregation, server formulas, offset/cursor/infinite pagination, known and unknown totals, edit/metadata correction, validation rejection, metadata/value conflict resolution, uncertain same-ID offline retry, backpressure, subscription refresh, pending undo with an abort-ignoring server, committed value/metadata compensating undo, loaded-row export rejection, complete-dataset `ExportArtifact`, malformed export artifact, React-handle Blob conversion, export abort/destroy, revision mismatch, and cache eviction using deterministic route fixtures. Fail the test on console errors, unhandled rejections, post-unmount publication, or request activity after destroy.

- [ ] **Step 4: Run the remote browser test**

Run: `corepack pnpm exec playwright test tests/data-table-remote.spec.ts --project=chromium`

Expected: PASS.

- [ ] **Step 5: Document the complete remote adapter**

Add a runnable TypeScript example to `docs/embedding.md` showing `capabilities`, abort-aware `query`, versioned `mutate`, and `subscribe`, including the statement that missing complete-dataset capability disables the corresponding global command.

- [ ] **Step 6: Run the remote subsystem and full unit suite**

Run:

```bash
corepack pnpm exec vitest run src/table/remote src/react/DataTable.remote.test.tsx
corepack pnpm test
```

Expected: all tests pass.

- [ ] **Step 7: Commit remote verification and docs**

```bash
git add src/table/remote/RemoteTableSession.property.test.ts tests/data-table-remote.spec.ts tests/fixtures/remoteTableApi.ts docs/embedding.md
git commit -m "test: pressure test remote data tables"
```
