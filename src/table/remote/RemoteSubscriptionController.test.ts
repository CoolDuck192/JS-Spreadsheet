import { describe, expect, it, vi } from "vitest";
import type { TableCapabilities } from "../core/capabilities";
import type { QueryRequest, QueryResult } from "../core/query";
import { createRemoteTableSource } from "./createRemoteTableSource";
import { OptimisticOverlayStore } from "./OptimisticOverlayStore";
import { RemoteMutationController } from "./RemoteMutationController";
import { RemoteQueryController } from "./RemoteQueryController";
import { RemoteSubscriptionController } from "./RemoteSubscriptionController";
import { createDeferred } from "./testUtils";
import type { RemoteMutationResult, RemoteSourceEvent } from "./types";

type Row = { id: string; salary: number; group: string };

describe("RemoteSubscriptionController", () => {
  it("ignores older/equal events and applies a newer existing-row upsert in place", async () => {
    const harness = await createHarness();

    harness.subscription.accept({
      kind: "rows-upserted",
      revision: "0",
      rows: [{ id: "1", salary: 90, group: "old" }]
    });
    harness.subscription.accept({
      kind: "rows-upserted",
      revision: "1",
      rows: [{ id: "1", salary: 100, group: "same" }]
    });
    expect(harness.queryController.getCanonicalRow("1")?.salary).toBe(100);

    harness.subscription.accept({
      kind: "rows-upserted",
      revision: "2",
      rows: [{ id: "1", salary: 110, group: "new" }]
    });
    expect(harness.queryController.getCanonicalRow("1")).toEqual({ id: "1", salary: 110, group: "new" });
    expect(harness.queryController.getCurrentRevision()).toBe("2");
  });

  it("invalidates for a newly inserted row because natural server order is unknown", async () => {
    const harness = await createHarness();
    harness.subscription.accept({
      kind: "rows-upserted",
      revision: "2",
      rows: [{ id: "2", salary: 80, group: "new" }]
    });

    expect(harness.queryController.getSnapshot().status).toBe("loading");
    await vi.waitFor(() => expect(harness.query).toHaveBeenCalledTimes(2));
  });

  it("invalidates projected queries for sort/filter/group/aggregate/page-sensitive changes", async () => {
    const projected: QueryRequest = {
      sorting: [{ columnId: "salary", direction: "asc" }],
      filter: {
        kind: "comparison",
        columnId: "group",
        operator: "eq",
        value: { type: "string", value: "A" }
      },
      grouping: [],
      aggregates: [],
      pagination: { kind: "none" }
    };
    const harness = await createHarness(projected);
    harness.subscription.accept({
      kind: "rows-upserted",
      revision: "2",
      rows: [{ id: "1", salary: 120, group: "B" }]
    });

    await vi.waitFor(() => expect(harness.query).toHaveBeenCalledTimes(2));
  });

  it("preserves a pending attempt and creates an explicit subscription conflict", async () => {
    const acknowledgement = createDeferred<readonly RemoteMutationResult<Row>[]>();
    const harness = await createHarness(undefined, async () => acknowledgement.promise);
    void harness.mutations.execute("operation-1", [{
      kind: "cell-value",
      rowId: "1",
      columnId: "salary",
      rawText: "120",
      parsedValue: 120,
      optimisticCell: {
        storedValue: 120,
        evaluatedValue: 120,
        displayValue: "120",
        metadata: {}
      }
    }]);
    expect(harness.overlays.getLatest("1", "salary")?.cell.storedValue).toBe(120);

    harness.subscription.accept({
      kind: "rows-upserted",
      revision: "2",
      rows: [{ id: "1", salary: 115, group: "A" }]
    });

    expect(harness.overlays.getLatest("1", "salary")).toMatchObject({ status: "conflict" });
    expect(harness.mutations.getConflicts()).toEqual([expect.objectContaining({
      operationId: "operation-1",
      rowId: "1",
      columnId: "salary",
      attemptedValue: 120,
      authoritativeValue: 115,
      revision: "2"
    })]);
  });

  it("settles a matching uncertain batch from subscription authority", async () => {
    const harness = await createHarness(undefined, async () => {
      throw new Error("connection lost after commit");
    });
    await harness.mutations.execute("operation-uncertain-match", [{
      kind: "cell-value",
      rowId: "1",
      columnId: "salary",
      rawText: "120",
      parsedValue: 120,
      optimisticCell: {
        storedValue: 120,
        evaluatedValue: 120,
        displayValue: "120",
        metadata: {}
      }
    }]);

    harness.subscription.accept({
      kind: "rows-upserted",
      revision: "2",
      rows: [{ id: "1", salary: 120, group: "A" }]
    });

    expect(harness.overlays.getLatest("1", "salary")).toBeUndefined();
    expect(harness.mutations.getPendingOperations()).toEqual([]);
    expect(harness.mutations.getConflicts()).toEqual([]);
  });

  it("turns an uncertain subscription mismatch into a terminal conflict", async () => {
    const harness = await createHarness(undefined, async () => {
      throw new Error("connection lost");
    });
    await harness.mutations.execute("operation-uncertain-conflict", [{
      kind: "cell-value",
      rowId: "1",
      columnId: "salary",
      rawText: "120",
      parsedValue: 120,
      optimisticCell: {
        storedValue: 120,
        evaluatedValue: 120,
        displayValue: "120",
        metadata: {}
      }
    }]);

    harness.subscription.accept({
      kind: "rows-upserted",
      revision: "2",
      rows: [{ id: "1", salary: 115, group: "A" }]
    });

    expect(harness.mutations.getPendingOperations()).toEqual([]);
    expect(harness.mutations.getConflicts()).toEqual([expect.objectContaining({
      operationId: "operation-uncertain-conflict",
      attemptedValue: 120,
      authoritativeValue: 115
    })]);
  });

  it("turns delete-versus-pending-edit into a conflict instead of dropping the attempt", async () => {
    const acknowledgement = createDeferred<readonly RemoteMutationResult<Row>[]>();
    const harness = await createHarness(undefined, async () => acknowledgement.promise);
    void harness.mutations.execute("operation-delete", [{
      kind: "cell-value",
      rowId: "1",
      columnId: "salary",
      rawText: "120",
      parsedValue: 120,
      optimisticCell: {
        storedValue: 120,
        evaluatedValue: 120,
        displayValue: "120",
        metadata: {}
      }
    }]);

    harness.subscription.accept({ kind: "rows-deleted", revision: "2", rowIds: ["1"] });

    expect(harness.mutations.getConflicts()).toEqual([expect.objectContaining({
      operationId: "operation-delete",
      rowId: "1",
      attemptedValue: 120,
      authoritativeValue: undefined,
      revision: "2"
    })]);
  });

  it("invalidates unknown revision order and rejects an older in-flight query after a newer event", async () => {
    const harness = await createHarness();
    harness.subscription.accept({ kind: "invalidate", revision: "opaque" });
    expect(harness.queryController.getSnapshot().status).toBe("loading");

    harness.queryController.noteCurrentRevision("3", false);
    harness.subscription.accept({
      kind: "rows-upserted",
      revision: "2",
      rows: [{ id: "1", salary: 105, group: "stale" }]
    });
    expect(harness.queryController.getCurrentRevision()).toBe("3");
  });

  it("subscribes once, forwards events, and stops all publication after destroy", async () => {
    let listener: ((event: RemoteSourceEvent<Row>) => void) | undefined;
    const unsubscribe = vi.fn();
    const subscribe = vi.fn((next: (event: RemoteSourceEvent<Row>) => void) => {
      listener = next;
      return unsubscribe;
    });
    const harness = await createHarness(undefined, undefined, subscribe);
    let publications = 0;
    harness.queryController.subscribe(() => { publications += 1; });
    harness.subscription.start();
    harness.subscription.start();
    expect(subscribe).toHaveBeenCalledTimes(1);

    harness.subscription.destroy();
    expect(unsubscribe).toHaveBeenCalledTimes(1);
    listener?.({
      kind: "rows-upserted",
      revision: "2",
      rows: [{ id: "1", salary: 120, group: "late" }]
    });
    expect(publications).toBe(0);
  });
});

async function createHarness(
  activeQuery: QueryRequest = unprojectedQuery(),
  mutateImplementation: ((...args: Parameters<NonNullable<ReturnType<typeof createSource>["mutate"]>>) => ReturnType<NonNullable<ReturnType<typeof createSource>["mutate"]>>) | undefined = undefined,
  subscribeImplementation?: (listener: (event: RemoteSourceEvent<Row>) => void) => () => void
) {
  const query = vi.fn(async () => result("1", [{ id: "1", salary: 100, group: "A" }]));
  const source = createSource({
    query,
    ...(mutateImplementation === undefined ? {} : { mutate: mutateImplementation }),
    ...(subscribeImplementation === undefined ? {} : { subscribe: subscribeImplementation })
  });
  const queryController = new RemoteQueryController(source);
  await queryController.load(activeQuery, "initial");
  const overlays = new OptimisticOverlayStore();
  const mutations = new RemoteMutationController({
    source,
    queryController,
    overlays,
    getActiveQuery: () => activeQuery,
    onChange: vi.fn()
  });
  let operation = 0;
  const subscription = new RemoteSubscriptionController({
    source,
    queryController,
    mutationController: mutations,
    getActiveQuery: () => activeQuery,
    createOperationId: () => `subscription-${++operation}`
  });
  return { source, queryController, overlays, mutations, subscription, query };
}

function createSource(overrides: Partial<Parameters<typeof createRemoteTableSource<Row>>[0]> = {}) {
  return createRemoteTableSource<Row>({
    getRowId: (row) => row.id,
    capabilities: capabilities(),
    paginationMode: "none",
    mutationMode: "versioned",
    undoMode: "none",
    compareRevisions: comparator,
    readCell: (row, columnId) => {
      const value = row[columnId as keyof Row];
      return { storedValue: value, evaluatedValue: value };
    },
    query: async () => result("1", []),
    mutate: async () => [],
    subscribe: () => () => {},
    ...overrides
  });
}

function capabilities(): TableCapabilities {
  const complete = { executor: "server" as const, scope: "completeDataset" as const };
  return {
    sort: { ...complete }, filter: { ...complete }, group: false, aggregate: false,
    pagination: false, edit: { ...complete }, bulkEdit: { ...complete },
    metadata: { ...complete }, validation: { ...complete }, formula: "server",
    subscription: true, undo: false, export: false
  };
}

function unprojectedQuery(): QueryRequest {
  return { sorting: [], filter: null, grouping: [], aggregates: [], pagination: { kind: "none" } };
}

function result(revision: string, rows: readonly Row[]): QueryResult<Row> {
  return {
    items: rows.map((row) => ({ kind: "data" as const, id: row.id, original: row, depth: 0 })),
    revision,
    completeness: "completeDataset",
    pageInfo: { kind: "none", total: { kind: "known", value: rows.length } }
  };
}

function comparator(candidate: string, current: string) {
  if (candidate === current) return "equal" as const;
  const left = Number(candidate);
  const right = Number(current);
  if (!Number.isFinite(left) || !Number.isFinite(right)) return "unknown" as const;
  return left > right ? "newer" as const : "older" as const;
}
