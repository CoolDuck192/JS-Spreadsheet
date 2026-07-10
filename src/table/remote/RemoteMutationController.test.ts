import { describe, expect, it, vi } from "vitest";
import type { TableCapabilities } from "../core/capabilities";
import type { QueryRequest, QueryResult } from "../core/query";
import { createRemoteTableSource } from "./createRemoteTableSource";
import { OptimisticOverlayStore } from "./OptimisticOverlayStore";
import { RemoteMutationController, type PreparedRemoteMutation } from "./RemoteMutationController";
import { RemoteQueryController } from "./RemoteQueryController";
import { createDeferred } from "./testUtils";
import type { RemoteMutationResult } from "./types";

type Row = { id: string; salary: number; label: string };

describe("RemoteMutationController", () => {
  it("keeps canonical rows untouched while pending and commits the acknowledged row", async () => {
    const acknowledgement = createDeferred<readonly RemoteMutationResult<Row>[]>();
    const harness = await createHarness(async () => acknowledgement.promise);
    const execution = harness.controller.execute("operation-1", [valueMutation(120)]);

    expect(harness.overlays.getLatest("1", "salary")).toMatchObject({
      clientMutationId: "operation-1:0",
      status: "pending",
      cell: { storedValue: 120, evaluatedValue: 120 }
    });
    expect(harness.queryController.getCanonicalRow("1")?.salary).toBe(100);
    expect(harness.controller.getPendingOperations()).toHaveLength(1);
    expect(harness.mutate).toHaveBeenCalledWith(
      [expect.objectContaining({
        clientMutationId: "operation-1:0",
        rowId: "1",
        columnId: "salary",
        baseRevision: "1",
        parsedValue: 120
      })],
      expect.objectContaining({ operationId: "operation-1", signal: expect.any(Object) })
    );

    acknowledgement.resolve([{
      clientMutationId: "operation-1:0",
      status: "committed",
      revision: "2",
      row: { id: "1", salary: 120, label: "derived-120" },
      rowVersion: "row-2"
    }]);
    await expect(execution).resolves.toMatchObject({ status: "committed", changed: true });
    expect(harness.overlays.getLatest("1", "salary")).toBeUndefined();
    expect(harness.queryController.getCanonicalRow("1")).toEqual({
      id: "1", salary: 120, label: "derived-120"
    });
    expect(harness.controller.getPendingOperations()).toEqual([]);
  });

  it("accepts a corrected row as authoritative", async () => {
    const harness = await createHarness(async (batch) => [{
      clientMutationId: batch[0].clientMutationId,
      status: "corrected",
      revision: "2",
      row: { id: "1", salary: 115, label: "capped" }
    }]);

    await expect(harness.controller.execute("operation-corrected", [valueMutation(120)]))
      .resolves.toMatchObject({ status: "committed" });
    expect(harness.queryController.getCanonicalRow("1")?.salary).toBe(115);
    expect(harness.overlays.size).toBe(0);
  });

  it("removes only rejected overlays and returns authoritative issues", async () => {
    const harness = await createHarness(async (batch) => [{
      clientMutationId: batch[0].clientMutationId,
      status: "rejected",
      revision: "2",
      issues: [{ code: "SALARY_LIMIT", message: "Salary is too high", columnId: "salary" }]
    }]);

    await expect(harness.controller.execute("operation-rejected", [valueMutation(120)]))
      .resolves.toMatchObject({
        status: "rejected",
        reason: "validation",
        issues: [{ code: "SALARY_LIMIT", message: "Salary is too high" }]
      });
    expect(harness.queryController.getCanonicalRow("1")?.salary).toBe(100);
    expect(harness.overlays.size).toBe(0);
  });

  it("retains an attempted overlay and authoritative row for conflicts", async () => {
    const harness = await createHarness(async (batch) => [{
      clientMutationId: batch[0].clientMutationId,
      status: "conflict",
      revision: "2",
      current: { id: "1", salary: 110, label: "server" },
      rowVersion: "row-2"
    }]);

    await expect(harness.controller.execute("operation-conflict", [valueMutation(120)]))
      .resolves.toMatchObject({ status: "conflict", revision: "2" });
    expect(harness.overlays.getLatest("1", "salary")).toMatchObject({ status: "conflict" });
    expect(harness.controller.getConflicts()).toEqual([expect.objectContaining({
      operationId: "operation-conflict",
      rowId: "1",
      columnId: "salary",
      attemptedValue: 120,
      authoritativeValue: 110,
      revision: "2"
    })]);
  });

  it("matches out-of-order results by mutation ID for a multi-cell batch", async () => {
    const harness = await createHarness(async (batch) => [
      {
        clientMutationId: batch[1].clientMutationId,
        status: "committed",
        revision: "2",
        row: { id: "1", salary: 120, label: "updated" }
      },
      {
        clientMutationId: batch[0].clientMutationId,
        status: "committed",
        revision: "2",
        row: { id: "1", salary: 120, label: "updated" }
      }
    ]);
    const labelMutation: PreparedRemoteMutation = {
      kind: "cell-value",
      rowId: "1",
      columnId: "label",
      rawText: "updated",
      parsedValue: "updated",
      optimisticCell: {
        storedValue: "updated", evaluatedValue: "updated", displayValue: "updated", metadata: {}
      }
    };

    await expect(harness.controller.execute("operation-batch", [valueMutation(120), labelMutation]))
      .resolves.toMatchObject({ status: "committed" });
    expect(harness.overlays.size).toBe(0);
  });

  it.each(["missing", "duplicate", "unknown"] as const)(
    "invalidates on %s mutation result IDs",
    async (kind) => {
      const harness = await createHarness(async (batch) => {
        const valid = {
          clientMutationId: batch[0].clientMutationId,
          status: "committed" as const,
          revision: "2",
          row: { id: "1", salary: 120, label: "updated" }
        };
        if (kind === "missing") return [];
        if (kind === "duplicate") return [valid, valid];
        return [{ ...valid, clientMutationId: "not-requested" }];
      });

      await expect(harness.controller.execute("operation-protocol", [valueMutation(120)]))
        .resolves.toMatchObject({
          status: "rejected",
          reason: "validation",
          issues: [{ code: "REMOTE_MUTATION_PROTOCOL_ERROR" }]
        });
      expect(harness.overlays.size).toBe(0);
      expect(harness.queryController.getSnapshot().status).toBe("error");
    }
  );

  it("keeps a transport-uncertain overlay and returns a pending operation", async () => {
    const harness = await createHarness(async () => { throw new Error("connection lost"); });

    await expect(harness.controller.execute("operation-uncertain", [valueMutation(120)]))
      .resolves.toEqual({ status: "pending", operationId: "operation-uncertain" });
    expect(harness.overlays.getLatest("1", "salary")).toMatchObject({ status: "uncertain" });
    expect(harness.controller.getPendingOperations()).toHaveLength(1);
  });

  it("shows the latest concurrent same-cell overlay while reconciling each acknowledgement once", async () => {
    const first = createDeferred<readonly RemoteMutationResult<Row>[]>();
    const second = createDeferred<readonly RemoteMutationResult<Row>[]>();
    const harness = await createHarness(vi.fn()
      .mockReturnValueOnce(first.promise)
      .mockReturnValueOnce(second.promise));
    const firstExecution = harness.controller.execute("operation-1", [valueMutation(110)]);
    const secondExecution = harness.controller.execute("operation-2", [valueMutation(120)]);
    expect(harness.overlays.getLatest("1", "salary")?.cell.storedValue).toBe(120);

    first.resolve([{
      clientMutationId: "operation-1:0", status: "committed", revision: "2",
      row: { id: "1", salary: 110, label: "first" }
    }]);
    await firstExecution;
    expect(harness.overlays.getLatest("1", "salary")?.cell.storedValue).toBe(120);
    expect(harness.queryController.getCanonicalRow("1")?.salary).toBe(110);

    second.resolve([{
      clientMutationId: "operation-2:0", status: "committed", revision: "3",
      row: { id: "1", salary: 120, label: "second" }
    }]);
    await secondExecution;
    expect(harness.overlays.size).toBe(0);
    expect(harness.queryController.getCanonicalRow("1")?.salary).toBe(120);
  });

  it("enforces operation and unique-cell backpressure without a host call", async () => {
    const pending = new Promise<readonly RemoteMutationResult<Row>[]>(() => undefined);
    const harness = await createHarness(async () => pending, {
      maxPendingOperations: 1,
      maxPendingCells: 1
    });
    void harness.controller.execute("operation-1", [valueMutation(110)]);

    const secondMutation: PreparedRemoteMutation = {
      kind: "cell-value",
      rowId: "1",
      columnId: "label",
      rawText: "next",
      parsedValue: "next",
      optimisticCell: { storedValue: "next", evaluatedValue: "next", displayValue: "next", metadata: {} }
    };
    await expect(harness.controller.execute("operation-2", [secondMutation])).resolves.toMatchObject({
      status: "rejected",
      issues: [{ code: "REMOTE_MUTATION_BACKPRESSURE" }]
    });
    expect(harness.mutate).toHaveBeenCalledTimes(1);
  });

  it("aborts pending work, removes overlays, and ignores a late result after destroy", async () => {
    const acknowledgement = createDeferred<readonly RemoteMutationResult<Row>[]>();
    let signal: { aborted: boolean } | undefined;
    const harness = await createHarness((_batch, context) => {
      signal = context.signal;
      return acknowledgement.promise;
    });
    const execution = harness.controller.execute("operation-destroy", [valueMutation(120)]);

    harness.controller.destroy();
    expect(signal?.aborted).toBe(true);
    expect(harness.overlays.size).toBe(0);
    acknowledgement.resolve([{
      clientMutationId: "operation-destroy:0", status: "committed", revision: "2",
      row: { id: "1", salary: 120, label: "late" }
    }]);
    await execution;
    expect(harness.queryController.getCanonicalRow("1")?.salary).toBe(100);
  });
});

function valueMutation(value: number): PreparedRemoteMutation {
  return {
    kind: "cell-value",
    rowId: "1",
    columnId: "salary",
    rawText: String(value),
    parsedValue: value,
    optimisticCell: {
      storedValue: value,
      evaluatedValue: value,
      displayValue: String(value),
      metadata: {}
    }
  };
}

async function createHarness(
  mutateImplementation: NonNullable<Parameters<typeof createRemoteTableSource<Row>>[0]["mutate"]>,
  limits?: { maxPendingOperations?: number; maxPendingCells?: number }
) {
  const mutate = vi.fn(mutateImplementation);
  const source = createRemoteTableSource<Row>({
    getRowId: (row) => row.id,
    capabilities: capabilities(),
    paginationMode: "none",
    mutationMode: "versioned",
    undoMode: "none",
    compareRevisions: numericComparator,
    readCell: (row, columnId) => {
      const value = row[columnId as keyof Row];
      return { storedValue: value, evaluatedValue: value };
    },
    query: async () => canonicalResult(),
    mutate
  });
  const queryController = new RemoteQueryController(source);
  await queryController.load(query(), "initial-query");
  const overlays = new OptimisticOverlayStore();
  const controller = new RemoteMutationController({
    source,
    queryController,
    overlays,
    getActiveQuery: query,
    onChange: vi.fn(),
    limits
  });
  return { source, queryController, overlays, controller, mutate };
}

function capabilities(): TableCapabilities {
  const complete = { executor: "server" as const, scope: "completeDataset" as const };
  return {
    sort: { ...complete }, filter: { ...complete }, group: false, aggregate: false,
    pagination: false, edit: { ...complete }, bulkEdit: { ...complete },
    metadata: { ...complete }, validation: { ...complete }, formula: "server",
    subscription: false, undo: false, export: false
  };
}

function query(): QueryRequest {
  return { sorting: [], filter: null, grouping: [], aggregates: [], pagination: { kind: "none" } };
}

function canonicalResult(): QueryResult<Row> {
  const row = { id: "1", salary: 100, label: "original" };
  return {
    items: [{ kind: "data", id: row.id, original: row, depth: 0 }],
    revision: "1",
    completeness: "completeDataset",
    pageInfo: { kind: "none", total: { kind: "known", value: 1 } }
  };
}

function numericComparator(candidate: string, current: string) {
  if (candidate === current) return "equal" as const;
  return Number(candidate) > Number(current) ? "newer" as const : "older" as const;
}
