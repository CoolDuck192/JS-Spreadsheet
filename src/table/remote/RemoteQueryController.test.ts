import { describe, expect, it, vi } from "vitest";
import type { QueryRequest, QueryResult } from "../core/query";
import type { TableCapabilities } from "../core/capabilities";
import {
  RemoteQueryController,
  RemoteTableError
} from "./RemoteQueryController";
import { createDeferred, createTestRemoteSource } from "./testUtils";

type Row = { id: string; name: string };

describe("RemoteQueryController", () => {
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

  it("aborts the previous request when query state changes", () => {
    const signals: Array<{ aborted: boolean }> = [];
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

  it("publishes loading and ready snapshots once each", async () => {
    const source = createTestRemoteSource<Row>({ query: async () => result("r1", []) });
    const controller = new RemoteQueryController(source);
    const phases: string[] = [];
    controller.subscribe(() => phases.push(controller.getSnapshot().status));

    await controller.load(queryWithFilter("Ada"), "query-1");

    expect(phases).toEqual(["loading", "ready"]);
    expect(controller.getSnapshot()).toMatchObject({ status: "ready", revision: "r1" });
  });

  it("accepts equal/newer revisions, ignores older revisions, and invalidates unknown order", async () => {
    const responses = [
      result("1", [{ id: "1", name: "Ada" }]),
      result("2", [{ id: "2", name: "Grace" }]),
      result("1", [{ id: "stale", name: "Stale" }]),
      result("opaque", [{ id: "unknown", name: "Unknown" }])
    ];
    const source = createTestRemoteSource<Row>({
      compareRevisions(candidate, current) {
        if (candidate === current) return "equal";
        const left = Number(candidate);
        const right = Number(current);
        if (!Number.isFinite(left) || !Number.isFinite(right)) return "unknown";
        return left > right ? "newer" : "older";
      },
      query: async () => responses.shift()!
    });
    const controller = new RemoteQueryController(source);

    await controller.load(queryWithFilter("Ada"), "query-1");
    await controller.load(queryWithFilter("Grace"), "query-2");
    expect(controller.getSnapshot().revision).toBe("2");
    await controller.load(queryWithFilter("Stale"), "query-3");
    expect(controller.getSnapshot()).toMatchObject({ status: "ready", revision: "2" });
    expect(controller.getSnapshot().items[0].id).toBe("2");
    await controller.load(queryWithFilter("Unknown"), "query-4");
    expect(controller.getSnapshot()).toMatchObject({
      status: "error",
      revision: "2",
      items: [],
      error: { code: "REVISION_ORDER_UNKNOWN", retryable: true }
    });
  });

  it("notifies reconciliation listeners only for revision-accepted query results", async () => {
    const responses = [
      result("1", [{ id: "1", name: "Ada" }]),
      result("0", [{ id: "stale", name: "Stale" }]),
      result("2", [{ id: "2", name: "Grace" }])
    ];
    const source = createTestRemoteSource<Row>({
      compareRevisions: numericComparator,
      query: async () => responses.shift()!
    });
    const controller = new RemoteQueryController(source);
    const accepted = vi.fn();
    controller.subscribeAccepted(accepted);

    await controller.load(queryWithFilter("Ada"), "query-1");
    await controller.load(queryWithFilter("Stale"), "query-2");
    await controller.load(queryWithFilter("Grace"), "query-3");

    expect(accepted).toHaveBeenCalledTimes(2);
    expect(accepted.mock.calls.map(([acceptance]) => ({
      generation: acceptance.generation,
      revision: acceptance.revision,
      filter: acceptance.query.filter,
      rowIds: acceptance.items.map((item: { id: string }) => item.id)
    }))).toEqual([
      { generation: 1, revision: "1", filter: queryWithFilter("Ada").filter, rowIds: ["1"] },
      { generation: 3, revision: "2", filter: queryWithFilter("Grace").filter, rowIds: ["2"] }
    ]);
  });

  it("keeps canonical rows published while an older refresh is in flight", async () => {
    const refresh = createDeferred<QueryResult<Row>>();
    const source = createTestRemoteSource<Row>({
      compareRevisions: numericComparator,
      query: vi.fn()
        .mockResolvedValueOnce(result("1", [{ id: "1", name: "Ada" }]))
        .mockReturnValueOnce(refresh.promise)
    });
    const controller = new RemoteQueryController(source);

    await controller.load(queryWithFilter("Ada"), "initial");
    const pending = controller.refresh("refresh");
    controller.applyCanonicalRows([{ id: "1", name: "Ada Lovelace" }], "2");
    refresh.resolve(result("1", [{ id: "1", name: "Ada" }]));
    await pending;

    expect(controller.getSnapshot()).toMatchObject({
      status: "ready",
      revision: "2",
      items: [{ kind: "data", id: "1", original: { id: "1", name: "Ada Lovelace" } }]
    });
  });

  it("keeps canonical rows published when an in-flight refresh fails", async () => {
    const refresh = createDeferred<QueryResult<Row>>();
    const source = createTestRemoteSource<Row>({
      compareRevisions: numericComparator,
      query: vi.fn()
        .mockResolvedValueOnce(result("1", [{ id: "1", name: "Ada" }]))
        .mockReturnValueOnce(refresh.promise)
    });
    const controller = new RemoteQueryController(source);

    await controller.load(queryWithFilter("Ada"), "initial");
    const pending = controller.refresh("refresh");
    controller.applyCanonicalRows([{ id: "1", name: "Ada Lovelace" }], "2");
    refresh.reject(new Error("refresh failed"));
    await pending;

    expect(controller.getSnapshot()).toMatchObject({
      status: "error",
      revision: "2",
      items: [{ kind: "data", id: "1", original: { id: "1", name: "Ada Lovelace" } }]
    });
  });

  it("replaces offset pages but appends infinite pages in request order", async () => {
    const offsetResponses = [
      result("1", [{ id: "1", name: "Ada" }], 0),
      result("2", [{ id: "2", name: "Grace" }], 50)
    ];
    const offset = new RemoteQueryController(createTestRemoteSource<Row>({
      compareRevisions: numericComparator,
      query: async () => offsetResponses.shift()!
    }));
    await offset.load(offsetRequest(0), "offset-1");
    await offset.load(offsetRequest(50), "offset-2");
    expect(offset.getSnapshot().items.map((item) => item.id)).toEqual(["2"]);

    const infiniteResponses: QueryResult<Row>[] = [
      infiniteResult("1", [{ id: "1", name: "Ada" }], "cursor-1"),
      infiniteResult("2", [{ id: "2", name: "Grace" }])
    ];
    const infinite = new RemoteQueryController(createTestRemoteSource<Row>({
      capabilities: infiniteCapabilities(),
      paginationMode: "infinite",
      compareRevisions: numericComparator,
      query: async () => infiniteResponses.shift()!
    }));
    await infinite.load(infiniteRequest(), "infinite-1");
    await infinite.load(infiniteRequest("cursor-1"), "infinite-2");
    expect(infinite.getSnapshot().items.map((item) => item.id)).toEqual(["1", "2"]);
  });

  it.each(["cursor", "infinite"] as const)(
    "replays the loaded %s window from the head after invalidation",
    async (kind) => {
      const pageRows = [rows(1, 50), rows(51, 50), rows(101, 50)];
      const responses = [
        accumulatedResult(kind, "1", pageRows[0], "page-2"),
        accumulatedResult(kind, "2", pageRows[1], "page-3"),
        accumulatedResult(kind, "3", pageRows[2]),
        accumulatedResult(kind, "4", pageRows[0], "refreshed-page-2"),
        accumulatedResult(kind, "5", pageRows[1], "refreshed-page-3"),
        accumulatedResult(kind, "6", pageRows[2])
      ];
      const query = vi.fn(async (_request: QueryRequest) => responses.shift()!);
      const controller = new RemoteQueryController(createTestRemoteSource<Row>({
        capabilities: accumulatedCapabilities(kind),
        paginationMode: kind,
        compareRevisions: numericComparator,
        query
      }));

      await controller.load(accumulatedRequest(kind), "page-1");
      await controller.load(accumulatedRequest(kind, "page-2"), "page-2");
      await controller.load(accumulatedRequest(kind, "page-3"), "page-3");
      expect(controller.getSnapshot().items).toHaveLength(150);

      controller.invalidate();
      await controller.refresh("refresh");

      expect(controller.getSnapshot()).toMatchObject({ status: "ready", revision: "6" });
      expect(controller.getSnapshot().items.map((item) => item.id)).toEqual(
        Array.from({ length: 150 }, (_, index) => String(index + 1))
      );
      expect(query.mock.calls.slice(3).map(([request]) => request.pagination)).toEqual([
        accumulatedRequest(kind).pagination,
        accumulatedRequest(kind, "refreshed-page-2").pagination,
        accumulatedRequest(kind, "refreshed-page-3").pagination
      ]);
    }
  );

  it("stops accumulated replay when a newer load supersedes it", async () => {
    const replayHead = createDeferred<QueryResult<Row>>();
    const query = vi.fn()
      .mockResolvedValueOnce(accumulatedResult("infinite", "1", rows(1, 1), "page-2"))
      .mockResolvedValueOnce(accumulatedResult("infinite", "2", rows(2, 1)))
      .mockReturnValueOnce(replayHead.promise)
      .mockResolvedValueOnce(accumulatedResult("infinite", "10", rows(999, 1), "user-page-2"))
      .mockResolvedValueOnce(accumulatedResult("infinite", "11", rows(2, 1)));
    const controller = new RemoteQueryController(createTestRemoteSource<Row>({
      capabilities: accumulatedCapabilities("infinite"),
      paginationMode: "infinite",
      compareRevisions: numericComparator,
      query
    }));
    await controller.load(accumulatedRequest("infinite"), "page-1");
    await controller.load(accumulatedRequest("infinite", "page-2"), "page-2");

    controller.invalidate();
    const replay = controller.refresh("refresh");
    await controller.load(accumulatedRequest("infinite"), "user-load");
    replayHead.resolve(accumulatedResult("infinite", "3", rows(1, 1), "stale-page-2"));
    await replay;

    expect(query).toHaveBeenCalledTimes(4);
    expect(controller.getSnapshot().items.map((item) => item.id)).toEqual(["999"]);
    expect(controller.getSnapshot().revision).toBe("10");
  });

  it("rejects refresh before load and stops all work after destroy", async () => {
    const query = vi.fn(async () => result("r1", []));
    const controller = new RemoteQueryController(createTestRemoteSource<Row>({ query }));
    let publications = 0;
    controller.subscribe(() => { publications += 1; });

    await expect(controller.refresh("refresh-before-load")).rejects.toMatchObject({ code: "NO_QUERY" });
    controller.destroy();
    await expect(controller.load(offsetRequest(0), "after-destroy")).rejects.toMatchObject({
      code: "REMOTE_SESSION_DESTROYED"
    });
    await expect(controller.refresh("refresh-after-destroy")).rejects.toBeInstanceOf(RemoteTableError);
    expect(query).not.toHaveBeenCalled();
    expect(publications).toBe(0);
  });

  it("rejects an unsupported pagination request before calling the host", async () => {
    const query = vi.fn(async () => result("r1", []));
    const controller = new RemoteQueryController(createTestRemoteSource<Row>({ query }));

    await expect(controller.load(infiniteRequest(), "wrong-mode")).rejects.toMatchObject({
      code: "UNSUPPORTED_PAGINATION"
    });
    expect(query).not.toHaveBeenCalled();
  });

  it("does not publish an in-flight result after destruction", async () => {
    const pending = createDeferred<QueryResult<Row>>();
    const controller = new RemoteQueryController(createTestRemoteSource<Row>({
      query: async () => pending.promise
    }));
    let publications = 0;
    controller.subscribe(() => { publications += 1; });
    const load = controller.load(offsetRequest(0), "pending");
    expect(publications).toBe(1);

    controller.destroy();
    pending.resolve(result("r1", [{ id: "1", name: "Ada" }]));
    await load;

    expect(publications).toBe(1);
  });
});

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

function offsetRequest(offset: number): QueryRequest {
  return { sorting: [], filter: null, grouping: [], aggregates: [], pagination: { kind: "offset", offset, limit: 50 } };
}

function infiniteRequest(after?: string): QueryRequest {
  return {
    sorting: [], filter: null, grouping: [], aggregates: [],
    pagination: after === undefined
      ? { kind: "infinite", limit: 50 }
      : { kind: "infinite", after, limit: 50 }
  };
}

function result(revision: string, rows: readonly Row[], offset = 0): QueryResult<Row> {
  return {
    items: rows.map((row) => ({ kind: "data" as const, id: row.id, original: row, depth: 0 })),
    revision,
    completeness: "loadedRows",
    pageInfo: {
      kind: "offset", offset, limit: 50,
      total: { kind: "known", value: rows.length }, hasMore: false
    }
  };
}

function infiniteResult(
  revision: string,
  rows: readonly Row[],
  nextCursor?: string
): QueryResult<Row> {
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

function accumulatedRequest(kind: "cursor" | "infinite", token?: string): QueryRequest {
  return {
    sorting: [], filter: null, grouping: [], aggregates: [],
    pagination: kind === "cursor"
      ? { kind, ...(token === undefined ? {} : { cursor: token }), limit: 50 }
      : { kind, ...(token === undefined ? {} : { after: token }), limit: 50 }
  };
}

function accumulatedResult(
  kind: "cursor" | "infinite",
  revision: string,
  resultRows: readonly Row[],
  nextCursor?: string
): QueryResult<Row> {
  return {
    items: resultRows.map((row) => ({ kind: "data" as const, id: row.id, original: row, depth: 0 })),
    revision,
    completeness: "loadedRows",
    pageInfo: kind === "cursor"
      ? { kind, ...(nextCursor === undefined ? {} : { nextCursor }), total: { kind: "unknown" } }
      : {
          kind,
          ...(nextCursor === undefined ? {} : { nextCursor }),
          loadedCount: resultRows.length,
          total: { kind: "unknown" }
        }
  };
}

function accumulatedCapabilities(kind: "cursor" | "infinite"): TableCapabilities {
  const source = createTestRemoteSource<Row>();
  return {
    ...source.capabilities,
    pagination: { executor: "server", scope: "completeDataset", modes: [kind] }
  };
}

function rows(start: number, length: number): Row[] {
  return Array.from({ length }, (_, index) => ({
    id: String(start + index),
    name: `Row ${start + index}`
  }));
}

function infiniteCapabilities(): TableCapabilities {
  const source = createTestRemoteSource<Row>();
  return {
    ...source.capabilities,
    pagination: { executor: "server", scope: "completeDataset", modes: ["infinite"] }
  };
}

function numericComparator(candidate: string, current: string) {
  if (candidate === current) return "equal" as const;
  return Number(candidate) > Number(current) ? "newer" as const : "older" as const;
}
