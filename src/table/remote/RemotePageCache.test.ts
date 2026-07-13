import { describe, expect, it, vi } from "vitest";
import type { QueryRequest, QueryResult } from "../core/query";
import {
  getRemotePageKey,
  RemotePageCache,
  RemotePageCacheError
} from "./RemotePageCache";

type Row = { id: string; name: string };

describe("RemotePageCache", () => {
  it("creates stable keys for every pagination mode", () => {
    expect(getRemotePageKey({ kind: "none" })).toBe("none");
    expect(getRemotePageKey({ kind: "offset", offset: 50, limit: 25 })).toBe("offset:50:25");
    expect(getRemotePageKey({ kind: "cursor", cursor: "next/page", limit: 25 }))
      .toBe("cursor:next%2Fpage:25");
    expect(getRemotePageKey({ kind: "infinite", after: "row:9", limit: 100 }))
      .toBe("infinite:row%3A9:100");
  });

  it("promotes reads for LRU eviction while preserving query aggregation order", () => {
    const cache = new RemotePageCache<Row>(2);
    const first = offsetRequest(0);
    const second = offsetRequest(10);
    const third = offsetRequest(20);
    cache.put(first, offsetResult("r1", 0, [{ id: "1", name: "Ada" }]));
    cache.put(second, offsetResult("r2", 10, [{ id: "2", name: "Grace" }]));

    expect(cache.get(first)?.revision).toBe("r1");
    cache.put(third, offsetResult("r3", 20, [{ id: "3", name: "Linus" }]));

    expect(cache.get(second)).toBeUndefined();
    expect(cache.get(first)?.items[0].id).toBe("1");
    expect(cache.get(third)?.items[0].id).toBe("3");
    expect(cache.combine()?.items.map((item) => item.id)).toEqual(["1", "3"]);
  });

  it("replaces an existing page without consuming capacity or changing its order", () => {
    const cache = new RemotePageCache<Row>();
    const first = offsetRequest(0);
    const second = offsetRequest(10);
    cache.put(first, offsetResult("r1", 0, [{ id: "1", name: "Ada" }]));
    cache.put(second, offsetResult("r2", 10, [{ id: "2", name: "Grace" }]));
    cache.put(first, offsetResult("r3", 0, [{ id: "1", name: "Ada Lovelace" }]));

    expect(cache.size).toBe(2);
    expect(cache.combine()?.items.map((item) => item.id)).toEqual(["1", "2"]);
    expect(cache.get(first)?.items[0]).toMatchObject({
      kind: "data",
      original: { name: "Ada Lovelace" }
    });
  });

  it("bounds the cache to five pages by default", () => {
    const cache = new RemotePageCache<Row>();
    for (let page = 0; page < 6; page += 1) {
      cache.put(
        offsetRequest(page * 10),
        offsetResult(`r${page}`, page * 10, [{ id: String(page), name: String(page) }])
      );
    }
    expect(cache.size).toBe(5);
    expect(cache.get(offsetRequest(0))).toBeUndefined();
  });

  it.each(["cursor", "infinite"] as const)(
    "bounds accumulated %s pages and records an explicit prefix gap",
    (kind) => {
      const cache = new RemotePageCache<Row>(1);
      const head = accumulatedRequest(kind);
      cache.put(
        head,
        accumulatedResult(kind, "r1", [{ id: "1", name: "Ada" }], "next")
      );
      cache.put(
        accumulatedRequest(kind, "next"),
        accumulatedResult(kind, "r2", [{ id: "2", name: "Grace" }])
      );

      expect(cache.size).toBe(1);
      expect(cache.combine()?.items.map((item) => item.id)).toEqual(["2"]);
      expect(cache.getPageGaps()).toEqual([{
        kind: "evicted-pages",
        at: 0,
        omittedPages: 1,
        omittedItems: 1
      }]);
      expect(cache.getRecoveryWindow()).toEqual({ headQuery: head, pageCount: 2 });
    }
  );

  it("validates only the candidate page IDs", () => {
    const reads = vi.fn();
    const cache = new RemotePageCache<Row>(3);
    cache.put(
      accumulatedRequest("infinite"),
      trackedAccumulatedResult("1", { id: "1", name: "Ada" }, "next", reads)
    );
    expect(reads).toHaveBeenCalledTimes(1);

    cache.put(
      accumulatedRequest("infinite", "next"),
      trackedAccumulatedResult("2", { id: "2", name: "Grace" }, "last", reads)
    );
    expect(reads).toHaveBeenCalledTimes(2);
  });

  it("releases incremental ID owners when accumulated pages are evicted", () => {
    const cache = new RemotePageCache<Row>(1);
    cache.put(
      accumulatedRequest("infinite"),
      accumulatedResult("infinite", "1", [{ id: "1", name: "Ada" }], "next")
    );
    cache.put(
      accumulatedRequest("infinite", "next"),
      accumulatedResult("infinite", "2", [{ id: "2", name: "Grace" }], "last")
    );
    expect(() => cache.put(
      accumulatedRequest("infinite", "last"),
      accumulatedResult("infinite", "3", [{ id: "1", name: "Ada again" }])
    )).not.toThrow();
  });

  it("coalesces FIFO prefix evictions even when an old page was recently read", () => {
    const cache = new RemotePageCache<Row>(2);
    const requests = [0, 1, 2, 3].map((page) =>
      accumulatedRequest("infinite", page === 0 ? undefined : `page-${page + 1}`));
    cache.put(requests[0], accumulatedResult("infinite", "1", [{ id: "1", name: "1" }], "page-2"));
    cache.put(requests[1], accumulatedResult("infinite", "2", [{ id: "2", name: "2" }], "page-3"));
    expect(cache.get(requests[0])).toBeDefined();
    cache.put(requests[2], accumulatedResult("infinite", "3", [{ id: "3", name: "3" }], "page-4"));
    cache.put(requests[3], accumulatedResult("infinite", "4", [{ id: "4", name: "4" }]));

    expect(cache.combine()?.items.map((item) => item.id)).toEqual(["3", "4"]);
    expect(cache.getPageGaps()).toEqual([{
      kind: "evicted-pages", at: 0, omittedPages: 2, omittedItems: 2
    }]);
    expect(cache.getRecoveryWindow()).toEqual({ headQuery: requests[0], pageCount: 4 });
  });

  it("clones only pages touched by canonical row events and advances revision in O(1)", () => {
    const cache = new RemotePageCache<Row>(3);
    const first = accumulatedRequest("infinite");
    const second = accumulatedRequest("infinite", "next");
    cache.put(first, accumulatedResult("infinite", "r1", [{ id: "1", name: "Ada" }], "next"));
    cache.put(second, accumulatedResult("infinite", "r2", [{ id: "2", name: "Grace" }]));
    const firstResult = cache.get(first)!;
    const secondResult = cache.get(second)!;
    const firstItems = firstResult.items;
    const secondItems = secondResult.items;

    cache.applyCanonicalRows([{ id: "1", name: "Ada Lovelace" }], "r3", (row) => row.id);

    expect(cache.get(first)).not.toBe(firstResult);
    expect(cache.get(second)).toBe(secondResult);
    expect(cache.get(first)!.items).not.toBe(firstItems);
    expect(cache.get(second)!.items).toBe(secondItems);
    expect(cache.combine()?.revision).toBe("r3");

    const changedFirstResult = cache.get(first)!;
    const changedFirstItems = changedFirstResult.items;
    cache.deleteCanonicalRows(["2"], "r4");
    expect(cache.get(first)).toBe(changedFirstResult);
    expect(cache.get(second)).not.toBe(secondResult);
    expect(cache.get(first)!.items).toBe(changedFirstItems);
    expect(cache.get(second)!.items).not.toBe(secondItems);
    expect(cache.combine()?.revision).toBe("r4");

    const changedSecondResult = cache.get(second)!;
    const changedSecondItems = changedSecondResult.items;
    cache.noteRevision("r5");
    expect(cache.get(first)).toBe(changedFirstResult);
    expect(cache.get(second)).toBe(changedSecondResult);
    expect(cache.get(first)!.items).toBe(changedFirstItems);
    expect(cache.get(second)!.items).toBe(changedSecondItems);
    expect(cache.combine()?.revision).toBe("r5");
  });

  it("invalidates the complete query cache when combined pages duplicate an ID", () => {
    const cache = new RemotePageCache<Row>();
    cache.put(offsetRequest(0), offsetResult("r1", 0, [{ id: "same", name: "Ada" }]));

    expect(() => cache.put(
      offsetRequest(10),
      offsetResult("r2", 10, [{ id: "same", name: "Grace" }])
    )).toThrow(RemotePageCacheError);
    expect(cache.size).toBe(0);
    expect(cache.combine()).toBeNull();
  });
});

function offsetRequest(offset: number): QueryRequest {
  return {
    sorting: [], filter: null, grouping: [], aggregates: [],
    pagination: { kind: "offset", offset, limit: 10 }
  };
}

function offsetResult(
  revision: string,
  offset: number,
  rows: readonly Row[]
): QueryResult<Row> {
  return {
    items: rows.map((row) => ({ kind: "data" as const, id: row.id, original: row, depth: 0 })),
    revision,
    completeness: "loadedRows",
    pageInfo: {
      kind: "offset",
      offset,
      limit: 10,
      total: { kind: "known", value: 100 },
      hasMore: offset + rows.length < 100
    }
  };
}

function accumulatedRequest(kind: "cursor" | "infinite", token?: string): QueryRequest {
  return {
    sorting: [], filter: null, grouping: [], aggregates: [],
    pagination: kind === "cursor"
      ? { kind, ...(token === undefined ? {} : { cursor: token }), limit: 10 }
      : { kind, ...(token === undefined ? {} : { after: token }), limit: 10 }
  };
}

function accumulatedResult(
  kind: "cursor" | "infinite",
  revision: string,
  rows: readonly Row[],
  nextCursor?: string
): QueryResult<Row> {
  return {
    items: rows.map((row) => ({ kind: "data" as const, id: row.id, original: row, depth: 0 })),
    revision,
    completeness: "loadedRows",
    pageInfo: kind === "cursor"
      ? { kind, ...(nextCursor === undefined ? {} : { nextCursor }), total: { kind: "unknown" } }
      : {
          kind,
          ...(nextCursor === undefined ? {} : { nextCursor }),
          loadedCount: rows.length,
          total: { kind: "unknown" }
        }
  };
}

function trackedAccumulatedResult(
  revision: string,
  row: Row,
  nextCursor: string | undefined,
  onIdRead: () => void
): QueryResult<Row> {
  return {
    items: [{
      kind: "data" as const,
      get id() {
        onIdRead();
        return row.id;
      },
      original: row,
      depth: 0
    }],
    revision,
    completeness: "loadedRows",
    pageInfo: {
      kind: "infinite",
      ...(nextCursor === undefined ? {} : { nextCursor }),
      loadedCount: 1,
      total: { kind: "unknown" }
    }
  };
}
