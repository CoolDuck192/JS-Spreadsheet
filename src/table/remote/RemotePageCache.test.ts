import { describe, expect, it } from "vitest";
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
