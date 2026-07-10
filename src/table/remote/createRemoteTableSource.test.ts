import { describe, expect, it, vi } from "vitest";
import type { QueryRequest, QueryResult } from "../core/query";
import type { TableCapabilities } from "../core/capabilities";
import { createRemoteTableSource } from "./createRemoteTableSource";
import type { CreateRemoteTableSourceOptions } from "./types";

type Row = { id: string; name: string };

describe("createRemoteTableSource", () => {
  it("requires stable ids and preserves declared dataset scope", async () => {
    const query = vi.fn(async () => offsetResult("r1", [{ id: "employee-1", name: "Ada" }]));
    const source = createRemoteTableSource(baseOptions({ query }));

    const result = await source.query(offsetRequest(), {
      signal: new AbortController().signal,
      generation: 1,
      operationId: "query-test-1"
    });

    expect(source.kind).toBe("remote");
    const first = result.items[0];
    expect(first.kind).toBe("data");
    if (first.kind !== "data") throw new Error("expected data row");
    expect(source.getRowId(first.original)).toBe("employee-1");
    expect(result.completeness).toBe("loadedRows");
    expect(source.capabilities.sort).toEqual({ executor: "server", scope: "completeDataset" });
    expect(query).toHaveBeenCalledTimes(1);
  });

  it.each([
    ["blank IDs", offsetResult("r1", [{ id: "", name: "Ada" }])],
    ["duplicate IDs", {
      ...offsetResult("r1", [{ id: "employee-1", name: "Ada" }]),
      items: [
        { kind: "data" as const, id: "employee-1", original: { id: "employee-1", name: "Ada" }, depth: 0 },
        { kind: "data" as const, id: "employee-1", original: { id: "employee-1", name: "Grace" }, depth: 0 }
      ]
    }],
    ["row/source ID mismatches", {
      ...offsetResult("r1", [{ id: "employee-1", name: "Ada" }]),
      items: [{
        kind: "data" as const,
        id: "different-id",
        original: { id: "employee-1", name: "Ada" },
        depth: 0
      }]
    }],
    ["page kind mismatches", {
      ...offsetResult("r1", [{ id: "employee-1", name: "Ada" }]),
      pageInfo: { kind: "cursor" as const, total: { kind: "known" as const, value: 1 } }
    }],
    ["paginated complete-dataset claims", {
      ...offsetResult("r1", [{ id: "employee-1", name: "Ada" }]),
      completeness: "completeDataset" as const
    }]
  ] satisfies ReadonlyArray<readonly [string, QueryResult<Row>]>) (
    "rejects %s returned by the host",
    async (_label, result) => {
      const source = createRemoteTableSource(baseOptions({ query: async () => result }));
      await expect(source.query(offsetRequest(), queryContext())).rejects.toThrow();
    }
  );

  it("rejects a request pagination mode that contradicts the source", async () => {
    const query = vi.fn(async () => offsetResult("r1", []));
    const source = createRemoteTableSource(baseOptions({ query }));

    await expect(source.query({
      ...offsetRequest(),
      pagination: { kind: "cursor", limit: 50 }
    }, queryContext())).rejects.toThrow(/pagination/i);
    expect(query).not.toHaveBeenCalled();
  });

  it.each([
    ["pagination mode absent from capabilities", () => baseOptions({ paginationMode: "cursor" })],
    ["edit without mutate", () => baseOptions({
      capabilities: { ...baseCapabilities(), edit: { executor: "server", scope: "completeDataset" } }
    })],
    ["bulk edit without mutate", () => baseOptions({
      capabilities: { ...baseCapabilities(), bulkEdit: { executor: "server", scope: "completeDataset" } }
    })],
    ["metadata without mutate", () => baseOptions({
      capabilities: { ...baseCapabilities(), metadata: { executor: "server", scope: "completeDataset" } },
      readCell: () => ({ storedValue: null, evaluatedValue: null })
    })],
    ["metadata without readCell", () => baseOptions({
      capabilities: { ...baseCapabilities(), metadata: { executor: "server", scope: "completeDataset" } },
      mutate: async () => []
    })],
    ["write capabilities in mutation none mode", () => baseOptions({
      capabilities: { ...baseCapabilities(), edit: { executor: "server", scope: "completeDataset" } },
      mutationMode: "none",
      mutate: async () => []
    })],
    ["server validation without mutate", () => baseOptions({
      capabilities: { ...baseCapabilities(), validation: { executor: "server", scope: "completeDataset" } }
    })],
    ["server formulas without mutate", () => baseOptions({
      capabilities: { ...baseCapabilities(), formula: "server" },
      readCell: () => ({ storedValue: null, evaluatedValue: null })
    })],
    ["server formulas without readCell", () => baseOptions({
      capabilities: { ...baseCapabilities(), formula: "server" },
      mutate: async () => []
    })],
    ["export without an exporter", () => baseOptions({
      capabilities: { ...baseCapabilities(), export: { executor: "server", scope: "completeDataset" } }
    })],
    ["subscriptions without subscribe", () => baseOptions({
      capabilities: { ...baseCapabilities(), subscription: true }
    })],
    ["compensating undo without versioned mutation support", () => baseOptions({
      capabilities: { ...baseCapabilities(), undo: { executor: "server", scope: "completeDataset" } },
      undoMode: "compensating",
      mutationMode: "none"
    })]
  ] satisfies ReadonlyArray<readonly [string, () => CreateRemoteTableSourceOptions<Row>]>) (
    "rejects contradictory declaration: %s",
    (_label, createOptions) => {
      expect(() => createRemoteTableSource(createOptions())).toThrow();
    }
  );
});

function baseOptions(
  overrides: Partial<CreateRemoteTableSourceOptions<Row>> = {}
): CreateRemoteTableSourceOptions<Row> {
  return {
    getRowId: (row) => row.id,
    capabilities: baseCapabilities(),
    paginationMode: "offset",
    mutationMode: "versioned",
    undoMode: "none",
    compareRevisions: (candidate, current) => candidate === current ? "equal" : "unknown",
    query: async () => offsetResult("r1", []),
    ...overrides
  };
}

function baseCapabilities(): TableCapabilities {
  return {
    sort: { executor: "server", scope: "completeDataset" },
    filter: { executor: "server", scope: "completeDataset" },
    group: false,
    aggregate: false,
    pagination: { executor: "server", scope: "completeDataset", modes: ["offset"] },
    edit: false,
    bulkEdit: false,
    metadata: false,
    validation: false,
    undo: false,
    export: false,
    formula: "none",
    subscription: false
  };
}

function offsetRequest(): QueryRequest {
  return {
    sorting: [],
    filter: null,
    grouping: [],
    aggregates: [],
    pagination: { kind: "offset", offset: 0, limit: 50 }
  };
}

function offsetResult(revision: string, rows: readonly Row[]): QueryResult<Row> {
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

function queryContext() {
  return {
    signal: new AbortController().signal,
    generation: 1,
    operationId: "query-test"
  };
}
