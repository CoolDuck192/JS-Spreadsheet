import { describe, expect, it, vi } from "vitest";
import type { CommandResult } from "../../core/commands/types";
import type { TableCapabilities, TableFeature } from "../core/capabilities";
import type { QueryRequest, QueryResult } from "../core/query";
import type { ColumnDef } from "../core/types";
import { defineTableSessionContract } from "../core/session.contract";
import { createRemoteTableSession } from "./RemoteTableSession";
import { createTestRemoteSource } from "./testUtils";
import type { RemoteMutation } from "./types";

type Row = { id: string; name: string; amount: number };

const row: Row = { id: "row-1", name: "Ada", amount: 10 };
const columns: readonly ColumnDef<Row>[] = [
  {
    id: "name", header: "Name", accessor: (value) => value.name,
    update: (value, name) => ({ ...value, name: String(name) })
  },
  {
    id: "amount", header: "Amount", dataType: "number", accessor: (value) => value.amount,
    update: (value, amount) => ({ ...value, amount: Number(amount) })
  }
];

defineTableSessionContract("remote", async () => {
  let queryRevision = "r1";
  let queryRow = row;
  let mutationCalls = 0;
  const source = createTestRemoteSource<Row>({
    capabilities: mutableCapabilities(),
    paginationMode: "none",
    mutationMode: "versioned",
    undoMode: "none",
    compareRevisions: (candidate, current) => candidate === current
      ? "equal"
      : Number(candidate.slice(1)) > Number(current.slice(1)) ? "newer" : "older",
    query: async (request) => result(request, queryRevision, queryRow),
    mutate: async (batch: readonly RemoteMutation[]) => {
      mutationCalls += 1;
      if (mutationCalls === 1) {
        queryRevision = "r2";
        queryRow = { ...row, name: "Authoritative" };
        return batch.map((mutation) => ({
          clientMutationId: mutation.clientMutationId,
          status: "conflict" as const,
          revision: queryRevision,
          current: queryRow
        }));
      }
      return new Promise(() => undefined);
    },
    export: async () => ({
      bytes: new Uint8Array([1]),
      mediaType: "text/csv",
      fileName: "rows.csv"
    })
  });
  const session = createRemoteTableSession({ source, columns });
  session.start();
  await vi.waitFor(() => expect(session.getSnapshot().status.phase).toBe("ready"));
  await session.dispatch({
    type: "edit-cells",
    edits: [{ rowId: "row-1", columnId: "name", rawText: "Attempted" }]
  });
  await vi.waitFor(() => expect(session.getSnapshot().conflicts).toHaveLength(1));
  const conflict = session.getSnapshot().conflicts[0];
  const featureOperations: Partial<Record<TableFeature, () => Promise<CommandResult>>> = {
    sort: () => session.dispatch({ type: "set-sorting", sorting: [{ columnId: "amount", direction: "desc" }] }),
    filter: () => session.dispatch({
      type: "set-filter",
      filter: { kind: "blank", columnId: "name", operator: "isNotBlank" }
    }),
    group: () => session.dispatch({ type: "set-grouping", grouping: [{ columnId: "name" }] }),
    aggregate: () => session.dispatch({
      type: "set-aggregates",
      aggregates: [{ id: "sum", columnId: "amount", function: "sum" }]
    }),
    pagination: () => session.dispatch({
      type: "set-pagination",
      pagination: { kind: "none" }
    }),
    edit: () => session.dispatch({
      type: "edit-cells",
      edits: [{ rowId: "row-1", columnId: "name", rawText: "Grace" }]
    }),
    bulkEdit: () => session.dispatch({
      type: "edit-cells",
      edits: [
        { rowId: "row-1", columnId: "name", rawText: "Grace" },
        { rowId: "row-1", columnId: "amount", rawText: "20" }
      ]
    }),
    metadata: () => session.dispatch({
      type: "update-cell-metadata",
      updates: [{ rowId: "row-1", columnId: "name", patch: { comment: "note" } }]
    }),
    validation: () => session.dispatch({
      type: "edit-cells",
      edits: [{ rowId: "row-1", columnId: "amount", rawText: "20" }]
    }),
    export: async () => {
      await vi.waitFor(() => expect(session.getSnapshot().status.phase).toBe("ready"));
      await session.export({ format: "csv", scope: "completeDataset" });
      return { status: "committed", revision: session.getSnapshot().revision, changed: false };
    }
  };
  return {
    session,
    editableCell: { rowId: "row-1", columnId: "name" },
    validRawText: "Grace",
    featureOperations,
    conflictResolution: {
      reload: {
        type: "reload-authoritative",
        operationId: conflict.operationId,
        rowId: conflict.rowId
      },
      retry: {
        type: "retry-with-revision",
        operationId: conflict.operationId,
        rowId: conflict.rowId,
        expectedRevision: conflict.revision
      }
    },
    cleanup: () => session.destroy()
  };
});

describe("remote read-only capability boundary", () => {
  it("rejects absent write features without calling mutate", async () => {
    const mutate = vi.fn(async () => []);
    const source = createTestRemoteSource<Row>({
      capabilities: readOnlyCapabilities(),
      paginationMode: "none",
      mutationMode: "none",
      undoMode: "none",
      mutate,
      query: async (request) => result(request),
      export: async () => ({ bytes: new Uint8Array([1]), mediaType: "text/csv", fileName: "rows.csv" })
    });
    const session = createRemoteTableSession({ source, columns });

    expect(await session.dispatch({
      type: "edit-cells",
      edits: [{ rowId: "row-1", columnId: "name", rawText: "Grace" }]
    })).toMatchObject({ status: "rejected", reason: "unsupported" });
    expect(await session.dispatch({
      type: "update-cell-metadata",
      updates: [{ rowId: "row-1", columnId: "name", patch: { comment: "note" } }]
    })).toMatchObject({ status: "rejected", reason: "unsupported" });
    expect(await session.undo()).toMatchObject({ status: "rejected", reason: "unsupported" });
    expect(mutate).not.toHaveBeenCalled();

    session.destroy();
  });
});

function readOnlyCapabilities(): TableCapabilities {
  const complete = { executor: "server" as const, scope: "completeDataset" as const };
  return {
    sort: { ...complete }, filter: { ...complete }, group: { ...complete }, aggregate: { ...complete },
    pagination: { ...complete, modes: ["none"] },
    edit: false, bulkEdit: false, metadata: false, validation: false,
    formula: "none", subscription: false, undo: false, export: { ...complete }
  };
}

function mutableCapabilities(): TableCapabilities {
  const capabilities = readOnlyCapabilities();
  const complete = { executor: "server" as const, scope: "completeDataset" as const };
  return {
    ...capabilities,
    edit: { ...complete },
    bulkEdit: { ...complete },
    metadata: { ...complete },
    validation: { ...complete },
    formula: "server"
  };
}

function result(
  request: QueryRequest,
  revision: string = "r1",
  currentRow: Row = row
): QueryResult<Row> {
  if (request.pagination.kind !== "none") throw new Error("expected unpaginated query");
  return {
    items: [{ kind: "data", id: currentRow.id, original: currentRow, depth: 0 }],
    revision,
    completeness: "completeDataset",
    pageInfo: { kind: "none", total: { kind: "known", value: 1 } }
  };
}
