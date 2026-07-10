import { describe, expect, it, vi } from "vitest";
import type { CommandResult } from "../../core/commands/types";
import type { TableCapabilities, TableFeature } from "../core/capabilities";
import type { QueryRequest, QueryResult } from "../core/query";
import type { ColumnDef } from "../core/types";
import { defineTableSessionContract } from "../core/session.contract";
import { createRemoteTableSession } from "./RemoteTableSession";
import { createTestRemoteSource } from "./testUtils";

type Row = { id: string; name: string; amount: number };

const row: Row = { id: "row-1", name: "Ada", amount: 10 };
const columns: readonly ColumnDef<Row>[] = [
  { id: "name", header: "Name", accessor: (value) => value.name },
  { id: "amount", header: "Amount", dataType: "number", accessor: (value) => value.amount }
];

defineTableSessionContract("remote read-only", () => {
  const source = createTestRemoteSource<Row>({
    capabilities: readOnlyCapabilities(),
    paginationMode: "none",
    mutationMode: "none",
    undoMode: "none",
    query: async (request) => result(request),
    export: async () => ({
      bytes: new Uint8Array([1]),
      mediaType: "text/csv",
      fileName: "rows.csv"
    })
  });
  const session = createRemoteTableSession({ source, columns });
  session.start();
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

function result(request: QueryRequest): QueryResult<Row> {
  if (request.pagination.kind !== "none") throw new Error("expected unpaginated query");
  return {
    items: [{ kind: "data", id: row.id, original: row, depth: 0 }],
    revision: "r1",
    completeness: "completeDataset",
    pageInfo: { kind: "none", total: { kind: "known", value: 1 } }
  };
}
