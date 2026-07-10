import type { TableCapabilities } from "../core/capabilities";
import type { QueryResult } from "../core/query";
import { createRemoteTableSource } from "./createRemoteTableSource";
import type {
  CreateRemoteTableSourceOptions,
  RemoteTableSource,
  RevisionComparison
} from "./types";

export function createDeferred<T>() {
  let resolve!: (value: T) => void;
  let reject!: (reason?: unknown) => void;
  const promise = new Promise<T>((resolvePromise, rejectPromise) => {
    resolve = resolvePromise;
    reject = rejectPromise;
  });
  return { promise, resolve, reject };
}

export function createTestRemoteSource<TRow>(
  overrides: Partial<CreateRemoteTableSourceOptions<TRow>> = {}
): RemoteTableSource<TRow> {
  const capabilities = overrides.capabilities ?? defaultCapabilities();
  const paginationMode = overrides.paginationMode ?? "offset";
  const options: CreateRemoteTableSourceOptions<TRow> = {
    getRowId: (row) => String((row as { id?: unknown }).id ?? ""),
    mutationMode: "versioned",
    undoMode: capabilities.undo ? "compensating" : "none",
    compareRevisions: defaultRevisionComparator,
    readCell: (row, columnId) => {
      const value = (row as Record<string, unknown>)[columnId];
      return { storedValue: value ?? null, evaluatedValue: value ?? null };
    },
    query: async () => emptyResult<TRow>(paginationMode),
    mutate: async () => [],
    subscribe: () => () => {},
    ...overrides,
    capabilities,
    paginationMode
  };
  return createRemoteTableSource(options);
}

export function defaultRemoteCapabilities(): TableCapabilities {
  return defaultCapabilities();
}

export function defaultRevisionComparator(
  candidate: string,
  current: string
): RevisionComparison {
  return candidate === current ? "equal" : "unknown";
}

function defaultCapabilities(): TableCapabilities {
  const complete = { executor: "server" as const, scope: "completeDataset" as const };
  return {
    sort: { ...complete },
    filter: { ...complete },
    group: false,
    aggregate: false,
    pagination: { ...complete, modes: ["offset"] },
    edit: { ...complete },
    bulkEdit: { ...complete },
    metadata: { ...complete },
    validation: { ...complete },
    formula: "server",
    subscription: true,
    undo: { ...complete },
    export: false
  };
}

function emptyResult<TRow>(
  mode: CreateRemoteTableSourceOptions<TRow>["paginationMode"]
): QueryResult<TRow> {
  const common = { items: [], revision: "test-revision", completeness: "loadedRows" as const };
  switch (mode) {
    case "none":
      return { ...common, completeness: "completeDataset", pageInfo: { kind: "none", total: { kind: "known", value: 0 } } };
    case "offset":
      return {
        ...common,
        pageInfo: {
          kind: "offset", offset: 0, limit: 50,
          total: { kind: "known", value: 0 }, hasMore: false
        }
      };
    case "cursor":
      return { ...common, pageInfo: { kind: "cursor", total: { kind: "unknown" } } };
    case "infinite":
      return {
        ...common,
        pageInfo: { kind: "infinite", loadedCount: 0, total: { kind: "unknown" } }
      };
  }
}
