import type { TableCapabilities } from "../core/capabilities";
import type { QueryRequest, QueryResult } from "../core/query";
import type {
  ExportArtifact,
  ExportOptions,
  TableAbortSignal,
  TableCellMetadata
} from "../core/types";

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
  | {
      clientMutationId: string;
      status: "rejected";
      revision: string;
      issues: readonly { code: string; message: string; columnId?: string }[];
    }
  | {
      clientMutationId: string;
      status: "conflict";
      revision: string;
      current: TRow;
      rowVersion?: string;
    };

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
  query(
    request: QueryRequest,
    context: { signal: TableAbortSignal; generation: number; operationId: string }
  ): Promise<QueryResult<TRow>>;
  mutate?(
    batch: readonly RemoteMutation[],
    context: { signal: TableAbortSignal; operationId: string }
  ): Promise<readonly RemoteMutationResult<TRow>[]>;
  export?(
    request: RemoteExportRequest,
    context: { signal: TableAbortSignal; operationId: string }
  ): Promise<ExportArtifact>;
  subscribe?(listener: (event: RemoteSourceEvent<TRow>) => void): () => void;
};

export type CreateRemoteTableSourceOptions<TRow> = Omit<RemoteTableSource<TRow>, "kind">;
