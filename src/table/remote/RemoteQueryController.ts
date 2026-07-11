import { serializeQueryRequest, type QueryRequest, type QueryResult, type QueryRow } from "../core/query";
import { RemotePageCache, RemotePageCacheError } from "./RemotePageCache";
import type { RemoteTableSource } from "./types";

export type RemoteQuerySnapshot<TRow> = {
  status: "idle" | "loading" | "ready" | "error";
  items: readonly QueryRow<TRow>[];
  revision: string | null;
  completeness: QueryResult<TRow>["completeness"];
  pageInfo: QueryResult<TRow>["pageInfo"];
  error?: { code: string; message: string; retryable: boolean };
};

export type RemoteQueryAcceptance<TRow> = {
  generation: number;
  revision: string;
  items: readonly QueryRow<TRow>[];
  completeness: QueryResult<TRow>["completeness"];
};

export type RemoteTableErrorCode =
  | "NO_QUERY"
  | "REMOTE_SESSION_DESTROYED"
  | "UNSUPPORTED_PAGINATION";

export class RemoteTableError extends Error {
  constructor(readonly code: RemoteTableErrorCode, message: string = code) {
    super(message);
    this.name = "RemoteTableError";
  }
}

export class RemoteQueryController<TRow> {
  private generation = 0;
  private abortController: AbortController | null = null;
  private snapshot: RemoteQuerySnapshot<TRow> = emptySnapshot();
  private readonly listeners = new Set<() => void>();
  private readonly acceptanceListeners = new Set<(acceptance: RemoteQueryAcceptance<TRow>) => void>();
  private readonly cache: RemotePageCache<TRow>;
  private destroyed = false;
  private lastQuery: QueryRequest | null = null;
  private queryFamily: string | null = null;
  private currentRevision: string | null = null;

  constructor(
    private readonly source: RemoteTableSource<TRow>,
    options: { maxCachedPages?: number } = {}
  ) {
    this.cache = new RemotePageCache(options.maxCachedPages);
  }

  async load(query: QueryRequest, operationId: string): Promise<void> {
    this.assertActive();
    if (query.pagination.kind !== this.source.paginationMode) {
      throw new RemoteTableError(
        "UNSUPPORTED_PAGINATION",
        `Pagination ${query.pagination.kind} is unsupported by ${this.source.paginationMode} source`
      );
    }

    const family = queryFamilyKey(query);
    if (family !== this.queryFamily) {
      this.cache.clear();
      this.queryFamily = family;
    }
    this.lastQuery = query;
    const generation = ++this.generation;
    this.abortController?.abort();
    const abortController = new AbortController();
    this.abortController = abortController;
    const before = this.snapshot;
    this.publish({ ...before, status: "loading", error: undefined });

    try {
      const result = await this.source.query(query, {
        signal: abortController.signal,
        generation,
        operationId
      });
      if (!this.isCurrent(generation, abortController)) return;

      if (this.currentRevision !== null) {
        const order = this.source.compareRevisions(result.revision, this.currentRevision);
        if (order === "older") {
          this.publish({ ...before, status: before.revision === null ? "idle" : "ready", error: undefined });
          return;
        }
        if (order === "unknown") {
          this.cache.clear();
          this.publish({
            ...before,
            status: "error",
            items: [],
            error: {
              code: "REVISION_ORDER_UNKNOWN",
              message: "The source could not order this response; refresh required.",
              retryable: true
            }
          });
          return;
        }
      }

      if (query.pagination.kind === "offset" || query.pagination.kind === "none") {
        this.cache.clear();
      }
      this.cache.put(query, result);
      const combined = this.cache.combine() ?? result;
      this.currentRevision = result.revision;
      this.publishAcceptance({
        generation,
        revision: result.revision,
        items: result.items,
        completeness: result.completeness
      });
      this.publish({ status: "ready", ...combined, error: undefined });
    } catch (error) {
      if (!this.isCurrent(generation, abortController)) return;
      if (error instanceof RemotePageCacheError) this.cache.clear();
      this.publish({ ...before, status: "error", error: normalizeRemoteError(error) });
    }
  }

  async refresh(operationId: string): Promise<void> {
    this.assertActive();
    if (!this.lastQuery) throw new RemoteTableError("NO_QUERY");
    await this.load(this.lastQuery, operationId);
  }

  getSnapshot = (): RemoteQuerySnapshot<TRow> => this.snapshot;

  getDiagnostics(): {
    generation: number;
    cachedPages: number;
    cachedItems: number;
    destroyed: boolean;
  } {
    return {
      generation: this.generation,
      cachedPages: this.cache.size,
      cachedItems: this.cache.combine()?.items.length ?? 0,
      destroyed: this.destroyed
    };
  }

  subscribe = (listener: () => void): (() => void) => {
    if (this.destroyed) return () => {};
    this.listeners.add(listener);
    return () => this.listeners.delete(listener);
  };

  subscribeAccepted = (
    listener: (acceptance: RemoteQueryAcceptance<TRow>) => void
  ): (() => void) => {
    if (this.destroyed) return () => {};
    this.acceptanceListeners.add(listener);
    return () => this.acceptanceListeners.delete(listener);
  };

  getCanonicalRow(rowId: string): TRow | undefined {
    return this.cache.getCanonicalRow(rowId);
  }

  getCurrentRevision(): string | null {
    return this.currentRevision;
  }

  applyCanonicalRows(rows: readonly TRow[], revision: string, publish = true): void {
    if (this.destroyed) return;
    this.cache.applyCanonicalRows(rows, revision, this.source.getRowId);
    this.currentRevision = revision;
    if (publish) this.publishFromCache();
  }

  deleteCanonicalRows(rowIds: readonly string[], revision: string, publish = true): void {
    if (this.destroyed) return;
    this.cache.deleteCanonicalRows(rowIds, revision);
    this.currentRevision = revision;
    if (publish) this.publishFromCache();
  }

  noteCurrentRevision(revision: string, publish = true): void {
    if (this.destroyed) return;
    this.cache.noteRevision(revision);
    this.currentRevision = revision;
    if (publish) this.publishFromCache();
  }

  publishCanonicalState(): void {
    if (!this.destroyed) this.publishFromCache();
  }

  invalidate(): void {
    if (this.destroyed) return;
    this.generation += 1;
    this.abortController?.abort();
    this.abortController = null;
    this.cache.clear();
    this.publish({
      ...emptySnapshot<TRow>(),
      status: "error",
      revision: this.currentRevision,
      error: {
        code: "REMOTE_INVALIDATED",
        message: "Remote data was invalidated; refresh required.",
        retryable: true
      }
    });
  }

  destroy(): void {
    if (this.destroyed) return;
    this.destroyed = true;
    this.generation += 1;
    this.abortController?.abort();
    this.abortController = null;
    this.cache.clear();
    this.listeners.clear();
    this.acceptanceListeners.clear();
  }

  private publishFromCache(): void {
    const combined = this.cache.combine();
    if (!combined) return;
    this.publish({ status: "ready", ...combined, error: undefined });
  }

  private publish(next: RemoteQuerySnapshot<TRow>): void {
    if (this.destroyed) return;
    this.snapshot = next;
    for (const listener of [...this.listeners]) {
      try { listener(); } catch { /* Host listeners are isolated. */ }
    }
  }

  private publishAcceptance(acceptance: RemoteQueryAcceptance<TRow>): void {
    if (this.destroyed) return;
    for (const listener of [...this.acceptanceListeners]) {
      try { listener(acceptance); } catch { /* Reconciliation listeners are isolated. */ }
    }
  }

  private assertActive(): void {
    if (this.destroyed) throw new RemoteTableError("REMOTE_SESSION_DESTROYED");
  }

  private isCurrent(generation: number, abortController: AbortController): boolean {
    return !this.destroyed
      && generation === this.generation
      && !abortController.signal.aborted;
  }
}

export function normalizeRemoteError(error: unknown): {
  code: string;
  message: string;
  retryable: boolean;
} {
  if (error && typeof error === "object") {
    const candidate = error as { code?: unknown; message?: unknown; retryable?: unknown };
    return {
      code: typeof candidate.code === "string" ? candidate.code : "REMOTE_QUERY_FAILED",
      message: typeof candidate.message === "string" ? candidate.message : "Remote query failed",
      retryable: typeof candidate.retryable === "boolean" ? candidate.retryable : true
    };
  }
  return {
    code: "REMOTE_QUERY_FAILED",
    message: typeof error === "string" ? error : "Remote query failed",
    retryable: true
  };
}

function emptySnapshot<TRow>(): RemoteQuerySnapshot<TRow> {
  return {
    status: "idle",
    items: [],
    revision: null,
    completeness: "loadedRows",
    pageInfo: { kind: "none", total: { kind: "known", value: 0 } }
  };
}

function queryFamilyKey(query: QueryRequest): string {
  switch (query.pagination.kind) {
    case "none":
      return serializeQueryRequest(query);
    case "offset":
      return serializeQueryRequest({ ...query, pagination: { ...query.pagination, offset: 0 } });
    case "cursor": {
      const { cursor: _cursor, ...pagination } = query.pagination;
      return serializeQueryRequest({ ...query, pagination });
    }
    case "infinite": {
      const { after: _after, ...pagination } = query.pagination;
      return serializeQueryRequest({ ...query, pagination });
    }
  }
}
