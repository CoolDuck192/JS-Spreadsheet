import type { PaginationRequest, QueryRequest, QueryResult, QueryRow } from "../core/query";

const DEFAULT_MAX_PAGES = 5;

type CacheEntry<TRow> = {
  key: string;
  request: QueryRequest;
  result: QueryResult<TRow>;
  order: number;
  lastUsed: number;
};

export class RemotePageCacheError extends Error {
  readonly code = "DUPLICATE_REMOTE_ROW_ID";

  constructor(message = "Remote pages contain duplicate visible row IDs") {
    super(message);
    this.name = "RemotePageCacheError";
  }
}

export class RemotePageCache<TRow> {
  private readonly entries = new Map<string, CacheEntry<TRow>>();
  private clock = 0;
  private order = 0;

  constructor(private readonly maxPages = DEFAULT_MAX_PAGES) {
    if (!Number.isInteger(maxPages) || maxPages < 1) {
      throw new RangeError("RemotePageCache maxPages must be a positive integer");
    }
  }

  get size(): number {
    return this.entries.size;
  }

  get(request: QueryRequest): QueryResult<TRow> | undefined {
    const entry = this.entries.get(getRemotePageKey(request.pagination));
    if (!entry) return undefined;
    entry.lastUsed = ++this.clock;
    return entry.result;
  }

  getRequests(): readonly QueryRequest[] {
    return this.orderedEntries().map((entry) => entry.request);
  }

  put(request: QueryRequest, result: QueryResult<TRow>): void {
    const key = getRemotePageKey(request.pagination);
    const existing = this.entries.get(key);
    const candidate: CacheEntry<TRow> = {
      key,
      request,
      result,
      order: existing?.order ?? ++this.order,
      lastUsed: ++this.clock
    };
    this.entries.set(key, candidate);

    const accumulatesVisiblePages = request.pagination.kind === "cursor"
      || request.pagination.kind === "infinite";
    if (!accumulatesVisiblePages) {
      while (this.entries.size > this.maxPages) {
        const oldest = [...this.entries.values()].reduce((left, right) =>
          left.lastUsed <= right.lastUsed ? left : right
        );
        this.entries.delete(oldest.key);
      }
    }

    try {
      validateUniqueVisibleIds(this.orderedEntries());
    } catch (error) {
      this.clear();
      throw error;
    }
  }

  combine(): QueryResult<TRow> | null {
    const entries = this.orderedEntries();
    if (entries.length === 0) return null;
    const items = entries.flatMap((entry) => entry.result.items);
    const latest = entries[entries.length - 1].result;
    const pageInfo = latest.pageInfo.kind === "infinite"
      ? { ...latest.pageInfo, loadedCount: items.length }
      : latest.pageInfo;
    return {
      items,
      revision: latest.revision,
      completeness: entries.every((entry) => entry.result.completeness === "completeDataset")
        ? "completeDataset"
        : "loadedRows",
      pageInfo
    };
  }

  getCanonicalRow(rowId: string): TRow | undefined {
    for (const entry of this.orderedEntries()) {
      const item = entry.result.items.find((candidate) => candidate.kind === "data" && candidate.id === rowId);
      if (item?.kind === "data") return item.original;
    }
    return undefined;
  }

  applyCanonicalRows(
    rows: readonly TRow[],
    revision: string,
    getRowId: (row: TRow) => string
  ): void {
    const replacements = new Map(rows.map((row) => [getRowId(row), row]));
    for (const entry of this.entries.values()) {
      entry.result = {
        ...entry.result,
        revision,
        items: entry.result.items.map((item): QueryRow<TRow> => {
          if (item.kind !== "data") return item;
          const replacement = replacements.get(item.id);
          return replacement === undefined ? item : { ...item, original: replacement };
        })
      };
    }
  }

  deleteCanonicalRows(rowIds: readonly string[], revision: string): void {
    const deleted = new Set(rowIds);
    for (const entry of this.entries.values()) {
      entry.result = {
        ...entry.result,
        revision,
        items: entry.result.items.filter((item) => !deleted.has(item.id))
      };
    }
  }

  noteRevision(revision: string): void {
    for (const entry of this.entries.values()) {
      entry.result = { ...entry.result, revision };
    }
  }

  clear(): void {
    this.entries.clear();
  }

  private orderedEntries(): CacheEntry<TRow>[] {
    return [...this.entries.values()].sort((left, right) => left.order - right.order);
  }
}

export function getRemotePageKey(pagination: PaginationRequest): string {
  switch (pagination.kind) {
    case "none":
      return "none";
    case "offset":
      return `offset:${pagination.offset}:${pagination.limit}`;
    case "cursor":
      return `cursor:${encodeURIComponent(pagination.cursor ?? "")}:${pagination.limit}`;
    case "infinite":
      return `infinite:${encodeURIComponent(pagination.after ?? "")}:${pagination.limit}`;
  }
}

function validateUniqueVisibleIds<TRow>(entries: readonly CacheEntry<TRow>[]): void {
  const ids = new Set<string>();
  for (const entry of entries) {
    for (const item of entry.result.items) {
      if (ids.has(item.id)) throw new RemotePageCacheError(`Duplicate remote row ID: ${item.id}`);
      ids.add(item.id);
    }
  }
}
