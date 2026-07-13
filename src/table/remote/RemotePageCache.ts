import type { PaginationRequest, QueryRequest, QueryResult, QueryRow } from "../core/query";
import type { TablePageGap } from "../core/types";

const DEFAULT_MAX_PAGES = 5;

type CacheEntry<TRow> = {
  key: string;
  request: QueryRequest;
  result: QueryResult<TRow>;
  order: number;
  lastUsed: number;
  visibleIds: readonly string[];
};

export type RemotePageCachePutResult =
  | { gapCreated: false }
  | {
      gapCreated: true;
      mode: "cursor" | "infinite";
      maxPages: number;
      gap: TablePageGap;
    };

export type RemotePageRecoveryWindow = {
  headQuery: QueryRequest;
  pageCount: number;
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
  private readonly visibleIdOwners = new Map<string, string>();
  private clock = 0;
  private order = 0;
  private pageGap: TablePageGap | null = null;
  private rootQuery: QueryRequest | null = null;
  private currentRevision: string | null = null;

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

  getPageGaps(): readonly TablePageGap[] {
    return this.pageGap ? [this.pageGap] : [];
  }

  getRecoveryWindow(): RemotePageRecoveryWindow | null {
    if (!this.rootQuery) return null;
    return {
      headQuery: this.rootQuery,
      pageCount: this.entries.size + (this.pageGap?.omittedPages ?? 0)
    };
  }

  put(request: QueryRequest, result: QueryResult<TRow>): RemotePageCachePutResult {
    const key = getRemotePageKey(request.pagination);
    const existing = this.entries.get(key);
    let visibleIds: readonly string[];
    try {
      visibleIds = collectVisibleIds(result.items);
      for (const id of visibleIds) {
        const owner = this.visibleIdOwners.get(id);
        if (owner !== undefined && owner !== key) {
          throw new RemotePageCacheError(`Duplicate remote row ID: ${id}`);
        }
      }
    } catch (error) {
      this.clear();
      throw error;
    }

    if (existing) this.releaseVisibleIds(existing);
    const candidate: CacheEntry<TRow> = {
      key,
      request,
      result,
      order: existing?.order ?? ++this.order,
      lastUsed: ++this.clock,
      visibleIds
    };
    this.entries.set(key, candidate);
    for (const id of visibleIds) this.visibleIdOwners.set(id, key);
    this.currentRevision = result.revision;

    const accumulatedMode = request.pagination.kind === "cursor"
      || request.pagination.kind === "infinite"
      ? request.pagination.kind
      : null;
    if (accumulatedMode) {
      if (this.rootQuery === null) {
        this.rootQuery = request;
      }
      const gapCreated = this.pageGap === null && this.entries.size > this.maxPages;
      while (this.entries.size > this.maxPages) {
        const oldest = this.orderedEntries()[0];
        this.evict(oldest);
        this.pageGap = {
          kind: "evicted-pages",
          at: 0,
          omittedPages: (this.pageGap?.omittedPages ?? 0) + 1,
          omittedItems: (this.pageGap?.omittedItems ?? 0)
            + oldest.result.items.filter((item) => item.kind === "data").length
        };
      }
      if (gapCreated && this.pageGap) {
        return {
          gapCreated: true,
          mode: accumulatedMode,
          maxPages: this.maxPages,
          gap: this.pageGap
        };
      }
    } else {
      while (this.entries.size > this.maxPages) {
        const oldest = [...this.entries.values()].reduce((left, right) =>
          left.lastUsed <= right.lastUsed ? left : right
        );
        this.evict(oldest);
      }
    }
    return { gapCreated: false };
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
      revision: this.currentRevision ?? latest.revision,
      completeness: entries.every((entry) => entry.result.completeness === "completeDataset")
        ? "completeDataset"
        : "loadedRows",
      pageInfo
    };
  }

  getCanonicalRow(rowId: string): TRow | undefined {
    const owner = this.visibleIdOwners.get(rowId);
    const entry = owner === undefined ? undefined : this.entries.get(owner);
    const item = entry?.result.items.find((candidate) => candidate.kind === "data" && candidate.id === rowId);
    return item?.kind === "data" ? item.original : undefined;
  }

  applyCanonicalRows(
    rows: readonly TRow[],
    revision: string,
    getRowId: (row: TRow) => string
  ): void {
    const replacementsByPage = new Map<string, Map<string, TRow>>();
    for (const row of rows) {
      const rowId = getRowId(row);
      const owner = this.visibleIdOwners.get(rowId);
      if (!owner) continue;
      const replacements = replacementsByPage.get(owner) ?? new Map<string, TRow>();
      replacements.set(rowId, row);
      replacementsByPage.set(owner, replacements);
    }
    for (const [key, replacements] of replacementsByPage) {
      const entry = this.entries.get(key);
      if (!entry) continue;
      entry.result = {
        ...entry.result,
        items: entry.result.items.map((item): QueryRow<TRow> => {
          if (item.kind !== "data") return item;
          const replacement = replacements.get(item.id);
          return replacement === undefined ? item : { ...item, original: replacement };
        })
      };
    }
    this.currentRevision = revision;
  }

  deleteCanonicalRows(rowIds: readonly string[], revision: string): void {
    const deletedByPage = new Map<string, Set<string>>();
    for (const rowId of rowIds) {
      const owner = this.visibleIdOwners.get(rowId);
      if (!owner) continue;
      const deleted = deletedByPage.get(owner) ?? new Set<string>();
      deleted.add(rowId);
      deletedByPage.set(owner, deleted);
    }
    for (const [key, deleted] of deletedByPage) {
      const entry = this.entries.get(key);
      if (!entry) continue;
      entry.result = {
        ...entry.result,
        items: entry.result.items.filter((item) => !deleted.has(item.id))
      };
      entry.visibleIds = entry.visibleIds.filter((id) => !deleted.has(id));
      for (const id of deleted) this.visibleIdOwners.delete(id);
    }
    this.currentRevision = revision;
  }

  noteRevision(revision: string): void {
    this.currentRevision = revision;
  }

  clear(): void {
    this.entries.clear();
    this.visibleIdOwners.clear();
    this.clock = 0;
    this.order = 0;
    this.pageGap = null;
    this.rootQuery = null;
    this.currentRevision = null;
  }

  private orderedEntries(): CacheEntry<TRow>[] {
    return [...this.entries.values()].sort((left, right) => left.order - right.order);
  }

  private evict(entry: CacheEntry<TRow>): void {
    this.entries.delete(entry.key);
    this.releaseVisibleIds(entry);
  }

  private releaseVisibleIds(entry: CacheEntry<TRow>): void {
    for (const id of entry.visibleIds) {
      if (this.visibleIdOwners.get(id) === entry.key) this.visibleIdOwners.delete(id);
    }
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

function collectVisibleIds<TRow>(items: readonly QueryRow<TRow>[]): readonly string[] {
  const ids = new Set<string>();
  for (const item of items) {
    const id = item.id;
    if (ids.has(id)) throw new RemotePageCacheError(`Duplicate remote row ID: ${id}`);
    ids.add(id);
  }
  return [...ids];
}
