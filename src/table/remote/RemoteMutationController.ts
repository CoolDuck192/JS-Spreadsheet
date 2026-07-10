import type { CommandResult } from "../../core/commands/types";
import type { QueryRequest } from "../core/query";
import type {
  TableCellIssue,
  TableConflict,
  TablePendingOperation
} from "../core/types";
import type { OptimisticCell, OptimisticOverlayStore } from "./OptimisticOverlayStore";
import type { RemoteQueryController } from "./RemoteQueryController";
import type {
  RemoteMutation,
  RemoteMutationResult,
  RemoteTableSource
} from "./types";

export type PreparedRemoteMutation = {
  rowId: string;
  columnId: string;
  rowVersion?: string;
  optimisticCell: OptimisticCell;
} & (
  | { kind: "cell-value"; rawText: string; parsedValue: unknown; formula?: string }
  | { kind: "cell-metadata"; metadata: OptimisticCell["metadata"] }
);

export type RemoteMutationControllerOptions<TRow> = {
  source: RemoteTableSource<TRow>;
  queryController: RemoteQueryController<TRow>;
  overlays: OptimisticOverlayStore;
  getActiveQuery(): QueryRequest;
  onChange(): void;
  limits?: { maxPendingOperations?: number; maxPendingCells?: number };
};

type MutationBatch = {
  operationId: string;
  mutationIds: readonly string[];
  cells: readonly { rowId: string; columnId: string }[];
  feature: "edit" | "metadata";
  startedAt: number;
  abortController: AbortController;
  status: "pending" | "uncertain";
};

const DEFAULT_MAX_PENDING_OPERATIONS = 100;
const DEFAULT_MAX_PENDING_CELLS = 1_000;

export class RemoteMutationController<TRow> {
  private readonly source: RemoteTableSource<TRow>;
  private readonly queryController: RemoteQueryController<TRow>;
  private readonly overlays: OptimisticOverlayStore;
  private readonly getActiveQuery: () => QueryRequest;
  private readonly onChange: () => void;
  private readonly maxPendingOperations: number;
  private readonly maxPendingCells: number;
  private readonly batches = new Map<string, MutationBatch>();
  private readonly conflicts = new Map<string, TableConflict<TRow>>();
  private issues: TableCellIssue[] = [];
  private destroyed = false;

  constructor(options: RemoteMutationControllerOptions<TRow>) {
    this.source = options.source;
    this.queryController = options.queryController;
    this.overlays = options.overlays;
    this.getActiveQuery = options.getActiveQuery;
    this.onChange = options.onChange;
    this.maxPendingOperations = positiveLimit(
      options.limits?.maxPendingOperations,
      DEFAULT_MAX_PENDING_OPERATIONS,
      "maxPendingOperations"
    );
    this.maxPendingCells = positiveLimit(
      options.limits?.maxPendingCells,
      DEFAULT_MAX_PENDING_CELLS,
      "maxPendingCells"
    );
  }

  async execute(
    operationId: string,
    prepared: readonly PreparedRemoteMutation[]
  ): Promise<CommandResult<TRow>> {
    const preflight = this.preflight(operationId, prepared);
    if (preflight) return preflight;
    this.issues = [];
    if (prepared.length === 0) {
      return {
        status: "committed",
        revision: this.queryController.getSnapshot().revision ?? "0",
        changed: false
      };
    }
    const revision = this.queryController.getCurrentRevision()!;

    const mutations: RemoteMutation[] = prepared.map((item, index) => {
      const clientMutationId = `${operationId}:${index}`;
      this.overlays.add({
        clientMutationId,
        operationId,
        rowId: item.rowId,
        columnId: item.columnId,
        baseRevision: revision,
        ...(item.rowVersion === undefined ? {} : { rowVersion: item.rowVersion }),
        cell: item.optimisticCell,
        status: "pending"
      });
      const base = {
        clientMutationId,
        rowId: item.rowId,
        columnId: item.columnId,
        baseRevision: revision,
        ...(item.rowVersion === undefined ? {} : { rowVersion: item.rowVersion })
      };
      return item.kind === "cell-value"
        ? {
            ...base,
            kind: "cell-value",
            rawText: item.rawText,
            parsedValue: item.parsedValue,
            ...(item.formula === undefined ? {} : { formula: item.formula })
          }
        : { ...base, kind: "cell-metadata", metadata: item.metadata };
    });
    const abortController = new AbortController();
    const batch: MutationBatch = {
      operationId,
      mutationIds: mutations.map((mutation) => mutation.clientMutationId),
      cells: prepared.map(({ rowId, columnId }) => ({ rowId, columnId })),
      feature: prepared.some((mutation) => mutation.kind === "cell-metadata") ? "metadata" : "edit",
      startedAt: Date.now(),
      abortController,
      status: "pending"
    };
    this.batches.set(operationId, batch);
    this.publish();

    let results: readonly RemoteMutationResult<TRow>[];
    try {
      results = await this.source.mutate!(mutations, {
        signal: abortController.signal,
        operationId
      });
    } catch (error) {
      if (this.destroyed || abortController.signal.aborted || isAbortError(error)) {
        this.removeBatch(batch, this.destroyed);
        return unsupported("Remote mutation was aborted");
      }
      for (const mutationId of batch.mutationIds) this.overlays.updateStatus(mutationId, "uncertain");
      batch.status = "uncertain";
      this.publish();
      return { status: "pending", operationId };
    }

    if (this.destroyed || abortController.signal.aborted) {
      this.removeBatch(batch, true);
      return unsupported("Remote mutation was aborted");
    }
    const protocolIssue = validateResults(batch.mutationIds, results);
    if (protocolIssue) {
      this.issues = [{ code: "REMOTE_MUTATION_PROTOCOL_ERROR", message: protocolIssue }];
      this.removeBatch(batch, true);
      this.queryController.invalidate();
      return validationIssue("REMOTE_MUTATION_PROTOCOL_ERROR", protocolIssue);
    }

    const byId = new Map(results.map((result) => [result.clientMutationId, result]));
    const ordered = batch.mutationIds.map((id) => byId.get(id)!);
    const issues: TableCellIssue[] = [];
    const committedRows: TRow[] = [];
    let conflictResult: Extract<RemoteMutationResult<TRow>, { status: "conflict" }> | null = null;
    let latestRevision = revision;
    let invalidated = false;

    for (const result of ordered) {
      const overlay = this.overlays.get(result.clientMutationId);
      if (!overlay) continue;
      const comparison = this.source.compareRevisions(
        result.revision,
        this.queryController.getCurrentRevision() ?? latestRevision
      );
      if (comparison === "older") {
        this.overlays.remove(result.clientMutationId);
        issues.push({
          code: "REMOTE_MUTATION_STALE_ACK",
          message: "The source returned an older mutation acknowledgement",
          rowId: overlay.rowId,
          columnId: overlay.columnId
        });
        invalidated = true;
        continue;
      }
      if (comparison === "unknown") {
        this.overlays.updateStatus(result.clientMutationId, "conflict");
        const current = authoritativeRow(result) ?? this.queryController.getCanonicalRow(overlay.rowId);
        if (current) {
          const conflict = createConflict(overlay, current, result.revision, this.source);
          this.conflicts.set(result.clientMutationId, conflict);
          if (result.status === "conflict") conflictResult = result;
        }
        invalidated = true;
        latestRevision = result.revision;
        continue;
      }

      latestRevision = result.revision;
      this.queryController.noteCurrentRevision(result.revision, false);
      switch (result.status) {
        case "committed":
        case "corrected":
          committedRows.push(result.row);
          this.queryController.applyCanonicalRows([result.row], result.revision, false);
          this.overlays.remove(result.clientMutationId);
          break;
        case "rejected":
          this.overlays.remove(result.clientMutationId);
          issues.push(...result.issues.map((issue) => ({
            ...issue,
            rowId: overlay.rowId,
            columnId: issue.columnId ?? overlay.columnId
          })));
          break;
        case "conflict": {
          this.queryController.applyCanonicalRows([result.current], result.revision, false);
          this.overlays.updateStatus(result.clientMutationId, "conflict");
          this.conflicts.set(
            result.clientMutationId,
            createConflict(overlay, result.current, result.revision, this.source)
          );
          conflictResult ??= result;
          break;
        }
      }
    }

    this.batches.delete(operationId);
    this.issues = issues;
    if (invalidated || (committedRows.length > 0 && queryIsProjected(this.getActiveQuery()))) {
      this.queryController.invalidate();
    } else if (committedRows.length > 0 || issues.length > 0 || conflictResult) {
      this.queryController.publishCanonicalState();
    } else {
      this.publish();
    }

    if (conflictResult) {
      return {
        status: "conflict",
        revision: conflictResult.revision,
        current: conflictResult.current
      };
    }
    if (issues.length > 0) {
      return { status: "rejected", reason: "validation", issues };
    }
    return { status: "committed", revision: latestRevision, changed: true };
  }

  getOverlay(rowId: string, columnId: string) {
    return this.overlays.getLatest(rowId, columnId);
  }

  preflight(
    operationId: string,
    prepared: readonly PreparedRemoteMutation[]
  ): Extract<CommandResult, { status: "rejected" }> | null {
    if (this.destroyed || !this.source.mutate) return unsupported("Remote mutation is unavailable");
    if (!operationId.trim() || this.batches.has(operationId)) {
      return validationIssue("REMOTE_MUTATION_PROTOCOL_ERROR", "Mutation operation ID is invalid or active");
    }
    if (!this.queryController.getCurrentRevision()) {
      return validationIssue("NO_QUERY", "A current source revision is required");
    }
    const newCellKeys = new Set(prepared.map((mutation) => cellKey(mutation.rowId, mutation.columnId)));
    const existingCellKeys = new Set(this.overlays.values().map((overlay) => cellKey(overlay.rowId, overlay.columnId)));
    const additionalCells = [...newCellKeys].filter((key) => !existingCellKeys.has(key)).length;
    return this.batches.size >= this.maxPendingOperations
      || this.overlays.uniqueCellCount + additionalCells > this.maxPendingCells
      ? validationIssue("REMOTE_MUTATION_BACKPRESSURE", "Remote mutation queue is full")
      : null;
  }

  getPendingOperations(): readonly TablePendingOperation[] {
    return [...this.batches.values()].map((batch) => ({
      id: batch.operationId,
      feature: batch.feature,
      startedAt: batch.startedAt,
      rowIds: [...new Set(batch.cells.map((cell) => cell.rowId))],
      cells: batch.cells
    }));
  }

  getConflicts(): readonly TableConflict<TRow>[] {
    return [...this.conflicts.values()];
  }

  getIssues(): readonly TableCellIssue[] {
    return this.issues;
  }

  getDiagnostics(): { pendingOperations: number; pendingCells: number; conflicts: number } {
    return {
      pendingOperations: this.batches.size,
      pendingCells: this.overlays.uniqueCellCount,
      conflicts: this.conflicts.size
    };
  }

  destroy(): void {
    if (this.destroyed) return;
    this.destroyed = true;
    for (const batch of this.batches.values()) batch.abortController.abort();
    this.batches.clear();
    this.conflicts.clear();
    this.issues = [];
    this.overlays.clear();
  }

  private removeBatch(batch: MutationBatch, silent: boolean): void {
    for (const mutationId of batch.mutationIds) {
      this.overlays.remove(mutationId);
      this.conflicts.delete(mutationId);
    }
    this.batches.delete(batch.operationId);
    if (!silent) this.publish();
  }

  private publish(): void {
    if (!this.destroyed) this.onChange();
  }
}

function validateResults<TRow>(
  expectedIds: readonly string[],
  results: readonly RemoteMutationResult<TRow>[]
): string | null {
  const expected = new Set(expectedIds);
  const seen = new Set<string>();
  for (const result of results) {
    if (!expected.has(result.clientMutationId)) return "Mutation result contains an unknown clientMutationId";
    if (seen.has(result.clientMutationId)) return "Mutation result contains a duplicate clientMutationId";
    seen.add(result.clientMutationId);
  }
  return seen.size === expected.size ? null : "Mutation result is missing a clientMutationId";
}

function authoritativeRow<TRow>(result: RemoteMutationResult<TRow>): TRow | undefined {
  return result.status === "committed" || result.status === "corrected"
    ? result.row
    : result.status === "conflict"
      ? result.current
      : undefined;
}

function createConflict<TRow>(
  overlay: NonNullable<ReturnType<OptimisticOverlayStore["get"]>>,
  current: TRow,
  revision: string,
  source: RemoteTableSource<TRow>
): TableConflict<TRow> {
  let authoritativeValue: unknown;
  try {
    authoritativeValue = source.readCell?.(current, overlay.columnId).evaluatedValue;
  } catch {
    authoritativeValue = undefined;
  }
  return {
    operationId: overlay.operationId,
    rowId: overlay.rowId,
    columnId: overlay.columnId,
    attemptedValue: overlay.cell.evaluatedValue,
    authoritativeValue,
    current,
    revision
  };
}

function queryIsProjected(query: QueryRequest): boolean {
  return query.sorting.length > 0
    || query.filter !== null
    || query.grouping.length > 0
    || query.aggregates.length > 0
    || query.pagination.kind !== "none";
}

function positiveLimit(value: number | undefined, fallback: number, name: string): number {
  const resolved = value ?? fallback;
  if (!Number.isInteger(resolved) || resolved < 1) throw new RangeError(`${name} must be a positive integer`);
  return resolved;
}

function validationIssue(
  code: string,
  message: string
): Extract<CommandResult, { status: "rejected" }> {
  return { status: "rejected", reason: "validation", issues: [{ code, message }] };
}

function unsupported(message: string): Extract<CommandResult, { status: "rejected" }> {
  return {
    status: "rejected",
    reason: "unsupported",
    issues: [{ code: "TABLE_CAPABILITY_UNSUPPORTED", message }]
  };
}

function isAbortError(error: unknown): boolean {
  return error instanceof DOMException && error.name === "AbortError"
    || Boolean(error && typeof error === "object" && "name" in error && error.name === "AbortError");
}

function cellKey(rowId: string, columnId: string): string {
  return `${rowId.length}:${rowId}${columnId}`;
}
