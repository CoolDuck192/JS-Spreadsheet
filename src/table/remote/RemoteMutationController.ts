import type { CommandResult } from "../../core/commands/types";
import type { QueryRequest } from "../core/query";
import type {
  TableCellIssue,
  TableConflict,
  TablePendingOperation
} from "../core/types";
import type { OptimisticCell, OptimisticOverlayStore } from "./OptimisticOverlayStore";
import type {
  RemoteQueryAcceptance,
  RemoteQueryController
} from "./RemoteQueryController";
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

export type RemoteAuthoritativeMutationCell = {
  storedValue: unknown;
  evaluatedValue: unknown;
  formula?: string;
  metadata: OptimisticCell["metadata"];
  rowVersion?: string;
};

export type RemoteMutationControllerOptions<TRow> = {
  source: RemoteTableSource<TRow>;
  queryController: RemoteQueryController<TRow>;
  overlays: OptimisticOverlayStore;
  getActiveQuery(): QueryRequest;
  onChange(): void;
  readAuthoritativeCell?(
    row: TRow,
    rowId: string,
    columnId: string
  ): RemoteAuthoritativeMutationCell | null;
  onAcknowledged?(acknowledgement: RemoteMutationAcknowledgement): void;
  onReconciled?(reconciliation: RemoteMutationReconciliation): void;
  limits?: { maxPendingOperations?: number; maxPendingCells?: number };
};

export type RemoteMutationAcknowledgement = {
  operationId: string;
  revision: string;
  rowVersions: readonly { rowId: string; columnId: string; rowVersion?: string }[];
};

export type RemoteMutationReconciliation = {
  operationId: string;
  outcome: "conflict" | "superseded";
};

export type RemoteMutationExecutionOptions = {
  abortController?: AbortController;
  baseRevision?: string;
};

type MutationBatch = {
  operationId: string;
  mutationIds: readonly string[];
  cells: readonly { rowId: string; columnId: string }[];
  feature: "edit" | "metadata";
  startedAt: number;
  abortController: AbortController;
  status: "pending" | "uncertain";
  cancelled: boolean;
  cancelledAtQueryGeneration: number;
  authoritativeRefreshCompleted: boolean;
  uncertainAtQueryGeneration: number;
  acknowledgeable: boolean;
  resolvedMutationIds: Set<string>;
};

export type RemoteMutationCancellation = {
  markAuthoritativeRefreshCompleted(): void;
};

const DEFAULT_MAX_PENDING_OPERATIONS = 100;
const DEFAULT_MAX_PENDING_CELLS = 1_000;

export class RemoteMutationController<TRow> {
  private readonly source: RemoteTableSource<TRow>;
  private readonly queryController: RemoteQueryController<TRow>;
  private readonly overlays: OptimisticOverlayStore;
  private readonly getActiveQuery: () => QueryRequest;
  private readonly onChange: () => void;
  private readonly authoritativeCellReader:
    | RemoteMutationControllerOptions<TRow>["readAuthoritativeCell"]
    | undefined;
  private readonly onAcknowledged: ((acknowledgement: RemoteMutationAcknowledgement) => void) | undefined;
  private readonly onReconciled: ((reconciliation: RemoteMutationReconciliation) => void) | undefined;
  private readonly unsubscribeAcceptedQuery: () => void;
  private readonly maxPendingOperations: number;
  private readonly maxPendingCells: number;
  private readonly batches = new Map<string, MutationBatch>();
  private readonly conflicts = new Map<string, TableConflict<TRow>>();
  private readonly attempts = new Map<string, PreparedRemoteMutation>();
  private readonly reconciliationTombstones = new Set<string>();
  private issues: TableCellIssue[] = [];
  private destroyed = false;

  constructor(options: RemoteMutationControllerOptions<TRow>) {
    this.source = options.source;
    this.queryController = options.queryController;
    this.overlays = options.overlays;
    this.getActiveQuery = options.getActiveQuery;
    this.onChange = options.onChange;
    this.authoritativeCellReader = options.readAuthoritativeCell;
    this.onAcknowledged = options.onAcknowledged;
    this.onReconciled = options.onReconciled;
    this.unsubscribeAcceptedQuery = this.queryController.subscribeAccepted((acceptance) => {
      this.acceptAcceptedQuery(acceptance);
    });
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
    prepared: readonly PreparedRemoteMutation[],
    options: RemoteMutationExecutionOptions = {}
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
    const revision = options.baseRevision ?? this.queryController.getCurrentRevision()!;

    const mutations: RemoteMutation[] = prepared.map((item, index) => {
      const clientMutationId = `${operationId}:${index}`;
      this.attempts.set(clientMutationId, item);
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
    const abortController = options.abortController ?? new AbortController();
    const batch: MutationBatch = {
      operationId,
      mutationIds: mutations.map((mutation) => mutation.clientMutationId),
      cells: prepared.map(({ rowId, columnId }) => ({ rowId, columnId })),
      feature: prepared.some((mutation) => mutation.kind === "cell-metadata") ? "metadata" : "edit",
      startedAt: Date.now(),
      abortController,
      status: "pending",
      cancelled: false,
      cancelledAtQueryGeneration: 0,
      authoritativeRefreshCompleted: false,
      uncertainAtQueryGeneration: 0,
      acknowledgeable: true,
      resolvedMutationIds: new Set()
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
      if (this.destroyed) {
        this.removeBatch(batch, this.destroyed);
        return unsupported("Remote mutation was aborted");
      }
      if (batch.cancelled) {
        this.removeBatch(batch, true);
        return unsupported("Remote mutation was cancelled for undo");
      }
      if (abortController.signal.aborted || isAbortError(error)) {
        this.removeBatch(batch, false);
        return unsupported("Remote mutation was aborted");
      }
      for (const mutationId of batch.mutationIds) this.overlays.updateStatus(mutationId, "uncertain");
      batch.status = "uncertain";
      batch.uncertainAtQueryGeneration = this.queryController.getDiagnostics().generation;
      this.publish();
      return { status: "pending", operationId };
    }

    if (this.destroyed) {
      this.removeBatch(batch, true);
      return unsupported("Remote mutation was aborted");
    }
    if (batch.cancelled) return this.reconcileCancelledBatch(batch, results);
    if (abortController.signal.aborted) {
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
      if (!overlay) {
        this.reconciliationTombstones.delete(result.clientMutationId);
        continue;
      }
      if ("rowVersion" in result && result.rowVersion !== undefined) {
        const attempt = this.attempts.get(result.clientMutationId);
        if (attempt) this.attempts.set(result.clientMutationId, { ...attempt, rowVersion: result.rowVersion });
      }
      const comparison = this.source.compareRevisions(
        result.revision,
        this.queryController.getCurrentRevision() ?? latestRevision
      );
      if (comparison === "older") {
        this.overlays.remove(result.clientMutationId);
        this.attempts.delete(result.clientMutationId);
        this.conflicts.delete(result.clientMutationId);
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
          this.supersedeOlderUncertainOverlays(overlay);
          committedRows.push(result.row);
          this.queryController.applyCanonicalRows([result.row], result.revision, false);
          this.overlays.remove(result.clientMutationId);
          this.attempts.delete(result.clientMutationId);
          this.conflicts.delete(result.clientMutationId);
          break;
        case "rejected":
          this.overlays.remove(result.clientMutationId);
          this.attempts.delete(result.clientMutationId);
          this.conflicts.delete(result.clientMutationId);
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

    const acknowledged = !invalidated && !conflictResult && issues.length === 0
      && ordered.every((result) => result.status === "committed" || result.status === "corrected");
    if (acknowledged) {
      const latestRowVersions = new Map<string, { revision: string; rowVersion: string }>();
      ordered.forEach((result, index) => {
        if (
          (result.status !== "committed" && result.status !== "corrected")
          || result.rowVersion === undefined
        ) return;
        const rowId = batch.cells[index].rowId;
        const current = latestRowVersions.get(rowId);
        if (
          !current
          || this.source.compareRevisions(result.revision, current.revision) !== "older"
        ) {
          latestRowVersions.set(rowId, { revision: result.revision, rowVersion: result.rowVersion });
        }
      });
      try {
        this.onAcknowledged?.({
          operationId,
          revision: latestRevision,
          rowVersions: batch.cells.map((cell) => {
            const rowVersion = latestRowVersions.get(cell.rowId)?.rowVersion;
            return { ...cell, ...(rowVersion === undefined ? {} : { rowVersion }) };
          })
        });
      } catch {
        // Journal callbacks never alter remote mutation reconciliation.
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

  cancelOperation(operationId: string): RemoteMutationCancellation | null {
    const batch = this.batches.get(operationId);
    if (!batch) return null;
    batch.cancelled = true;
    batch.cancelledAtQueryGeneration = this.queryController.getDiagnostics().generation;
    batch.abortController.abort();
    for (const mutationId of batch.mutationIds) {
      this.reconciliationTombstones.add(mutationId);
      this.overlays.remove(mutationId);
      this.conflicts.delete(mutationId);
      this.attempts.delete(mutationId);
    }
    this.batches.delete(operationId);
    this.publish();
    return {
      markAuthoritativeRefreshCompleted: () => {
        batch.authoritativeRefreshCompleted = true;
      }
    };
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

  hasOverlayForRow(rowId: string): boolean {
    return this.overlays.values().some((overlay) => overlay.rowId === rowId);
  }

  private acceptAcceptedQuery(acceptance: RemoteQueryAcceptance<TRow>): void {
    const rows = acceptance.items.flatMap((item) => item.kind === "data" ? [item.original] : []);
    this.reconcileUncertainBatches(rows, acceptance.revision, acceptance.generation);
  }

  private reconcileUncertainBatches(
    rows: readonly TRow[],
    revision: string,
    generation?: number
  ): void {
    if (this.destroyed) return;
    const rowsById = new Map(rows.map((row) => [this.source.getRowId(row), row]));
    let changed = false;

    for (const batch of [...this.batches.values()]) {
      if (
        batch.status !== "uncertain"
        || (generation !== undefined && generation <= batch.uncertainAtQueryGeneration)
      ) continue;

      const resolutions: Array<{
        mutationId: string;
        overlay: NonNullable<ReturnType<OptimisticOverlayStore["get"]>>;
        attempt: PreparedRemoteMutation;
        row: TRow;
        authoritative: RemoteAuthoritativeMutationCell;
      }> = [];
      let authoritativeForWholeBatch = true;

      for (const mutationId of batch.mutationIds) {
        if (batch.resolvedMutationIds.has(mutationId)) continue;
        const overlay = this.overlays.get(mutationId);
        const attempt = this.attempts.get(mutationId);
        if (!overlay || !attempt) {
          batch.resolvedMutationIds.add(mutationId);
          batch.acknowledgeable = false;
          changed = true;
          continue;
        }

        const latest = this.overlays.getLatest(overlay.rowId, overlay.columnId);
        if (latest && latest.sequence > overlay.sequence) {
          this.overlays.remove(mutationId);
          this.attempts.delete(mutationId);
          this.conflicts.delete(mutationId);
          batch.resolvedMutationIds.add(mutationId);
          batch.acknowledgeable = false;
          changed = true;
          continue;
        }

        const comparison = this.source.compareRevisions(revision, overlay.baseRevision);
        const row = rowsById.get(overlay.rowId);
        const authoritative = row && comparison !== "older" && comparison !== "unknown"
          ? this.readAuthoritativeCell(row, overlay.rowId, overlay.columnId)
          : null;
        if (!row || !authoritative) {
          authoritativeForWholeBatch = false;
          break;
        }
        resolutions.push({ mutationId, overlay, attempt, row, authoritative });
      }

      if (!authoritativeForWholeBatch) continue;
      let hasConflict = false;
      const rowVersions = new Map<string, string | undefined>();
      for (const resolution of resolutions) {
        const { mutationId, overlay, attempt, row, authoritative } = resolution;
        rowVersions.set(mutationId, authoritative.rowVersion);
        batch.resolvedMutationIds.add(mutationId);
        if (attemptMatchesAuthority(attempt, authoritative)) {
          this.overlays.remove(mutationId);
          this.attempts.delete(mutationId);
          this.conflicts.delete(mutationId);
        } else {
          hasConflict = true;
          this.overlays.updateStatus(mutationId, "conflict");
          if (authoritative.rowVersion !== undefined) {
            this.attempts.set(mutationId, { ...attempt, rowVersion: authoritative.rowVersion });
          }
          this.conflicts.set(mutationId, {
            operationId: overlay.operationId,
            rowId: overlay.rowId,
            columnId: overlay.columnId,
            attemptedValue: overlay.cell.evaluatedValue,
            authoritativeValue: authoritative.evaluatedValue,
            current: row,
            revision
          });
        }
        changed = true;
      }

      if (batch.resolvedMutationIds.size !== batch.mutationIds.length) continue;
      this.batches.delete(batch.operationId);
      if (batch.acknowledgeable && !hasConflict) {
        this.notifyAcknowledged({
          operationId: batch.operationId,
          revision,
          rowVersions: batch.cells.map((cell, index) => {
            const rowVersion = rowVersions.get(batch.mutationIds[index]);
            return { ...cell, ...(rowVersion === undefined ? {} : { rowVersion }) };
          })
        });
      } else {
        this.notifyReconciled({
          operationId: batch.operationId,
          outcome: hasConflict ? "conflict" : "superseded"
        });
      }
      changed = true;
    }

    if (changed) this.publish();
  }

  acceptAuthoritativeRows(rows: readonly TRow[], revision: string): void {
    this.reconcileUncertainBatches(rows, revision);
    for (const row of rows) {
      const rowId = this.source.getRowId(row);
      for (const overlay of this.overlays.values().filter((candidate) => candidate.rowId === rowId)) {
        const authoritativeValue = this.readAuthoritativeCell(row, rowId, overlay.columnId)?.evaluatedValue;
        if (!Object.is(authoritativeValue, overlay.cell.evaluatedValue)) {
          this.overlays.updateStatus(overlay.clientMutationId, "conflict");
          this.conflicts.set(overlay.clientMutationId, {
            operationId: overlay.operationId,
            rowId,
            columnId: overlay.columnId,
            attemptedValue: overlay.cell.evaluatedValue,
            authoritativeValue,
            current: row,
            revision
          });
        }
      }
    }
    this.queryController.applyCanonicalRows(rows, revision, false);
    this.queryController.noteCurrentRevision(revision, false);
  }

  acceptAuthoritativeDeletions(rowIds: readonly string[], revision: string): void {
    for (const rowId of rowIds) {
      const current = this.queryController.getCanonicalRow(rowId);
      if (!current) continue;
      for (const overlay of this.overlays.values().filter((candidate) => candidate.rowId === rowId)) {
        this.overlays.updateStatus(overlay.clientMutationId, "conflict");
        this.conflicts.set(overlay.clientMutationId, {
          operationId: overlay.operationId,
          rowId,
          columnId: overlay.columnId,
          attemptedValue: overlay.cell.evaluatedValue,
          authoritativeValue: undefined,
          current,
          revision
        });
      }
    }
    this.queryController.deleteCanonicalRows(rowIds, revision, false);
    this.queryController.noteCurrentRevision(revision, false);
  }

  abandonConflict(operationId: string, rowId: string): boolean {
    const mutationIds = this.conflictMutationIds(operationId, rowId);
    if (mutationIds.length === 0) return false;
    const acknowledgementPending = this.batches.has(operationId);
    for (const mutationId of mutationIds) {
      this.conflicts.delete(mutationId);
      this.overlays.remove(mutationId);
      this.attempts.delete(mutationId);
      if (acknowledgementPending) this.reconciliationTombstones.add(mutationId);
    }
    this.detachResolvedBatch(operationId, mutationIds);
    this.publish();
    return true;
  }

  consumeConflictForRetry(
    operationId: string,
    rowId: string,
    expectedRevision: string
  ): readonly PreparedRemoteMutation[] | null {
    const mutationIds = this.conflictMutationIds(operationId, rowId);
    if (
      mutationIds.length === 0
      || mutationIds.some((mutationId) => this.conflicts.get(mutationId)?.revision !== expectedRevision)
    ) {
      return null;
    }
    const attempts = mutationIds.flatMap((mutationId) => {
      const attempt = this.attempts.get(mutationId);
      return attempt ? [attempt] : [];
    });
    if (attempts.length !== mutationIds.length) return null;
    const acknowledgementPending = this.batches.has(operationId);
    for (const mutationId of mutationIds) {
      this.conflicts.delete(mutationId);
      this.overlays.remove(mutationId);
      this.attempts.delete(mutationId);
      if (acknowledgementPending) this.reconciliationTombstones.add(mutationId);
    }
    this.detachResolvedBatch(operationId, mutationIds);
    this.publish();
    return attempts;
  }

  getConflict(operationId: string, rowId: string): TableConflict<TRow> | undefined {
    return this.conflictMutationIds(operationId, rowId)
      .map((mutationId) => this.conflicts.get(mutationId))
      .find((conflict): conflict is TableConflict<TRow> => conflict !== undefined);
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
    this.unsubscribeAcceptedQuery();
    for (const batch of this.batches.values()) batch.abortController.abort();
    this.batches.clear();
    this.conflicts.clear();
    this.attempts.clear();
    this.reconciliationTombstones.clear();
    this.issues = [];
    this.overlays.clear();
  }

  private readAuthoritativeCell(
    row: TRow,
    rowId: string,
    columnId: string
  ): RemoteAuthoritativeMutationCell | null {
    try {
      if (this.authoritativeCellReader) {
        return this.authoritativeCellReader(row, rowId, columnId);
      }
      if (!this.source.readCell) return null;
      const cell = this.source.readCell(row, columnId);
      return {
        storedValue: cell.storedValue,
        evaluatedValue: cell.evaluatedValue,
        ...(cell.formula === undefined ? {} : { formula: cell.formula }),
        metadata: cell.metadata ?? {},
        ...(cell.rowVersion === undefined ? {} : { rowVersion: cell.rowVersion })
      };
    } catch {
      return null;
    }
  }

  private supersedeOlderUncertainOverlays(
    current: NonNullable<ReturnType<OptimisticOverlayStore["get"]>>
  ): void {
    const affected = new Set<MutationBatch>();
    for (const candidate of this.overlays.values()) {
      if (
        candidate.clientMutationId === current.clientMutationId
        || (candidate.status !== "uncertain" && candidate.status !== "conflict")
        || candidate.rowId !== current.rowId
        || candidate.columnId !== current.columnId
        || candidate.sequence >= current.sequence
      ) continue;
      const batch = this.batches.get(candidate.operationId);
      this.overlays.remove(candidate.clientMutationId);
      this.attempts.delete(candidate.clientMutationId);
      this.conflicts.delete(candidate.clientMutationId);
      if (!batch) continue;
      batch.resolvedMutationIds.add(candidate.clientMutationId);
      batch.acknowledgeable = false;
      affected.add(batch);
    }
    for (const batch of affected) {
      if (batch.resolvedMutationIds.size !== batch.mutationIds.length) continue;
      this.batches.delete(batch.operationId);
      this.notifyReconciled({ operationId: batch.operationId, outcome: "superseded" });
    }
  }

  private notifyAcknowledged(acknowledgement: RemoteMutationAcknowledgement): void {
    try {
      this.onAcknowledged?.(acknowledgement);
    } catch {
      // Journal callbacks never alter remote mutation reconciliation.
    }
  }

  private notifyReconciled(reconciliation: RemoteMutationReconciliation): void {
    try {
      this.onReconciled?.(reconciliation);
    } catch {
      // Journal callbacks never alter remote mutation reconciliation.
    }
  }

  private removeBatch(batch: MutationBatch, silent: boolean): void {
    for (const mutationId of batch.mutationIds) {
      this.overlays.remove(mutationId);
      this.conflicts.delete(mutationId);
      this.attempts.delete(mutationId);
      this.reconciliationTombstones.delete(mutationId);
    }
    this.batches.delete(batch.operationId);
    if (!silent) this.publish();
  }

  private reconcileCancelledBatch(
    batch: MutationBatch,
    results: readonly RemoteMutationResult<TRow>[]
  ): CommandResult<TRow> {
    this.removeBatch(batch, true);
    const query = this.queryController.getSnapshot();
    const queryGeneration = this.queryController.getDiagnostics().generation;
    const authoritativeRefreshCompleted = batch.authoritativeRefreshCompleted
      || (queryGeneration >= batch.cancelledAtQueryGeneration + 2 && query.status !== "loading");
    const protocolIssue = validateResults(batch.mutationIds, results);
    if (protocolIssue) {
      this.issues = [{ code: "REMOTE_MUTATION_PROTOCOL_ERROR", message: protocolIssue }];
      return validationIssue("REMOTE_MUTATION_PROTOCOL_ERROR", protocolIssue);
    }
    const byId = new Map(results.map((result) => [result.clientMutationId, result]));
    for (const mutationId of batch.mutationIds) {
      const result = byId.get(mutationId)!;
      const currentRevision = this.queryController.getCurrentRevision();
      const comparison = currentRevision
        ? this.source.compareRevisions(result.revision, currentRevision)
        : "newer";
      if (comparison === "older" || comparison === "unknown") continue;
      if (authoritativeRefreshCompleted) continue;
      this.queryController.noteCurrentRevision(result.revision, false);
      const row = authoritativeRow(result);
      if (row) this.queryController.applyCanonicalRows([row], result.revision, false);
    }
    if (!authoritativeRefreshCompleted) this.queryController.publishCanonicalState();
    return unsupported("Remote mutation was cancelled for undo");
  }

  private publish(): void {
    if (!this.destroyed) this.onChange();
  }

  private conflictMutationIds(operationId: string, rowId: string): string[] {
    return [...this.conflicts.entries()]
      .filter(([, conflict]) => conflict.operationId === operationId && conflict.rowId === rowId)
      .map(([mutationId]) => mutationId);
  }

  private detachResolvedBatch(operationId: string, mutationIds: readonly string[]): void {
    const batch = this.batches.get(operationId);
    if (!batch || batch.mutationIds.some((mutationId) => !mutationIds.includes(mutationId))) return;
    batch.abortController.abort();
    this.batches.delete(operationId);
  }
}

function attemptMatchesAuthority(
  attempt: PreparedRemoteMutation,
  authoritative: RemoteAuthoritativeMutationCell
): boolean {
  if (attempt.kind === "cell-metadata") {
    return JSON.stringify(authoritative.metadata) === JSON.stringify(attempt.metadata);
  }
  if (attempt.formula !== undefined) {
    return authoritative.formula === attempt.formula;
  }
  return Object.is(authoritative.storedValue, attempt.parsedValue)
    || Object.is(authoritative.evaluatedValue, attempt.optimisticCell.evaluatedValue);
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
