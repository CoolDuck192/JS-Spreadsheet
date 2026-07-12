import type { CommandResult } from "../../core/commands/types";
import { parseCellInput } from "../../core/values/parseCellInput";
import {
  resolveTableOperationStates,
  type TableFeature,
  type TableFeatureConfiguration
} from "../core/capabilities";
import { createCommandIdFactory, type CommandIdFactory } from "../core/commandId";
import { normalizeColumns } from "../core/columnHelper";
import type { QueryRequest, QueryRow, TotalCount } from "../core/query";
import { safeInvokeTableExtension } from "../core/safeInvoke";
import type {
  ChangeContext,
  ColumnDef,
  ExportArtifact,
  ExportOptions,
  TableCellIssue,
  TableCellMetadata,
  TableCellSnapshot,
  TableDiagnosticEvent,
  TableIntent,
  TableSession,
  TableStateUpdater,
  TableValidation,
  TableViewSnapshot,
  TableViewState
} from "../core/types";
import {
  OptimisticOverlayStore,
  type OptimisticCell
} from "./OptimisticOverlayStore";
import {
  RemoteMutationController,
  type PreparedRemoteMutation,
  type RemoteMutationAcknowledgement,
  type RemoteMutationReconciliation
} from "./RemoteMutationController";
import {
  RemoteOperationJournal,
  type RemoteOperationJournalChange,
  type RemoteOperationJournalEntry
} from "./RemoteOperationJournal";
import {
  RemoteQueryController,
  type RemoteQuerySnapshot
} from "./RemoteQueryController";
import { RemoteSubscriptionController } from "./RemoteSubscriptionController";
import type { RemoteTableSource } from "./types";

export type RemoteTableSessionOptions<
  TRow,
  TColumn extends ColumnDef<TRow, any> = ColumnDef<TRow, any>
> = {
  source: RemoteTableSource<TRow>;
  columns: readonly TColumn[];
  state?: Partial<TableViewState>;
  defaultState?: Partial<TableViewState>;
  onStateChange?: (updater: TableStateUpdater, context: ChangeContext) => void;
  features?: TableFeatureConfiguration;
  cache?: { maxPages?: number };
  mutationLimits?: { maxPendingOperations?: number; maxPendingCells?: number };
  commandIdFactory?: CommandIdFactory;
  onDiagnostic?: (event: TableDiagnosticEvent) => void;
};

export type RemoteTableDiagnostics = {
  queryGeneration: number;
  cachedPages: number;
  cachedItems: number;
  cachedGapPages: number;
  cachedGapItems: number;
  hasPageGaps: boolean;
  pendingMutations: number;
  conflicts: number;
  journalEntries: number;
  destroyed: boolean;
};

export interface RemoteTableSession<
  TRow,
  TColumn extends ColumnDef<TRow, any> = ColumnDef<TRow, any>
> extends TableSession<TRow, TColumn> {
  updateOptions(options: RemoteTableSessionOptions<TRow, TColumn>): void;
  start(): void;
  stop(): void;
  getDiagnostics(): RemoteTableDiagnostics;
}

export class RemoteSessionError extends Error {
  constructor(readonly code: string, message: string = code) {
    super(message);
    this.name = "RemoteSessionError";
  }
}

const GROUPING_PAGINATION_ISSUE: TableCellIssue = {
  code: "TABLE_GROUPING_PAGINATION_CONFLICT",
  message: "Grouping requires pagination kind none"
};

export function createRemoteTableSession<
  TRow,
  TColumn extends ColumnDef<TRow, any> = ColumnDef<TRow, any>
>(options: RemoteTableSessionOptions<TRow, TColumn>): RemoteTableSession<TRow, TColumn> {
  return new RemoteTableSessionImpl(options);
}

class RemoteTableSessionImpl<
  TRow,
  TColumn extends ColumnDef<TRow, any>
> implements RemoteTableSession<TRow, TColumn> {
  private options: RemoteTableSessionOptions<TRow, TColumn>;
  private columns: readonly TColumn[];
  private columnsById: ReadonlyMap<string, TColumn>;
  private columnsIdentity: readonly TColumn[];
  private state: TableViewState;
  private controlledStateKeys: Set<keyof TableViewState>;
  private commandIdFactory: CommandIdFactory;
  private controller: RemoteQueryController<TRow> | null = null;
  private mutationController: RemoteMutationController<TRow> | null = null;
  private subscriptionController: RemoteSubscriptionController<TRow> | null = null;
  private readonly operationJournal = new RemoteOperationJournal();
  private readonly compensationTargets = new Map<string, string>();
  private readonly compensationConflictTargets = new Map<string, string>();
  private readonly overlays = new OptimisticOverlayStore();
  private controllerUnsubscribe: (() => void) | null = null;
  private controllerSource: RemoteTableSource<TRow> | null = null;
  private readonly exportControllers = new Set<AbortController>();
  private readonly listeners = new Set<() => void>();
  private snapshot: TableViewSnapshot<TRow, TColumn> | null = null;
  private localRevision = 0;
  private started = false;
  private destroyed = false;
  private needsQuery = true;
  private sourceChanged = false;
  private invalidControlledState = false;
  private invalidPublicationPending = false;
  private lastQuery: QueryRequest | null = null;

  constructor(options: RemoteTableSessionOptions<TRow, TColumn>) {
    validateOptions(options);
    this.options = options;
    this.columnsIdentity = options.columns;
    this.columns = normalizeColumns(options.columns) as readonly TColumn[];
    this.columnsById = new Map(this.columns.map((column) => [column.id, column]));
    this.state = mergeState<TRow, TColumn>(defaultState(options.source), options.defaultState, options.state, this.columns);
    this.controlledStateKeys = new Set(Object.keys(options.state ?? {}) as Array<keyof TableViewState>);
    this.commandIdFactory = options.commandIdFactory ?? createCommandIdFactory();
    this.invalidControlledState = hasGroupingPaginationConflict(this.state);
    this.invalidPublicationPending = this.invalidControlledState;
  }

  updateOptions(options: RemoteTableSessionOptions<TRow, TColumn>): void {
    if (this.destroyed) return;
    validateOptions(options);
    let projectionChanged = false;
    if (options.columns !== this.columnsIdentity) {
      this.columnsIdentity = options.columns;
      this.columns = normalizeColumns(options.columns) as readonly TColumn[];
      this.columnsById = new Map(this.columns.map((column) => [column.id, column]));
      projectionChanged = true;
    }

    if (options.source !== this.options.source) {
      this.sourceChanged = true;
      this.needsQuery = true;
      projectionChanged = true;
    }
    if (options.cache?.maxPages !== this.options.cache?.maxPages) {
      this.sourceChanged = true;
      this.needsQuery = true;
      projectionChanged = true;
    }

    const candidate = mergeControlledState<TRow, TColumn>(this.state, options.state, this.columns);
    const invalid = hasGroupingPaginationConflict(candidate);
    if (invalid) {
      if (!this.invalidControlledState) this.invalidPublicationPending = true;
      this.invalidControlledState = true;
      this.needsQuery = false;
      projectionChanged = true;
    } else {
      if (this.invalidControlledState) projectionChanged = true;
      this.invalidControlledState = false;
      this.invalidPublicationPending = false;
      if (!stateEqual(candidate, this.state)) {
        if (!queryStateEqual(candidate, this.state)) this.needsQuery = true;
        this.state = candidate;
        projectionChanged = true;
      }
    }

    if (!stateEqual(options.features ?? {}, this.options.features ?? {})) projectionChanged = true;
    this.options = options;
    this.controlledStateKeys = new Set(Object.keys(options.state ?? {}) as Array<keyof TableViewState>);
    if (projectionChanged) this.snapshot = null;
  }

  start(): void {
    if (this.destroyed) return;
    this.started = true;
    if (this.invalidControlledState) {
      if (this.invalidPublicationPending) {
        this.invalidPublicationPending = false;
        this.localRevision += 1;
        this.snapshot = null;
        this.publish();
      }
      return;
    }
    if (!this.controller || this.sourceChanged || this.controllerSource !== this.options.source) {
      this.teardownController();
      const source = this.options.source;
      const controller = new RemoteQueryController(source, {
        maxCachedPages: this.options.cache?.maxPages,
        onCacheGap: (warning) => this.safeDiagnostic({
          category: "remote",
          commandId: warning.operationId,
          metadata: {
            code: "REMOTE_CACHE_WINDOW_GAP",
            outcome: "warning",
            mode: warning.mode,
            maxPages: warning.maxPages,
            cachedPages: warning.cachedPages,
            omittedPages: warning.omittedPages,
            omittedItems: warning.omittedItems
          }
        })
      });
      this.controller = controller;
      this.controllerSource = source;
      this.mutationController = new RemoteMutationController({
        source,
        queryController: controller,
        overlays: this.overlays,
        getActiveQuery: () => this.lastQuery ?? queryFromState(this.state),
        readAuthoritativeCell: (row, rowId, columnId) => {
          const authoritative = this.readAuthoritativeCell(row, rowId, columnId);
          if ("issue" in authoritative) return null;
          return {
            storedValue: authoritative.cell.storedValue,
            evaluatedValue: authoritative.cell.evaluatedValue,
            ...(authoritative.cell.formula === undefined
              ? {}
              : { formula: authoritative.cell.formula }),
            metadata: authoritative.cell.metadata,
            ...(authoritative.rowVersion === undefined
              ? {}
              : { rowVersion: authoritative.rowVersion })
          };
        },
        onChange: () => {
          if (this.destroyed || this.controller !== controller) return;
          this.localRevision += 1;
          this.snapshot = null;
          this.publish();
        },
        onAcknowledged: (acknowledgement) => this.handleMutationAcknowledged(acknowledgement),
        onReconciled: (reconciliation) => this.handleMutationReconciled(reconciliation),
        limits: this.options.mutationLimits
      });
      if (source.capabilities.subscription) {
        this.subscriptionController = new RemoteSubscriptionController({
          source,
          queryController: controller,
          mutationController: this.mutationController,
          getActiveQuery: () => this.lastQuery ?? queryFromState(this.state),
          createOperationId: this.commandIdFactory
        });
        this.subscriptionController.start();
      }
      this.controllerUnsubscribe = controller.subscribe(() => {
        if (this.destroyed || this.controller !== controller || this.options.source !== source) return;
        this.localRevision += 1;
        this.snapshot = null;
        this.publish();
        this.emitRemoteDiagnostic(controller.getSnapshot());
      });
      this.sourceChanged = false;
      this.needsQuery = true;
    }
    if (this.needsQuery) this.beginQuery();
  }

  stop(): void {
    if (this.destroyed) return;
    this.started = false;
    this.teardownController();
    this.abortExports();
    this.needsQuery = true;
    this.snapshot = null;
  }

  getSnapshot(): TableViewSnapshot<TRow, TColumn> {
    if (this.snapshot) return this.snapshot;
    const query = this.sourceChanged
      ? idleQuerySnapshot<TRow>()
      : this.controller?.getSnapshot() ?? idleQuerySnapshot<TRow>();
    const rows = query.items;
    const operationStates = {
      ...resolveTableOperationStates(this.options.source.capabilities, this.options.features ?? {})
    };
    if (this.options.source.paginationMode !== "none" && operationStates.group.enabled) {
      operationStates.group = {
        ...operationStates.group,
        enabled: false,
        reason: "Grouping requires pagination mode none"
      };
    }
    if (this.invalidControlledState) {
      operationStates.pagination = {
        enabled: false,
        reason: GROUPING_PAGINATION_ISSUE.message
      };
    }
    const baseIssues: TableCellIssue[] = this.invalidControlledState
      ? [GROUPING_PAGINATION_ISSUE]
      : query.error
        ? [{ code: query.error.code, message: query.error.message }]
        : [];
    const issues = [...baseIssues, ...(this.mutationController?.getIssues() ?? [])];
    const snapshot: TableViewSnapshot<TRow, TColumn> = {
      revision: `${query.revision ?? "0"}:${this.localRevision}`,
      rows,
      columns: this.columns,
      rowCount: rows.length,
      totalRowCount: totalFromPageInfo(query.pageInfo),
      completeness: query.completeness,
      state: this.state,
      selection: this.state.selection,
      status: this.invalidControlledState
        ? { phase: "error", message: GROUPING_PAGINATION_ISSUE.message }
        : query.error
          ? { phase: "error", message: query.error.message }
          : { phase: query.status },
      issues,
      capabilities: this.options.source.capabilities,
      operationStates,
      pendingOperations: this.mutationController?.getPendingOperations() ?? [],
      conflicts: this.mutationController?.getConflicts() ?? [],
      canUndo: operationStates.undo.enabled
        && this.options.source.undoMode === "compensating"
        && this.operationJournal.canUndo,
      canRedo: false,
      pageInfo: query.pageInfo,
      pageGaps: query.pageGaps,
      getCell: (rowId, columnId) => this.readCell(rows, rowId, columnId),
      getRowIndex: (rowId) => rows.findIndex((row) => row.id === rowId),
      getColumnIndex: (columnId) => this.state.columnOrder.indexOf(columnId)
    };
    this.snapshot = snapshot;
    return snapshot;
  }

  subscribe(listener: () => void): () => void {
    if (this.destroyed) return () => {};
    this.listeners.add(listener);
    return () => this.listeners.delete(listener);
  }

  async dispatch(intent: TableIntent<TRow>): Promise<CommandResult> {
    const commandId = this.commandIdFactory();
    if (this.destroyed) return unsupported("Remote table session is destroyed");
    const feature = featureForRemoteIntent(intent);
    if (feature) {
      const operation = this.getSnapshot().operationStates[feature];
      if (!operation.enabled) return unsupported(operation.reason ?? "Unsupported operation");
    }

    switch (intent.type) {
      case "edit-cells":
      case "clear-cells":
        return this.editCells(intent, commandId);
      case "update-cell-metadata":
        return this.updateCellMetadata(intent, commandId);
      case "insert-rows":
      case "delete-rows":
        return unsupported("Remote row mutations are not enabled yet");
      case "undo":
        return this.undoRemote(commandId);
      case "redo":
        return unsupported("Remote redo is not supported");
      case "reload-authoritative":
        return this.reloadAuthoritative(intent, commandId);
      case "retry-with-revision":
        return this.retryConflict(intent, commandId);
      case "refresh":
        try {
          await this.refreshWithId(commandId);
          return { status: "committed", revision: this.getSnapshot().revision, changed: true };
        } catch (error) {
          return unsupported(errorMessage(error));
        }
      default:
        return this.updateViewState(intent, commandId);
    }
  }

  async refresh(): Promise<void> {
    await this.refreshWithId(this.commandIdFactory());
  }

  async undo(): Promise<CommandResult> {
    return this.dispatch({ type: "undo" });
  }

  async redo(): Promise<CommandResult> {
    return this.dispatch({ type: "redo" });
  }

  async export(options: ExportOptions): Promise<ExportArtifact> {
    if (this.destroyed) throw new RemoteSessionError("REMOTE_SESSION_DESTROYED");
    const operation = this.getSnapshot().operationStates.export;
    if (!operation.enabled || !this.options.source.export) {
      throw new RemoteSessionError(
        "TABLE_CAPABILITY_UNSUPPORTED",
        operation.reason ?? "Remote export is unsupported"
      );
    }
    const capability = this.options.source.capabilities.export;
    if (options.scope === "completeDataset" && capability && capability.scope !== "completeDataset") {
      throw new RemoteSessionError(
        "TABLE_CAPABILITY_UNSUPPORTED",
        "This source can only export loaded rows"
      );
    }
    const query = this.lastQuery;
    const revision = this.controller?.getSnapshot().revision;
    if (!query || !revision) throw new RemoteSessionError("NO_QUERY", "A completed query is required before export");
    const operationId = this.commandIdFactory();
    const abortController = new AbortController();
    this.exportControllers.add(abortController);
    try {
      const artifact = await this.options.source.export(
        { ...options, query, revision },
        { signal: abortController.signal, operationId }
      );
      if (abortController.signal.aborted || this.destroyed) {
        throw new RemoteSessionError("REMOTE_SESSION_DESTROYED");
      }
      validateExportArtifact(artifact, options.format);
      return { ...artifact, bytes: new Uint8Array(artifact.bytes) };
    } finally {
      this.exportControllers.delete(abortController);
    }
  }

  getDiagnostics(): RemoteTableDiagnostics {
    const query = this.controller?.getDiagnostics();
    return {
      queryGeneration: query?.generation ?? 0,
      cachedPages: query?.cachedPages ?? 0,
      cachedItems: query?.cachedItems ?? 0,
      cachedGapPages: query?.cachedGapPages ?? 0,
      cachedGapItems: query?.cachedGapItems ?? 0,
      hasPageGaps: query?.hasPageGaps ?? false,
      pendingMutations: this.mutationController?.getDiagnostics().pendingOperations ?? 0,
      conflicts: this.mutationController?.getDiagnostics().conflicts ?? 0,
      journalEntries: this.operationJournal.size,
      destroyed: this.destroyed
    };
  }

  destroy(): void {
    if (this.destroyed) return;
    this.stop();
    this.destroyed = true;
    this.listeners.clear();
    this.snapshot = null;
  }

  private beginQuery(): void {
    if (!this.started || !this.controller || this.invalidControlledState) return;
    const query = queryFromState(this.state);
    if (query.pagination.kind !== this.options.source.paginationMode) {
      this.invalidControlledState = true;
      this.snapshot = null;
      return;
    }
    this.needsQuery = false;
    this.lastQuery = query;
    const operationId = this.commandIdFactory();
    void this.controller.load(query, operationId).catch((error) => {
      this.safeDiagnostic({
        category: "remote",
        commandId: operationId,
        metadata: { code: errorCode(error), outcome: "rejected" }
      });
    });
  }

  private async refreshWithId(operationId: string): Promise<void> {
    if (this.destroyed) throw new RemoteSessionError("REMOTE_SESSION_DESTROYED");
    if (!this.controller || !this.started) throw new RemoteSessionError("NO_QUERY");
    await this.controller.refresh(operationId);
  }

  private updateViewState(
    intent: Exclude<TableIntent<TRow>,
      | { type: "edit-cells" }
      | { type: "clear-cells" }
      | { type: "update-cell-metadata" }
      | { type: "insert-rows" }
      | { type: "delete-rows" }
      | { type: "undo" }
      | { type: "redo" }
      | { type: "refresh" }
      | { type: "reload-authoritative" }
      | { type: "retry-with-revision" }
    >,
    commandId: string
  ): CommandResult {
    const candidate = stateForIntent(this.state, intent);
    if (candidate instanceof RemoteSessionError) {
      return { status: "rejected", reason: "validation", issues: [{ code: candidate.code, message: candidate.message }] };
    }
    const invalidColumn = validateViewStateColumns<TRow, TColumn>(candidate, this.columnsById);
    if (invalidColumn) {
      return { status: "rejected", reason: "validation", issues: [invalidColumn] };
    }
    if (stateEqual(candidate, this.state)) {
      return { status: "committed", revision: this.getSnapshot().revision, changed: false };
    }

    const before = this.state;
    const changedKeys = stateChangedKeys(before, candidate);
    const controlled = changedKeys.filter((key) => this.controlledStateKeys.has(key));
    const uncontrolled = changedKeys.filter((key) => !this.controlledStateKeys.has(key));
    if (controlled.length > 0) {
      const updater: TableStateUpdater = (hostState) => {
        let next = hostState;
        for (const key of controlled) next = { ...next, [key]: candidate[key] };
        return next;
      };
      const context: ChangeContext = { commandId, reason: intent.type, revision: this.getSnapshot().revision };
      try { this.options.onStateChange?.(updater, context); } catch { /* Host callbacks are isolated. */ }
    }

    if (uncontrolled.length > 0) {
      let next = before;
      for (const key of uncontrolled) next = { ...next, [key]: candidate[key] };
      this.state = next;
      this.localRevision += 1;
      this.snapshot = null;
      this.publish();
      if (!queryStateEqual(before, next)) {
        this.needsQuery = true;
        this.start();
      }
    }
    return { status: "committed", revision: this.getSnapshot().revision, changed: true };
  }

  private async editCells(
    intent: Extract<TableIntent<TRow>, { type: "edit-cells" | "clear-cells" }>,
    commandId: string
  ): Promise<CommandResult<TRow>> {
    if (!this.controller || !this.mutationController) return unsupported("Remote table is not started");
    const edits = intent.type === "edit-cells"
      ? intent.edits
      : intent.cells.map((cell) => ({ ...cell, rawText: "" }));
    const seen = new Set<string>();
    const prepared: PreparedRemoteMutation[] = [];
    const journalChanges: RemoteOperationJournalChange[] = [];

    for (const edit of edits) {
      const key = `${edit.rowId.length}:${edit.rowId}${edit.columnId}`;
      if (seen.has(key)) {
        return validationResult("TABLE_CELL_DUPLICATE", "A mutation batch may edit each cell once");
      }
      seen.add(key);
      const row = this.controller.getCanonicalRow(edit.rowId);
      if (!row) return validationResult("TABLE_ROW_NOT_FOUND", "Row not found", edit.rowId, edit.columnId);
      const column = this.columnsById.get(edit.columnId);
      if (!column) return validationResult("TABLE_COLUMN_NOT_FOUND", "Column not found", edit.rowId, edit.columnId);
      const context = this.columnContext(row, edit.rowId, edit.columnId);
      if (!this.isEditable(column, context)) {
        return {
          status: "rejected",
          reason: "permission",
          issues: [{ code: "TABLE_CELL_READ_ONLY", message: "Cell is read-only" }]
        };
      }
      const authoritative = this.readAuthoritativeCell(row, edit.rowId, edit.columnId);
      if ("issue" in authoritative) {
        return { status: "rejected", reason: "unsupported", issues: [authoritative.issue] };
      }
      const input = parseCellInput(edit.rawText);
      let parsedValue: unknown = input.stored;
      let evaluatedValue: unknown = input.formula
        ? authoritative.cell.evaluatedValue
        : input.stored;
      let formula = input.formula;
      if (column.parse) {
        const parsed = safeInvokeTableExtension("parse", () => column.parse!(input, context));
        if (!parsed.ok) return { status: "rejected", reason: "unsupported", issues: [parsed.issue] };
        if (!parsed.value.ok) return { status: "rejected", reason: "validation", issues: parsed.value.issues };
        parsedValue = parsed.value.value;
        evaluatedValue = parsed.value.evaluatedValue ?? parsed.value.value;
        formula = parsed.value.formula ?? formula;
      }
      if (formula && !this.getSnapshot().operationStates.formula.enabled) {
        return unsupported(this.getSnapshot().operationStates.formula.reason ?? "Formulas are unsupported");
      }
      if (!compatibleValue(parsedValue, column.dataType, Boolean(formula))) {
        return validationResult("TABLE_VALUE_TYPE", "Cell value has an incompatible type", edit.rowId, edit.columnId);
      }
      if (column.validate) {
        const validation = safeInvokeTableExtension("validate", () => column.validate!({
          ...context,
          raw: edit.rawText,
          parsed: parsedValue,
          evaluated: evaluatedValue
        }));
        if (!validation.ok) return { status: "rejected", reason: "unsupported", issues: [validation.issue] };
        if (validation.value.length > 0) return { status: "rejected", reason: "validation", issues: validation.value };
      }
      const metadataIssues = validateRemoteMetadata(
        parsedValue,
        authoritative.cell.metadata.validation,
        edit.rowId,
        edit.columnId
      );
      if (metadataIssues.length > 0) {
        return { status: "rejected", reason: "validation", issues: metadataIssues };
      }
      const formatted = column.format
        ? safeInvokeTableExtension("format", () => column.format!(evaluatedValue, context))
        : null;
      if (formatted && !formatted.ok) {
        return { status: "rejected", reason: "unsupported", issues: [formatted.issue] };
      }
      prepared.push({
        kind: "cell-value",
        rowId: edit.rowId,
        columnId: edit.columnId,
        rawText: edit.rawText,
        parsedValue,
        ...(formula === undefined ? {} : { formula }),
        ...(authoritative.rowVersion === undefined ? {} : { rowVersion: authoritative.rowVersion }),
        optimisticCell: {
          storedValue: formula ?? parsedValue,
          evaluatedValue,
          displayValue: formatted?.ok ? formatted.value : formatDefault(evaluatedValue),
          ...(formula === undefined ? {} : { formula }),
          metadata: authoritative.cell.metadata
        }
      });
      journalChanges.push({
        kind: "value",
        rowId: edit.rowId,
        columnId: edit.columnId,
        originalValue: authoritative.cell.formula ?? authoritative.cell.storedValue,
        committedValue: formula ?? parsedValue
      });
    }

    if (prepared.length === 0) {
      return { status: "committed", revision: this.getSnapshot().revision, changed: false };
    }
    const preflight = this.mutationController.preflight(commandId, prepared);
    if (preflight) return preflight;
    const journalIssue = this.executeJournaledMutation(commandId, prepared, journalChanges);
    if (journalIssue) return journalIssue;
    return { status: "pending", operationId: commandId };
  }

  private async updateCellMetadata(
    intent: Extract<TableIntent<TRow>, { type: "update-cell-metadata" }>,
    commandId: string
  ): Promise<CommandResult<TRow>> {
    if (!this.controller || !this.mutationController || !this.options.source.readCell) {
      return unsupported("Remote metadata is unavailable");
    }
    const prepared: PreparedRemoteMutation[] = [];
    const journalChanges: RemoteOperationJournalChange[] = [];
    const seen = new Set<string>();
    for (const update of intent.updates) {
      const key = `${update.rowId.length}:${update.rowId}${update.columnId}`;
      if (seen.has(key)) return validationResult("TABLE_CELL_DUPLICATE", "A metadata batch may update each cell once");
      seen.add(key);
      const row = this.controller.getCanonicalRow(update.rowId);
      if (!row) return validationResult("TABLE_ROW_NOT_FOUND", "Row not found", update.rowId, update.columnId);
      if (!this.columnsById.has(update.columnId)) {
        return validationResult("TABLE_COLUMN_NOT_FOUND", "Column not found", update.rowId, update.columnId);
      }
      if (update.patch.formula && !this.getSnapshot().operationStates.formula.enabled) {
        return unsupported(this.getSnapshot().operationStates.formula.reason ?? "Formulas are unsupported");
      }
      const authoritative = this.readAuthoritativeCell(row, update.rowId, update.columnId);
      if ("issue" in authoritative) {
        return { status: "rejected", reason: "unsupported", issues: [authoritative.issue] };
      }
      const metadata = cleanMetadata({ ...authoritative.cell.metadata, ...update.patch });
      prepared.push({
        kind: "cell-metadata",
        rowId: update.rowId,
        columnId: update.columnId,
        metadata,
        ...(authoritative.rowVersion === undefined ? {} : { rowVersion: authoritative.rowVersion }),
        optimisticCell: {
          storedValue: authoritative.cell.storedValue,
          evaluatedValue: authoritative.cell.evaluatedValue,
          displayValue: authoritative.cell.displayValue,
          ...(authoritative.cell.formula === undefined ? {} : { formula: authoritative.cell.formula }),
          metadata
        }
      });
      journalChanges.push({
        kind: "metadata",
        rowId: update.rowId,
        columnId: update.columnId,
        originalMetadata: authoritative.cell.metadata,
        committedMetadata: metadata
      });
    }
    if (prepared.length === 0) {
      return { status: "committed", revision: this.getSnapshot().revision, changed: false };
    }
    const preflight = this.mutationController.preflight(commandId, prepared);
    if (preflight) return preflight;
    const journalIssue = this.executeJournaledMutation(commandId, prepared, journalChanges);
    if (journalIssue) return journalIssue;
    return { status: "pending", operationId: commandId };
  }

  private executeJournaledMutation(
    operationId: string,
    prepared: readonly PreparedRemoteMutation[],
    changes: readonly RemoteOperationJournalChange[]
  ): Extract<CommandResult, { status: "rejected" }> | null {
    const mutations = this.mutationController;
    if (!mutations) return null;
    const journaled = this.options.source.undoMode === "compensating"
      && this.getSnapshot().operationStates.undo.enabled;
    const abortController = journaled ? new AbortController() : undefined;
    if (abortController) {
      try {
        this.operationJournal.begin(operationId, abortController, changes);
      } catch (error) {
        return validationResult("REMOTE_MUTATION_PROTOCOL_ERROR", errorMessage(error));
      }
    }
    void mutations.execute(operationId, prepared, abortController ? { abortController } : {}).then((result) => {
      if (this.operationJournal.isTombstoned(operationId)) {
        this.operationJournal.reconcileTombstone(operationId);
      } else if (
        journaled
        && result.status !== "committed"
        && result.status !== "pending"
      ) {
        this.operationJournal.discardPending(operationId);
        this.publishJournalChange();
      }
      this.snapshot = null;
      this.refreshAfterInvalidation(operationId);
    }).catch(() => {
      if (this.operationJournal.isTombstoned(operationId)) {
        this.operationJournal.reconcileTombstone(operationId);
      } else if (journaled) {
        this.operationJournal.discardPending(operationId);
        this.publishJournalChange();
      }
    });
    return null;
  }

  private handleMutationAcknowledged(acknowledgement: RemoteMutationAcknowledgement): void {
    const compensatedOperationId = this.compensationTargets.get(acknowledgement.operationId);
    if (compensatedOperationId) {
      this.compensationTargets.delete(acknowledgement.operationId);
      this.operationJournal.completeCompensation(compensatedOperationId);
      for (const [operationId, target] of this.compensationConflictTargets) {
        if (target === compensatedOperationId) this.compensationConflictTargets.delete(operationId);
      }
    } else {
      this.operationJournal.acknowledge(
        acknowledgement.operationId,
        acknowledgement.revision,
        acknowledgement.rowVersions
      );
    }
    this.snapshot = null;
  }

  private handleMutationReconciled(reconciliation: RemoteMutationReconciliation): void {
    const compensatedOperationId = this.compensationTargets.get(reconciliation.operationId);
    if (compensatedOperationId) {
      this.compensationTargets.delete(reconciliation.operationId);
      if (reconciliation.outcome === "conflict") {
        this.compensationConflictTargets.set(
          reconciliation.operationId,
          compensatedOperationId
        );
      }
      this.operationJournal.releaseCompensation(compensatedOperationId);
    } else {
      this.operationJournal.discardPending(reconciliation.operationId);
    }
    this.publishJournalChange();
  }

  private undoRemote(commandId: string): CommandResult<TRow> {
    const controller = this.controller;
    const mutations = this.mutationController;
    if (
      !controller
      || !mutations
      || this.options.source.undoMode !== "compensating"
      || !this.getSnapshot().operationStates.undo.enabled
    ) {
      return unsupported("Remote compensation is unavailable");
    }
    const claim = this.operationJournal.claimLatestForUndo();
    if (!claim) return unsupported("No remote compensation is available");
    if (claim.kind === "pending") {
      const cancellation = mutations.cancelOperation(claim.entry.operationId);
      controller.invalidate();
      void controller.refresh(`${commandId}:refresh`).then((accepted) => {
        if (accepted) cancellation?.markAuthoritativeRefreshCompleted();
      }).catch(() => {});
      return { status: "pending", operationId: commandId };
    }

    const prepared = this.prepareCompensatingMutations(claim.entry);
    if (!Array.isArray(prepared)) {
      this.operationJournal.releaseCompensation(claim.entry.operationId);
      return prepared;
    }
    const preflight = mutations.preflight(commandId, prepared);
    if (preflight) {
      this.operationJournal.releaseCompensation(claim.entry.operationId);
      return preflight;
    }
    this.executeCompensatingMutation(
      commandId,
      prepared,
      claim.entry.operationId,
      claim.entry.acknowledgedRevision
    );
    return { status: "pending", operationId: commandId };
  }

  private executeCompensatingMutation(
    commandId: string,
    prepared: readonly PreparedRemoteMutation[],
    compensatedOperationId: string,
    baseRevision?: string
  ): void {
    const mutations = this.mutationController;
    if (!mutations) return;
    this.compensationTargets.set(commandId, compensatedOperationId);
    void mutations.execute(commandId, prepared, baseRevision === undefined ? {} : { baseRevision }).then((result) => {
      const retryableOperationId = this.compensationTargets.get(commandId);
      if (retryableOperationId) {
        if (result.status !== "pending") {
          this.compensationTargets.delete(commandId);
          if (result.status === "conflict") {
            this.compensationConflictTargets.set(commandId, retryableOperationId);
          }
          this.operationJournal.releaseCompensation(retryableOperationId);
          this.publishJournalChange();
        }
      }
      this.snapshot = null;
      this.refreshAfterInvalidation(commandId);
    }).catch(() => {
      const retryableOperationId = this.compensationTargets.get(commandId);
      if (!retryableOperationId) return;
      this.compensationTargets.delete(commandId);
      this.operationJournal.releaseCompensation(retryableOperationId);
      this.publishJournalChange();
    });
  }

  private prepareCompensatingMutations(
    entry: RemoteOperationJournalEntry
  ): PreparedRemoteMutation[] | Extract<CommandResult, { status: "rejected" }> {
    const controller = this.controller;
    if (!controller) return unsupported("Remote table is not started");
    const prepared: PreparedRemoteMutation[] = [];
    for (const change of entry.changes) {
      const row = controller.getCanonicalRow(change.rowId);
      if (!row) return validationResult("TABLE_ROW_NOT_FOUND", "Row not found", change.rowId, change.columnId);
      const column = this.columnsById.get(change.columnId);
      if (!column) return validationResult("TABLE_COLUMN_NOT_FOUND", "Column not found", change.rowId, change.columnId);
      const authoritative = this.readAuthoritativeCell(row, change.rowId, change.columnId);
      if ("issue" in authoritative) {
        return { status: "rejected", reason: "unsupported", issues: [authoritative.issue] };
      }
      const acknowledgedRowVersion = entry.rowVersions.find((candidate) =>
        candidate.rowId === change.rowId && candidate.columnId === change.columnId
      )?.rowVersion;
      const rowVersion = acknowledgedRowVersion ?? authoritative.rowVersion;
      if (change.kind === "metadata") {
        const metadata = cleanMetadata(change.originalMetadata);
        prepared.push({
          kind: "cell-metadata",
          rowId: change.rowId,
          columnId: change.columnId,
          metadata,
          ...(rowVersion === undefined ? {} : { rowVersion }),
          optimisticCell: {
            ...authoritative.cell,
            metadata
          }
        });
        continue;
      }
      const rawText = rawTextForCompensation(change.originalValue);
      const formula = typeof change.originalValue === "string" && change.originalValue.startsWith("=")
        ? change.originalValue
        : undefined;
      const formatted = column.format
        ? safeInvokeTableExtension("format", () => column.format!(change.originalValue, this.columnContext(
            row,
            change.rowId,
            change.columnId
          )))
        : null;
      prepared.push({
        kind: "cell-value",
        rowId: change.rowId,
        columnId: change.columnId,
        rawText,
        parsedValue: change.originalValue,
        ...(formula === undefined ? {} : { formula }),
        ...(rowVersion === undefined ? {} : { rowVersion }),
        optimisticCell: {
          storedValue: change.originalValue,
          evaluatedValue: change.originalValue,
          displayValue: formatted?.ok ? formatted.value : formatDefault(change.originalValue),
          ...(formula === undefined ? {} : { formula }),
          metadata: authoritative.cell.metadata
        }
      });
    }
    return prepared;
  }

  private publishJournalChange(): void {
    if (this.destroyed) return;
    this.localRevision += 1;
    this.snapshot = null;
    this.publish();
  }

  private reloadAuthoritative(
    intent: Extract<TableIntent<TRow>, { type: "reload-authoritative" }>,
    commandId: string
  ): CommandResult<TRow> {
    const controller = this.controller;
    const mutations = this.mutationController;
    const conflict = mutations?.getConflict(intent.operationId, intent.rowId);
    if (!controller || !mutations || !conflict) return conflictNotCurrent();
    if (!mutations.abandonConflict(intent.operationId, intent.rowId)) return conflictNotCurrent();
    this.compensationConflictTargets.delete(intent.operationId);
    controller.invalidate();
    void controller.refresh(commandId).catch(() => {});
    return { status: "pending", operationId: commandId };
  }

  private retryConflict(
    intent: Extract<TableIntent<TRow>, { type: "retry-with-revision" }>,
    commandId: string
  ): CommandResult<TRow> {
    const controller = this.controller;
    const mutations = this.mutationController;
    const conflict = mutations?.getConflict(intent.operationId, intent.rowId);
    const currentRevision = controller?.getCurrentRevision();
    if (
      !controller
      || !mutations
      || !conflict
      || intent.expectedRevision !== conflict.revision
      || currentRevision !== intent.expectedRevision
    ) {
      return conflictNotCurrent();
    }
    const prepared = mutations.consumeConflictForRetry(
      intent.operationId,
      intent.rowId,
      intent.expectedRevision
    );
    if (!prepared) return conflictNotCurrent();
    const compensatedOperationId = this.compensationConflictTargets.get(intent.operationId);
    const preflight = mutations.preflight(commandId, prepared);
    if (preflight) return preflight;
    if (compensatedOperationId) {
      if (!this.operationJournal.reserveCompensation(compensatedOperationId)) {
        return conflictNotCurrent();
      }
      this.compensationConflictTargets.delete(intent.operationId);
      this.executeCompensatingMutation(commandId, prepared, compensatedOperationId);
    } else {
      void mutations.execute(commandId, prepared).then(() => {
        this.snapshot = null;
        this.refreshAfterInvalidation(commandId);
      }).catch(() => {});
    }
    return { status: "pending", operationId: commandId };
  }

  private readAuthoritativeCell(row: TRow, rowId: string, columnId: string):
    | { cell: OptimisticCell; rowVersion?: string }
    | { issue: TableCellIssue } {
    if (!this.options.source.readCell) {
      const evaluated = this.evaluateColumn(row, rowId, columnId, new Set());
      return {
        cell: {
          storedValue: evaluated.storedValue,
          evaluatedValue: evaluated.evaluatedValue,
          displayValue: evaluated.displayValue ?? formatDefault(evaluated.evaluatedValue),
          ...(evaluated.formula === undefined ? {} : { formula: evaluated.formula }),
          metadata: evaluated.metadata,
          issues: evaluated.issues
        }
      };
    }
    const read = safeInvokeTableExtension("accessor", () => this.options.source.readCell!(row, columnId));
    if (!read.ok) return { issue: { ...read.issue, rowId, columnId } };
    return {
      cell: {
        storedValue: read.value.storedValue,
        evaluatedValue: read.value.evaluatedValue,
        displayValue: read.value.displayValue ?? formatDefault(read.value.evaluatedValue),
        ...(read.value.formula === undefined ? {} : { formula: read.value.formula }),
        metadata: read.value.metadata ?? {},
        issues: (read.value.issues ?? []).map((issue) => ({
          ...issue,
          rowId,
          columnId: issue.columnId ?? columnId
        }))
      },
      ...(read.value.rowVersion === undefined ? {} : { rowVersion: read.value.rowVersion })
    };
  }

  private refreshAfterInvalidation(commandId: string): void {
    if (this.controller?.getSnapshot().status !== "error" || !this.started) return;
    void this.controller.refresh(`${commandId}:refresh`).catch(() => {});
  }

  private readCell(
    rows: readonly QueryRow<TRow>[],
    rowId: string,
    columnId: string
  ): TableCellSnapshot {
    const column = this.columnsById.get(columnId);
    if (!column) return missingCell(rowId, columnId, "TABLE_COLUMN_NOT_FOUND", "Column not found");
    const queryRow = rows.find((row) => row.id === rowId);
    if (!queryRow) return missingCell(rowId, columnId, "TABLE_ROW_NOT_FOUND", "Row not found");
    if (queryRow.kind !== "data") return summaryCell(queryRow, columnId, this.state);

    const row = queryRow.original;
    const context = this.columnContext(row, rowId, columnId);
    const overlay = this.mutationController?.getOverlay(rowId, columnId);
    if (overlay) {
      return {
        rowId,
        columnId,
        storedValue: overlay.cell.storedValue,
        evaluatedValue: overlay.cell.evaluatedValue,
        displayValue: overlay.cell.displayValue,
        ...(overlay.cell.formula === undefined ? {} : { formula: overlay.cell.formula }),
        metadata: overlay.cell.metadata,
        editable: overlay.status !== "conflict" && this.isEditable(column, context),
        issues: overlay.cell.issues ?? []
      };
    }
    const sourceCell = this.options.source.readCell
      ? safeInvokeTableExtension("accessor", () => this.options.source.readCell!(row, columnId))
      : null;
    if (sourceCell && !sourceCell.ok) {
      return {
        ...missingCell(rowId, columnId, sourceCell.issue.code, sourceCell.issue.message),
        issues: [{ ...sourceCell.issue, rowId, columnId }]
      };
    }
    const evaluation = sourceCell?.ok
      ? {
          storedValue: sourceCell.value.storedValue,
          evaluatedValue: sourceCell.value.evaluatedValue,
          displayValue: sourceCell.value.displayValue,
          formula: sourceCell.value.formula,
          metadata: sourceCell.value.metadata ?? {},
          issues: (sourceCell.value.issues ?? []).map((issue) => ({ ...issue, rowId, columnId: issue.columnId ?? columnId }))
        }
      : this.evaluateColumn(row, rowId, columnId, new Set());
    const formatted = evaluation.displayValue === undefined && column.format
      ? safeInvokeTableExtension("format", () => column.format!(evaluation.evaluatedValue, context))
      : null;
    const formatIssues = formatted && !formatted.ok ? [{ ...formatted.issue, rowId, columnId }] : [];
    return {
      rowId,
      columnId,
      storedValue: evaluation.storedValue,
      evaluatedValue: evaluation.evaluatedValue,
      displayValue: evaluation.displayValue
        ?? (formatted?.ok ? formatted.value : formatDefault(evaluation.evaluatedValue)),
      ...(evaluation.formula ? { formula: evaluation.formula } : {}),
      metadata: evaluation.metadata,
      editable: this.isEditable(column, context),
      issues: [...evaluation.issues, ...formatIssues]
    };
  }

  private evaluateColumn(
    row: TRow,
    rowId: string,
    columnId: string,
    stack: Set<string>
  ): {
    storedValue: unknown;
    evaluatedValue: unknown;
    displayValue?: string;
    formula?: string;
    metadata: TableCellMetadata;
    issues: TableCellIssue[];
  } {
    if (stack.has(columnId)) {
      return {
        storedValue: null,
        evaluatedValue: { kind: "error", code: "#ERROR!" },
        displayValue: "#ERROR!",
        metadata: {},
        issues: [{ code: "TABLE_CALCULATED_COLUMN_CYCLE", message: "Calculated column cycle", rowId, columnId }]
      };
    }
    const column = this.columnsById.get(columnId);
    if (!column) return { storedValue: null, evaluatedValue: null, metadata: {}, issues: [] };
    stack.add(columnId);
    let value: unknown = null;
    let issue: TableCellIssue | null = null;
    if (column.kind === "computed") {
      const computed = safeInvokeTableExtension("calculate", () => column.calculate({
        row,
        rowId,
        columnId,
        getValue: (dependency) => this.evaluateColumn(row, rowId, dependency, stack).evaluatedValue
      }));
      if (computed.ok) value = computed.value;
      else issue = { ...computed.issue, rowId, columnId };
    } else if ("accessor" in column && typeof column.accessor === "function") {
      const accessed = safeInvokeTableExtension("accessor", () => column.accessor(row));
      if (accessed.ok) value = accessed.value;
      else issue = { ...accessed.issue, rowId, columnId };
    }
    stack.delete(columnId);
    return issue
      ? {
          storedValue: null,
          evaluatedValue: { kind: "error", code: "#ERROR!" },
          displayValue: "#ERROR!",
          metadata: {},
          issues: [issue]
        }
      : { storedValue: value, evaluatedValue: value, metadata: {}, issues: [] };
  }

  private columnContext(row: TRow, rowId: string, columnId: string) {
    return {
      row,
      rowId,
      columnId,
      getValue: (dependency: string) => this.evaluateColumn(row, rowId, dependency, new Set()).evaluatedValue
    };
  }

  private isEditable(column: TColumn, context: ReturnType<RemoteTableSessionImpl<TRow, TColumn>["columnContext"]>): boolean {
    if (!this.options.source.capabilities.edit) return false;
    if (!(column.parse || ("update" in column && typeof column.update === "function"))) return false;
    const editableOption = column.editable;
    const permittedOption = column.permitted;
    const editable = typeof editableOption === "function"
      ? safeBoolean(() => editableOption(context))
      : editableOption !== false;
    const permitted = permittedOption ? safeBoolean(() => permittedOption(context)) : true;
    return editable && permitted;
  }

  private teardownController(): void {
    this.subscriptionController?.destroy();
    this.subscriptionController = null;
    this.mutationController?.destroy();
    this.mutationController = null;
    this.operationJournal.clear();
    this.compensationTargets.clear();
    this.compensationConflictTargets.clear();
    this.overlays.clear();
    this.controllerUnsubscribe?.();
    this.controllerUnsubscribe = null;
    this.controller?.destroy();
    this.controller = null;
    this.controllerSource = null;
  }

  private abortExports(): void {
    for (const controller of this.exportControllers) controller.abort();
    this.exportControllers.clear();
  }

  private publish(): void {
    if (this.destroyed) return;
    for (const listener of [...this.listeners]) {
      try { listener(); } catch { /* Host listeners are isolated. */ }
    }
  }

  private emitRemoteDiagnostic(query: RemoteQuerySnapshot<TRow>): void {
    this.safeDiagnostic({
      category: "remote",
      metadata: {
        outcome: query.status,
        generation: this.controller?.getDiagnostics().generation ?? 0,
        count: query.items.length,
        ...(query.error ? { code: query.error.code } : {})
      }
    });
  }

  private safeDiagnostic(event: TableDiagnosticEvent): void {
    try { this.options.onDiagnostic?.(event); } catch { /* Diagnostics never alter state. */ }
  }
}

function defaultState<TRow>(source: RemoteTableSource<TRow>): TableViewState {
  return {
    sorting: [], filter: null, grouping: [], aggregates: [],
    pagination: defaultPagination(source.paginationMode),
    selection: null, selectedRowIds: [], expandedRowIds: [], columnOrder: [],
    columnVisibility: {}, columnWidths: {}, columnPinning: { left: [], right: [] }
  };
}

function defaultPagination(mode: RemoteTableSource<unknown>["paginationMode"]): TableViewState["pagination"] {
  switch (mode) {
    case "none": return { kind: "none" };
    case "offset": return { kind: "offset", offset: 0, limit: 50 };
    case "cursor": return { kind: "cursor", limit: 50 };
    case "infinite": return { kind: "infinite", limit: 50 };
  }
}

function mergeState<TRow, TColumn extends ColumnDef<TRow, any>>(
  base: TableViewState,
  defaults: Partial<TableViewState> | undefined,
  controlled: Partial<TableViewState> | undefined,
  columns: readonly TColumn[]
): TableViewState {
  return normalizeLayout<TRow, TColumn>({ ...base, ...defaults, ...controlled }, columns);
}

function mergeControlledState<TRow, TColumn extends ColumnDef<TRow, any>>(
  current: TableViewState,
  controlled: Partial<TableViewState> | undefined,
  columns: readonly TColumn[]
): TableViewState {
  return normalizeLayout<TRow, TColumn>({ ...current, ...controlled }, columns);
}

function normalizeLayout<TRow, TColumn extends ColumnDef<TRow, any>>(
  state: TableViewState,
  columns: readonly TColumn[]
): TableViewState {
  const ids = columns.map((column) => column.id);
  const order = state.columnOrder.length === ids.length && new Set(state.columnOrder).size === ids.length
    ? state.columnOrder
    : ids;
  return {
    ...state,
    columnOrder: [...order],
    columnVisibility: Object.fromEntries(columns.map((column) => [
      column.id,
      state.columnVisibility[column.id] ?? column.visible !== false
    ])),
    columnWidths: Object.fromEntries(columns.map((column) => [
      column.id,
      state.columnWidths[column.id] ?? column.width ?? 120
    ])),
    columnPinning: {
      left: state.columnPinning.left.filter((id) => ids.includes(id)),
      right: state.columnPinning.right.filter((id) => ids.includes(id))
    }
  };
}

function stateForIntent<TRow>(
  state: TableViewState,
  intent: Exclude<TableIntent<TRow>,
    | { type: "edit-cells" }
    | { type: "clear-cells" }
    | { type: "update-cell-metadata" }
    | { type: "insert-rows" }
    | { type: "delete-rows" }
    | { type: "undo" }
    | { type: "redo" }
    | { type: "refresh" }
    | { type: "reload-authoritative" }
    | { type: "retry-with-revision" }
  >
): TableViewState | RemoteSessionError {
  switch (intent.type) {
    case "set-selection": return { ...state, selection: intent.selection };
    case "set-row-selection": return { ...state, selectedRowIds: [...intent.rowIds] };
    case "set-row-expanded": return {
      ...state,
      expandedRowIds: intent.expanded
        ? [...new Set([...state.expandedRowIds, intent.rowId])]
        : state.expandedRowIds.filter((id) => id !== intent.rowId)
    };
    case "set-sorting": return { ...state, sorting: [...intent.sorting] };
    case "set-filter": return { ...state, filter: intent.filter };
    case "set-grouping": return {
      ...state,
      grouping: [...intent.grouping],
      ...(intent.grouping.length > 0 ? { pagination: { kind: "none" } as const } : {})
    };
    case "set-aggregates": return { ...state, aggregates: [...intent.aggregates] };
    case "set-pagination":
      return state.grouping.length > 0 && intent.pagination.kind !== "none"
        ? new RemoteSessionError(GROUPING_PAGINATION_ISSUE.code, GROUPING_PAGINATION_ISSUE.message)
        : { ...state, pagination: intent.pagination };
    case "set-column-order": return { ...state, columnOrder: [...intent.columnIds] };
    case "resize-column": return {
      ...state,
      columnWidths: { ...state.columnWidths, [intent.columnId]: intent.width }
    };
    case "set-column-visibility": return {
      ...state,
      columnVisibility: { ...state.columnVisibility, [intent.columnId]: intent.visible }
    };
    case "set-column-pinning": return pinColumn(state, intent.columnId, intent.pin);
  }
}

function queryFromState(state: TableViewState): QueryRequest {
  return {
    sorting: state.sorting,
    filter: state.filter,
    grouping: state.grouping,
    aggregates: state.aggregates,
    pagination: state.pagination,
    ...(state.expandedRowIds.length === 0 ? {} : { tree: { expandedRowIds: state.expandedRowIds } })
  };
}

function featureForRemoteIntent<TRow>(intent: TableIntent<TRow>): TableFeature | null {
  switch (intent.type) {
    case "set-sorting": return "sort";
    case "set-filter": return "filter";
    case "set-grouping": return "group";
    case "set-aggregates": return "aggregate";
    case "set-pagination": return "pagination";
    case "edit-cells": return intent.edits.length > 1 ? "bulkEdit" : "edit";
    case "clear-cells": return intent.cells.length > 1 ? "bulkEdit" : "edit";
    case "update-cell-metadata": return "metadata";
    case "undo":
    case "redo": return "undo";
    default: return null;
  }
}

function summaryCell<TRow>(row: Exclude<QueryRow<TRow>, { kind: "data" }>, columnId: string, state: TableViewState): TableCellSnapshot {
  const aggregateId = state.aggregates.find((aggregate) => aggregate.columnId === columnId)?.id;
  let value: unknown = null;
  if (row.kind === "group" && row.columnId === columnId) value = scalarValue(row.key);
  else value = aggregateId === undefined
    ? row.aggregates[columnId]
    : row.aggregates[aggregateId] ?? row.aggregates[columnId];
  return {
    rowId: row.id,
    columnId,
    storedValue: value ?? null,
    evaluatedValue: value ?? null,
    displayValue: formatDefault(value),
    metadata: {},
    editable: false,
    issues: []
  };
}

function scalarValue(value: { type: string; value?: unknown }): unknown {
  return value.type === "null" ? null : value.value;
}

function missingCell(rowId: string, columnId: string, code: string, message: string): TableCellSnapshot {
  return {
    rowId, columnId, storedValue: null, evaluatedValue: null, displayValue: "",
    metadata: {}, editable: false, issues: [{ code, message, rowId, columnId }]
  };
}

function totalFromPageInfo(pageInfo: RemoteQuerySnapshot<unknown>["pageInfo"]): TotalCount {
  return pageInfo.total;
}

function formatDefault(value: unknown): string {
  if (value === null || value === undefined) return "";
  if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
  if (typeof value === "object" && "code" in value && typeof value.code === "string") return value.code;
  return String(value);
}

function rawTextForCompensation(value: unknown): string {
  if (value === null || value === undefined) return "";
  if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
  return String(value);
}

function validateViewStateColumns<TRow, TColumn extends ColumnDef<TRow, any>>(
  state: TableViewState,
  columns: ReadonlyMap<string, TColumn>
): TableCellIssue | null {
  if (state.columnOrder.length !== columns.size || new Set(state.columnOrder).size !== columns.size) {
    return { code: "TABLE_COLUMN_ORDER_INVALID", message: "Column order must contain every column once" };
  }
  const missing = state.columnOrder.find((id) => !columns.has(id));
  return missing ? { code: "TABLE_COLUMN_NOT_FOUND", message: `Column not found: ${missing}`, columnId: missing } : null;
}

function pinColumn(state: TableViewState, columnId: string, pin: "left" | "right" | false): TableViewState {
  const left = state.columnPinning.left.filter((id) => id !== columnId);
  const right = state.columnPinning.right.filter((id) => id !== columnId);
  if (pin === "left") left.push(columnId);
  if (pin === "right") right.push(columnId);
  return { ...state, columnPinning: { left, right } };
}

function stateChangedKeys(before: TableViewState, after: TableViewState): Array<keyof TableViewState> {
  return (Object.keys(before) as Array<keyof TableViewState>).filter((key) => !stateEqual(before[key], after[key]));
}

function queryStateEqual(left: TableViewState, right: TableViewState): boolean {
  return stateEqual(queryFromState(left), queryFromState(right));
}

function stateEqual(left: unknown, right: unknown): boolean {
  return JSON.stringify(left) === JSON.stringify(right);
}

function hasGroupingPaginationConflict(state: TableViewState): boolean {
  return state.grouping.length > 0 && state.pagination.kind !== "none";
}

function validateOptions<TRow, TColumn extends ColumnDef<TRow, any>>(
  options: RemoteTableSessionOptions<TRow, TColumn>
): void {
  if (options.source.kind !== "remote") throw new Error("RemoteTableSession requires a remote source");
  if (options.mutationLimits?.maxPendingOperations !== undefined && options.mutationLimits.maxPendingOperations < 1) {
    throw new RangeError("maxPendingOperations must be positive");
  }
  if (options.mutationLimits?.maxPendingCells !== undefined && options.mutationLimits.maxPendingCells < 1) {
    throw new RangeError("maxPendingCells must be positive");
  }
}

function validateExportArtifact(artifact: ExportArtifact, format: ExportOptions["format"]): void {
  if (
    !(artifact.bytes instanceof Uint8Array)
    || !artifact.mediaType.trim()
    || !artifact.fileName.trim().toLowerCase().endsWith(`.${format}`)
  ) {
    throw new RemoteSessionError("REMOTE_EXPORT_PROTOCOL_ERROR", "Remote source returned a malformed export artifact");
  }
}

function unsupported(message: string): Extract<CommandResult, { status: "rejected" }> {
  return {
    status: "rejected",
    reason: "unsupported",
    issues: [{ code: "TABLE_CAPABILITY_UNSUPPORTED", message }]
  };
}

function validationResult(
  code: string,
  message: string,
  rowId?: string,
  columnId?: string
): Extract<CommandResult, { status: "rejected" }> {
  const issue: TableCellIssue = {
    code,
    message,
    ...(rowId === undefined ? {} : { rowId }),
    ...(columnId === undefined ? {} : { columnId })
  };
  return { status: "rejected", reason: "validation", issues: [issue] };
}

function conflictNotCurrent(): Extract<CommandResult, { status: "rejected" }> {
  return validationResult(
    "REMOTE_CONFLICT_NOT_CURRENT",
    "The remote conflict is no longer current"
  );
}

function compatibleValue(
  value: unknown,
  dataType: ColumnDef<unknown, any>["dataType"],
  isFormula: boolean
): boolean {
  if (isFormula || value === null || value === undefined || dataType === undefined || dataType === "custom") return true;
  switch (dataType) {
    case "text": return typeof value === "string";
    case "number": return typeof value === "number" && Number.isFinite(value);
    case "boolean": return typeof value === "boolean";
    case "date":
    case "datetime":
      return typeof value === "number" && Number.isFinite(value) || typeof value === "string";
  }
}

function validateRemoteMetadata(
  value: unknown,
  validation: TableValidation | undefined,
  rowId: string,
  columnId: string
): TableCellIssue[] {
  if (!validation || (value === null && validation.allowBlank !== false)) return [];
  let valid = true;
  switch (validation.kind) {
    case "list":
      valid = typeof value === "string" && validation.values.includes(value);
      break;
    case "number":
      valid = typeof value === "number"
        && Number.isFinite(value)
        && (validation.min === undefined || value >= validation.min)
        && (validation.max === undefined || value <= validation.max);
      break;
    case "textLength":
      valid = typeof value === "string"
        && (validation.min === undefined || value.length >= validation.min)
        && (validation.max === undefined || value.length <= validation.max);
      break;
  }
  return valid ? [] : [{
    code: "TABLE_VALIDATION_FAILED",
    message: "Cell value does not satisfy its validation rule",
    rowId,
    columnId
  }];
}

function cleanMetadata(metadata: TableCellMetadata): TableCellMetadata {
  return Object.fromEntries(
    Object.entries(metadata).filter(([, value]) => value !== undefined)
  ) as TableCellMetadata;
}

function safeBoolean(callback: () => boolean): boolean {
  try { return callback(); } catch { return false; }
}

function errorMessage(error: unknown): string {
  return error instanceof Error ? error.message : "Remote operation failed";
}

function errorCode(error: unknown): string {
  return error && typeof error === "object" && "code" in error && typeof error.code === "string"
    ? error.code
    : "REMOTE_OPERATION_FAILED";
}

function idleQuerySnapshot<TRow>(): RemoteQuerySnapshot<TRow> {
  return {
    status: "idle", items: [], revision: null, completeness: "loadedRows",
    pageInfo: { kind: "none", total: { kind: "known", value: 0 } },
    pageGaps: []
  };
}
