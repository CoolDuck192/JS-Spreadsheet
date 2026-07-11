import type { CellRange, WorkbookHistory, WorkbookModel } from "../../types";
import {
  createWorkbookTableSession,
  type WorkbookTableSession
} from "../../table/workbook/WorkbookTableSession";
import {
  createFormulaEngine,
  type ComputedCellValue,
  type FormulaEngine
} from "../../lib/formulaEngine";
import type { CommandEnvelope, CommandResult, TableIssue } from "../commands/types";
import { createRandomId, type IdGenerator, type IdKind } from "../ids";
import {
  applyWorkbookMutation,
  collectCommandDiagnostics,
  type WorkbookCommand,
  type WorkbookIdReservation,
  type WorkbookMutationContext
} from "./commands";
import { migrateWorkbookModel } from "./migrateWorkbook";
import {
  commitWorkbookHistory,
  createWorkbookHistory,
  redoWorkbookHistory,
  undoWorkbookHistory
} from "./history";

export type { WorkbookCommand, WorkbookIdReservation } from "./commands";

export type WorkbookSnapshot = {
  workbook: WorkbookModel;
  selection: CellRange;
  revision: string;
  canUndo: boolean;
  canRedo: boolean;
  persistence: {
    status: "idle" | "saving" | "failed";
    operation?: "load" | "save";
    message?: string;
  };
};

export type WorkbookCommandResult = CommandResult<WorkbookSnapshot>;

export type WorkbookDiagnosticEvent = {
  category: "command" | "validation" | "persistence" | "performance";
  commandId?: string;
  durationMs?: number;
  metadata: Readonly<Record<string, string | number | boolean | null>>;
};

export type CreateWorkbookSessionOptions = {
  workbook: WorkbookModel;
  selection?: CellRange;
  formulaEngineFactory?: typeof createFormulaEngine;
  now?: () => number;
  createCommandId?: () => string;
  createId?: IdGenerator;
  history?: { maxEntries?: number; maxWeight?: number };
  onDiagnostic?: (event: WorkbookDiagnosticEvent) => void;
};

export interface WorkbookSession {
  getSnapshot(): WorkbookSnapshot;
  getCellEvaluation(sheetId: string, address: string): ComputedCellValue;
  subscribe(listener: () => void): () => void;
  subscribeDiagnostics(listener: (event: WorkbookDiagnosticEvent) => void): () => void;
  dispatch(command: WorkbookCommand | CommandEnvelope<WorkbookCommand>): WorkbookCommandResult;
  replaceWorkbook(
    workbook: WorkbookModel,
    options?: { history?: "reset" | "preserve"; origin?: "external" | "import" }
  ): WorkbookCommandResult;
  table(tableId: string): WorkbookTableSession;
  destroy(): void;
}

type PersistenceState = WorkbookSnapshot["persistence"];

type DraftState = {
  workbook: WorkbookModel;
  selection: CellRange;
  persistence: PersistenceState;
  revisionAffectingChange: boolean;
  resetHistory: boolean;
  historyAffectingChange: boolean;
  historyOverride?: WorkbookHistory;
  selectionHistoryOverride?: SelectionHistory;
};

type SelectionHistory = {
  past: CellRange[];
  present: CellRange;
  future: CellRange[];
};

type AppliedDraft = { status: "applied"; state: DraftState };
type RejectedDraft = {
  status: "rejected";
  reason: "validation" | "permission" | "unsupported";
  issues?: readonly TableIssue[];
};
type DraftResult = AppliedDraft | RejectedDraft;

type IdAllocationTrace = Readonly<{
  kind: IdKind;
  occurrence: number;
  id: string;
  reserved: boolean;
}>;

type TransactionIdPreflight = Readonly<{
  status: "ready";
  createId: IdGenerator;
  allocations: IdAllocationTrace[];
  reservedIds: ReadonlySet<string>;
  finish(): RejectedDraft | undefined;
}>;

type TransactionIdReplay = Readonly<{
  createId: IdGenerator;
  finish(): RejectedDraft | undefined;
}>;

const DEFAULT_SELECTION: CellRange = {
  start: { row: 0, column: 0 },
  end: { row: 0, column: 0 }
};

const PROJECTION_UNAVAILABLE_VALUE: ComputedCellValue = Object.freeze({
  kind: "error",
  code: "#PROJECTION!"
});

export function createWorkbookSession(options: CreateWorkbookSessionOptions): WorkbookSession {
  const formulaEngineFactory = options.formulaEngineFactory ?? createFormulaEngine;
  let formulaEngine: FormulaEngine | null = formulaEngineFactory(options.workbook);
  let projectionUnavailable = false;
  const listeners = new Set<() => void>();
  const diagnosticListeners = new Set<(event: WorkbookDiagnosticEvent) => void>();
  const tableSessions = new Map<string, WorkbookTableSession>();
  const createCommandId = options.createCommandId ?? createDefaultCommandIdFactory();
  const createId = options.createId ?? createRandomId;
  const now = options.now ?? createDefaultNow();

  let history = createWorkbookHistory(options.workbook, options.history);
  let selection = options.selection ?? DEFAULT_SELECTION;
  let selectionHistory = createSelectionHistory(selection);
  let persistence: PersistenceState = { status: "idle" };
  let revision = 0;
  let destroyed = false;
  let snapshot = createSnapshot(history, selection, revision, persistence);

  function dispatch(
    commandOrEnvelope: WorkbookCommand | CommandEnvelope<WorkbookCommand>
  ): WorkbookCommandResult {
    if (destroyed) {
      return destroyedResult();
    }

    const envelope = "intent" in commandOrEnvelope
      ? commandOrEnvelope
      : { id: createCommandId(), intent: commandOrEnvelope };
    const command = envelope.intent;
    const startedAt = now();

    if (
      envelope.expectedRevision !== undefined
      && envelope.expectedRevision !== snapshot.revision
    ) {
      const result: WorkbookCommandResult = {
        status: "conflict",
        revision: snapshot.revision,
        current: snapshot
      };
      emitDiagnostic(createDiagnostic(
        command,
        envelope.id,
        startedAt,
        now(),
        "command",
        "conflict",
        false
      ));
      return result;
    }

    if (projectionUnavailable && commandRequiresProjection(command) && !recoverProjection()) {
      const result = projectionUnavailableResult();
      emitDiagnostic(createDiagnostic(
        command,
        envelope.id,
        startedAt,
        now(),
        "command",
        "rejected",
        false
      ));
      return result;
    }

    let scratchEngine: FormulaEngine | undefined;
    let scratchWorkbook = history.present;
    let transactionDepth = 0;

    const ensureScratchProjection = (workbook: WorkbookModel): FormulaEngine => {
      if (!scratchEngine) {
        scratchEngine = formulaEngineFactory(scratchWorkbook);
      }
      if (scratchWorkbook !== workbook) {
        scratchEngine.update(workbook);
        scratchWorkbook = workbook;
      }
      return scratchEngine;
    };

    const mutationContext: WorkbookMutationContext = {
      createId,
      evaluateCell(workbook, sheetId, address) {
        return ensureScratchProjection(workbook).getComputedValue(sheetId, address);
      }
    };

    const initialState: DraftState = {
      workbook: history.present,
      selection,
      persistence,
      revisionAffectingChange: false,
      resetHistory: false,
      historyAffectingChange: false
    };

    let applied: DraftResult;
    let reductionFailure: { cause: unknown } | undefined;
    try {
      const enterTransaction = () => {
        transactionDepth += 1;
      };
      const leaveTransaction = () => {
        transactionDepth -= 1;
      };
      const projectTransactionWorkbook = (workbook: WorkbookModel) => {
        if (transactionDepth > 0) {
          ensureScratchProjection(workbook);
        }
      };
      if (command.type === "transaction" && transactionNeedsIdPreflight(command)) {
        const prepared = createTransactionIdPreflight(initialState.workbook, command);
        if (prepared.status === "rejected") {
          applied = prepared;
        } else {
          const preflight = reduce(
            command,
            initialState,
            { ...mutationContext, createId: prepared.createId },
            enterTransaction,
            leaveTransaction,
            projectTransactionWorkbook
          );
          if (preflight.status === "rejected") {
            applied = preflight;
          } else {
            const reservationFailure = prepared.finish();
            if (reservationFailure) {
              applied = reservationFailure;
            } else {
              const replay = createTransactionIdReplay(
                initialState.workbook,
                prepared.allocations,
                prepared.reservedIds,
                createId
              );
              const reduced = reduce(
                command,
                initialState,
                { ...mutationContext, createId: replay.createId },
                enterTransaction,
                leaveTransaction,
                projectTransactionWorkbook
              );
              applied = replay.finish() ?? reduced;
            }
          }
        }
      } else {
        applied = reduce(
          command,
          initialState,
          mutationContext,
          enterTransaction,
          leaveTransaction,
          projectTransactionWorkbook
        );
      }
    } catch (cause) {
      reductionFailure = { cause };
      applied = invalidCommandResult();
    } finally {
      try {
        scratchEngine?.destroy();
      } catch {
        // Host cleanup must not replace the command result or its diagnostic.
      }
    }

    if (applied.status === "rejected") {
      const result: WorkbookCommandResult = {
        status: "rejected",
        reason: applied.reason,
        ...(applied.issues ? { issues: applied.issues } : {})
      };
      emitDiagnostic(createDiagnostic(
        command,
        envelope.id,
        startedAt,
        now(),
        applied.reason === "validation" ? "validation" : "command",
        "rejected",
        false,
        reductionFailure
      ));
      return result;
    }

    const nextState = applied.state;
    const workbookChanged = nextState.workbook !== history.present;
    const selectionChanged = !rangesEqual(nextState.selection, selection);
    const persistenceChanged = !persistenceEqual(nextState.persistence, persistence);
    const historyChanged = nextState.historyOverride !== undefined && nextState.historyOverride !== history;
    const historyReset = nextState.resetHistory && (history.past.length > 0 || history.future.length > 0);
    const changed = workbookChanged || selectionChanged || persistenceChanged || historyChanged || historyReset;

    if (!changed) {
      const result: WorkbookCommandResult = {
        status: "committed",
        revision: snapshot.revision,
        changed: false
      };
      emitDiagnostic(createDiagnostic(
        command,
        envelope.id,
        startedAt,
        now(),
        diagnosticCategory(command),
        "committed",
        false
      ));
      return result;
    }

    let nextHistory = history;
    let nextSelectionHistory = selectionHistory;
    if (nextState.historyOverride) {
      const overrideSelections = nextState.selectionHistoryOverride ?? selectionHistory;
      if (nextState.historyOverride.present === nextState.workbook) {
        nextHistory = nextState.historyOverride;
        nextSelectionHistory = rangesEqual(overrideSelections.present, nextState.selection)
          ? overrideSelections
          : { ...overrideSelections, present: cloneRange(nextState.selection) };
      } else {
        nextHistory = commitWorkbookHistory(nextState.historyOverride, nextState.workbook);
        nextSelectionHistory = commitSelectionHistory(
          overrideSelections,
          nextHistory,
          nextState.selection
        );
      }
    } else if (workbookChanged) {
      if (nextState.resetHistory) {
        nextHistory = createWorkbookHistory(nextState.workbook, history.limits);
        nextSelectionHistory = createSelectionHistory(nextState.selection);
      } else if (nextState.historyAffectingChange) {
        nextHistory = commitWorkbookHistory(history, nextState.workbook);
        nextSelectionHistory = commitSelectionHistory(
          selectionHistory,
          nextHistory,
          nextState.selection
        );
      } else {
        nextHistory = { ...history, present: nextState.workbook };
        nextSelectionHistory = {
          ...selectionHistory,
          present: cloneRange(nextState.selection)
        };
      }
    } else if (nextState.resetHistory) {
      nextHistory = createWorkbookHistory(nextState.workbook, history.limits);
      nextSelectionHistory = createSelectionHistory(nextState.selection);
    } else if (selectionChanged) {
      nextSelectionHistory = {
        ...selectionHistory,
        present: cloneRange(nextState.selection)
      };
    }

    if (workbookChanged) {
      try {
        formulaEngine!.update(nextState.workbook);
      } catch {
        recoverProjection();
        const result: WorkbookCommandResult = invalidCommandResult();
        emitDiagnostic(createDiagnostic(
          command,
          envelope.id,
          startedAt,
          now(),
          "command",
          "rejected",
          false
        ));
        return result;
      }
    }

    history = nextHistory;
    selectionHistory = nextSelectionHistory;
    selection = nextState.selection;
    persistence = nextState.persistence;
    if (nextState.revisionAffectingChange && (workbookChanged || selectionChanged)) {
      revision += 1;
    }
    snapshot = createSnapshot(history, selection, revision, persistence);
    publish();

    const result: WorkbookCommandResult = {
      status: "committed",
      revision: snapshot.revision,
      changed: true
    };
    emitDiagnostic(createDiagnostic(
      command,
      envelope.id,
      startedAt,
      now(),
      diagnosticCategory(command),
      "committed",
      true
    ));
    return result;
  }

  function emitDiagnostic(event: WorkbookDiagnosticEvent): void {
    if (destroyed) {
      return;
    }
    invokeSafely(options.onDiagnostic, event);
    for (const listener of [...diagnosticListeners]) {
      invokeSafely(listener, event);
    }
  }

  function publish(): void {
    for (const listener of [...listeners]) {
      try {
        listener();
      } catch {
        // Host listeners are isolated so one integration cannot block another.
      }
    }
  }

  function recoverProjection(): boolean {
    const previousEngine = formulaEngine;
    if (previousEngine) {
      try {
        previousEngine.rebuild(history.present);
        projectionUnavailable = false;
        return true;
      } catch {
        formulaEngine = null;
        projectionUnavailable = true;
        try {
          previousEngine.destroy();
        } catch {
          // A poisoned engine must not prevent replacement with a clean projection.
        }
      }
    }

    try {
      formulaEngine = formulaEngineFactory(history.present);
      projectionUnavailable = false;
      return true;
    } catch {
      formulaEngine = null;
      projectionUnavailable = true;
      return false;
    }
  }

  const session: WorkbookSession = {
    getSnapshot() {
      return snapshot;
    },
    getCellEvaluation(sheetId, address) {
      return projectionUnavailable || !formulaEngine
        ? PROJECTION_UNAVAILABLE_VALUE
        : formulaEngine.getComputedValue(sheetId, address);
    },
    subscribe(listener) {
      if (destroyed) {
        return () => {};
      }
      listeners.add(listener);
      return () => listeners.delete(listener);
    },
    subscribeDiagnostics(listener) {
      if (destroyed) {
        return () => {};
      }
      diagnosticListeners.add(listener);
      return () => diagnosticListeners.delete(listener);
    },
    dispatch,
    replaceWorkbook(workbook, replaceOptions) {
      return dispatch({
        type: "workbook.replace",
        workbook,
        history: replaceOptions?.history === "preserve" ? "commit" : "reset"
      });
    },
    table(tableId) {
      if (destroyed) {
        throw new Error("Workbook session is destroyed");
      }
      const existing = tableSessions.get(tableId);
      if (existing) {
        return existing;
      }
      const child = createWorkbookTableSession(session, tableId, () => {
        if (tableSessions.get(tableId) === child) {
          tableSessions.delete(tableId);
        }
      });
      tableSessions.set(tableId, child);
      return child;
    },
    destroy() {
      if (destroyed) {
        return;
      }
      destroyed = true;
      for (const tableSession of [...tableSessions.values()]) {
        tableSession.destroy();
      }
      tableSessions.clear();
      listeners.clear();
      diagnosticListeners.clear();
      const engine = formulaEngine;
      formulaEngine = null;
      projectionUnavailable = true;
      try {
        engine?.destroy();
      } catch {
        // Cleanup remains idempotent even for a failed host-provided engine.
      }
    }
  };

  return session;

  function reduce(
    command: WorkbookCommand,
    state: DraftState,
    context: WorkbookMutationContext,
    enterTransaction: () => void,
    leaveTransaction: () => void,
    projectTransactionWorkbook: (workbook: WorkbookModel) => void
  ): DraftResult {
    if (command.type === "transaction") {
      enterTransaction();
      let candidate = state;
      try {
        for (const child of command.commands) {
          const result = reduce(
            child,
            candidate,
            context,
            enterTransaction,
            leaveTransaction,
            projectTransactionWorkbook
          );
          if (result.status === "rejected") {
            return result;
          }
          candidate = result.state;
        }
        return { status: "applied", state: candidate };
      } finally {
        leaveTransaction();
      }
    }

    if (command.type === "selection.set") {
      const changed = !rangesEqual(state.selection, command.selection);
      return {
        status: "applied",
        state: changed
          ? {
              ...state,
              selection: command.selection,
              revisionAffectingChange: true
            }
          : state
      };
    }

    if (command.type === "persistence.status") {
      const nextPersistence: PersistenceState = {
        status: command.status,
        ...(command.operation === undefined ? {} : { operation: command.operation }),
        ...(command.message === undefined ? {} : { message: command.message })
      };
      return {
        status: "applied",
        state: persistenceEqual(state.persistence, nextPersistence)
          ? state
          : { ...state, persistence: nextPersistence }
      };
    }

    if (command.type === "history.undo" || command.type === "history.redo") {
      const currentHistory = state.historyOverride ?? history;
      const currentSelectionHistory = state.selectionHistoryOverride ?? selectionHistory;
      const baseHistory = state.workbook === currentHistory.present
        ? currentHistory
        : commitWorkbookHistory(currentHistory, state.workbook);
      const alignedSelectionHistory = rangesEqual(currentSelectionHistory.present, state.selection)
        ? currentSelectionHistory
        : { ...currentSelectionHistory, present: cloneRange(state.selection) };
      const baseSelectionHistory = baseHistory === currentHistory
        ? alignedSelectionHistory
        : commitSelectionHistory(currentSelectionHistory, baseHistory, state.selection);
      const nextHistory = command.type === "history.undo"
        ? undoWorkbookHistory(baseHistory)
        : redoWorkbookHistory(baseHistory);
      if (nextHistory === baseHistory) {
        return { status: "applied", state };
      }
      const nextSelectionHistory = command.type === "history.undo"
        ? undoSelectionHistory(baseSelectionHistory, nextHistory)
        : redoSelectionHistory(baseSelectionHistory, nextHistory);
      projectTransactionWorkbook(nextHistory.present);
      return {
        status: "applied",
        state: {
          ...state,
          workbook: nextHistory.present,
          selection: nextSelectionHistory.present,
          revisionAffectingChange: true,
          historyOverride: nextHistory,
          selectionHistoryOverride: nextSelectionHistory
        }
      };
    }

    if (command.type === "workbook.replace") {
      if (
        migrateWorkbookModel(command.workbook) === null
        || !hasUniqueStructuredTableIds(command.workbook)
      ) {
        return {
          status: "rejected",
          reason: "validation",
          issues: [{
            code: "WORKBOOK_REPLACEMENT_INVALID",
            message: "Replacement workbook must satisfy the persisted workbook contract"
          }]
        };
      }
      if (state.workbook === command.workbook) {
        return {
          status: "applied",
          state: command.history === "reset" && (history.past.length > 0 || history.future.length > 0)
            ? { ...state, resetHistory: true }
            : state
        };
      }
      projectTransactionWorkbook(command.workbook);
      return {
        status: "applied",
        state: {
          ...state,
          workbook: command.workbook,
          revisionAffectingChange: true,
          resetHistory: command.history === "reset",
          historyAffectingChange: state.historyAffectingChange || command.history === "commit"
        }
      };
    }

    const mutation = applyWorkbookMutation(state.workbook, command, context);
    if (mutation.status === "rejected") {
      return mutation;
    }
    if (mutation.workbook === state.workbook) {
      return { status: "applied", state };
    }
    if (
      commandGeneratesIds(command) &&
      (
        migrateWorkbookModel(mutation.workbook) === null ||
        !hasUniqueStructuredTableIds(mutation.workbook)
      )
    ) {
      return {
        status: "rejected",
        reason: "validation",
        issues: [{
          code: "TABLE_GENERATED_ID_INVALID",
          message: "Generated structured-table IDs must be nonblank and unique"
        }]
      };
    }
    return {
      status: "applied",
      state: {
        ...state,
        workbook: mutation.workbook,
        revisionAffectingChange: true,
        historyAffectingChange: state.historyAffectingChange || command.type !== "sheet.activate"
      }
    };
  }
}

function createSnapshot(
  history: WorkbookHistory,
  selection: CellRange,
  revision: number,
  persistence: PersistenceState
): WorkbookSnapshot {
  return {
    workbook: history.present,
    selection,
    revision: String(revision),
    canUndo: history.past.length > 0,
    canRedo: history.future.length > 0,
    persistence
  };
}

function createDiagnostic(
  command: WorkbookCommand,
  commandId: string,
  startedAt: number,
  completedAt: number,
  category: WorkbookDiagnosticEvent["category"],
  outcome: "committed" | "rejected" | "conflict",
  changed: boolean,
  failure?: { cause: unknown }
): WorkbookDiagnosticEvent {
  return {
    category,
    commandId,
    durationMs: Math.max(0, completedAt - startedAt),
    metadata: {
      commandType: command.type,
      outcome,
      changed,
      ...collectCommandDiagnostics(command),
      ...(failure ? serializeDiagnosticError(failure.cause) : {})
    }
  };
}

function serializeDiagnosticError(
  cause: unknown
): Readonly<Record<string, string | number | boolean | null>> {
  let errorType: string = typeof cause;
  let detail = `[${errorType}]`;
  let stack: unknown;
  let isError = false;
  try {
    isError = cause instanceof Error;
  } catch {
    // Host-provided thrown values can be proxies with hostile prototype traps.
  }
  if (isError) {
    const error = cause as Error;
    errorType = "Error";
    try {
      if (typeof error.name === "string") errorType = safeErrorType(error.name);
    } catch {
      // Error metadata is optional and must not replace the original rejection.
    }
    try {
      if (typeof error.message === "string") detail = error.message;
    } catch {
      // Error metadata is optional and must not replace the original rejection.
    }
    try {
      stack = error.stack;
    } catch {
      // Error metadata is optional and must not replace the original rejection.
    }
  } else {
    try {
      detail = String(cause);
    } catch {
      // Keep the type-only fallback when coercion is hostile.
    }
  }
  const stackFrame = sanitizeDiagnosticStackFrame(stack);
  return {
    errorType,
    errorFingerprint: fingerprintDiagnosticDetail(`${errorType}\0${detail}`),
    ...(stackFrame ? { errorStackFrame: stackFrame.slice(0, 512) } : {})
  };
}

function sanitizeDiagnosticStackFrame(stack: unknown): string | undefined {
  if (typeof stack !== "string") return undefined;
  for (const line of stack.split("\n").slice(1)) {
    const location = line.trim().match(/(?:^|[/\\])([A-Za-z0-9_.-]+\.(?:[cm]?[jt]sx?)):(\d+):(\d+)\)?$/);
    if (location) return `${location[1]}:${location[2]}:${location[3]}`;
  }
  return undefined;
}

function safeErrorType(name: string): string {
  return /^[A-Za-z][A-Za-z0-9_.:-]{0,63}$/.test(name) ? name : "Error";
}

function fingerprintDiagnosticDetail(value: string): string {
  let hash = 0x811c9dc5;
  for (let index = 0; index < value.length; index += 1) {
    hash ^= value.charCodeAt(index);
    hash = Math.imul(hash, 0x01000193);
  }
  return (hash >>> 0).toString(16).padStart(8, "0");
}

function diagnosticCategory(command: WorkbookCommand): WorkbookDiagnosticEvent["category"] {
  if (containsFailedPersistence(command)) {
    return "persistence";
  }
  return "command";
}

function containsFailedPersistence(command: WorkbookCommand): boolean {
  if (command.type === "persistence.status") {
    return command.status === "failed";
  }
  return command.type === "transaction" && command.commands.some(containsFailedPersistence);
}

function createSelectionHistory(selection: CellRange): SelectionHistory {
  return {
    past: [],
    present: cloneRange(selection),
    future: []
  };
}

function commitSelectionHistory(
  current: SelectionHistory,
  nextWorkbookHistory: WorkbookHistory,
  selection: CellRange
): SelectionHistory {
  const candidates = [...current.past, current.present];
  return {
    past: candidates.slice(Math.max(0, candidates.length - nextWorkbookHistory.past.length)),
    present: cloneRange(selection),
    future: []
  };
}

function undoSelectionHistory(
  current: SelectionHistory,
  nextWorkbookHistory: WorkbookHistory
): SelectionHistory {
  const previous = current.past.at(-1);
  if (!previous) return current;
  const pastCandidates = current.past.slice(0, -1);
  const futureCandidates = [current.present, ...current.future];
  return {
    past: pastCandidates.slice(Math.max(0, pastCandidates.length - nextWorkbookHistory.past.length)),
    present: cloneRange(previous),
    future: futureCandidates.slice(0, nextWorkbookHistory.future.length)
  };
}

function redoSelectionHistory(
  current: SelectionHistory,
  nextWorkbookHistory: WorkbookHistory
): SelectionHistory {
  const next = current.future[0];
  if (!next) return current;
  const pastCandidates = [...current.past, current.present];
  const futureCandidates = current.future.slice(1);
  return {
    past: pastCandidates.slice(Math.max(0, pastCandidates.length - nextWorkbookHistory.past.length)),
    present: cloneRange(next),
    future: futureCandidates.slice(0, nextWorkbookHistory.future.length)
  };
}

function cloneRange(range: CellRange): CellRange {
  return {
    start: { ...range.start },
    end: { ...range.end }
  };
}

function transactionNeedsIdPreflight(
  command: Extract<WorkbookCommand, { type: "transaction" }>
): boolean {
  return command.idReservations !== undefined
    || command.commands.some(commandOrChildNeedsIdPreflight);
}

function commandOrChildNeedsIdPreflight(command: WorkbookCommand): boolean {
  return command.type === "transaction"
    ? command.idReservations !== undefined
      || command.commands.some(commandOrChildNeedsIdPreflight)
    : commandGeneratesIds(command);
}

function commandGeneratesIds(command: WorkbookCommand): boolean {
  return command.type === "columns.insert"
    || command.type === "sheet.duplicate"
    || command.type === "table.create"
    || command.type === "table.resize"
    || command.type === "table.insertRows";
}

function hasUniqueStructuredTableIds(workbook: WorkbookModel): boolean {
  const reserved = new Set<string>();
  for (const table of workbook.tables) {
    if (!reserve(table.id)) return false;
    for (const column of table.columns) {
      if (!reserve(column.id)) return false;
    }
    for (const rowId of table.rowIds) {
      if (!reserve(rowId)) return false;
    }
  }
  return true;

  function reserve(id: unknown): boolean {
    if (typeof id !== "string" || id.trim().length === 0 || reserved.has(id)) {
      return false;
    }
    reserved.add(id);
    return true;
  }
}

function createTransactionIdPreflight(
  workbook: WorkbookModel,
  command: Extract<WorkbookCommand, { type: "transaction" }>
): TransactionIdPreflight | RejectedDraft {
  if (hasNestedIdReservations(command.commands)) {
    return invalidGeneratedIdsResult();
  }

  const existingIds = collectStructuredTableIds(workbook);
  const reservationSource: unknown = command.idReservations;
  if (reservationSource !== undefined && !Array.isArray(reservationSource)) {
    return invalidGeneratedIdsResult();
  }
  const reservations = (reservationSource ?? []) as readonly unknown[];
  const bySlot = new Map<string, WorkbookIdReservation>();
  const reservationIds = new Set<string>();
  for (const candidate of reservations) {
    if (!isWorkbookIdReservation(candidate)) {
      return invalidGeneratedIdsResult();
    }
    const slot = idReservationSlot(candidate.kind, candidate.occurrence);
    if (
      bySlot.has(slot)
      || reservationIds.has(candidate.id)
      || existingIds.has(candidate.id)
    ) {
      return invalidGeneratedIdsResult();
    }
    bySlot.set(slot, candidate);
    reservationIds.add(candidate.id);
  }

  const commandStrings = collectCommandStringValues(command);
  const unavailable = new Set([...existingIds, ...reservationIds, ...commandStrings]);
  const usedSlots = new Set<string>();
  const allocations: IdAllocationTrace[] = [];
  const occurrences: Record<IdKind, number> = {
    table: 0,
    "table-column": 0,
    "table-row": 0
  };
  let sequence = 0;
  const createId: IdGenerator = (kind) => {
    const occurrence = occurrences[kind]++;
    const slot = idReservationSlot(kind, occurrence);
    const reservation = bySlot.get(slot);
    if (reservation) {
      usedSlots.add(slot);
      allocations.push({
        kind,
        occurrence,
        id: reservation.id,
        reserved: true
      });
      return reservation.id;
    }
    let candidate: string;
    do {
      sequence += 1;
      candidate = `__preflight-${kind}-${sequence}`;
    } while (unavailable.has(candidate));
    unavailable.add(candidate);
    allocations.push({ kind, occurrence, id: candidate, reserved: false });
    return candidate;
  };

  return {
    status: "ready",
    createId,
    allocations,
    reservedIds: reservationIds,
    finish() {
      return usedSlots.size === bySlot.size ? undefined : invalidGeneratedIdsResult();
    }
  };
}

function createTransactionIdReplay(
  workbook: WorkbookModel,
  allocations: readonly IdAllocationTrace[],
  reservedIds: ReadonlySet<string>,
  hostCreateId: IdGenerator
): TransactionIdReplay {
  const seen = collectStructuredTableIds(workbook);
  let allocationIndex = 0;
  let invalid = false;
  const createId: IdGenerator = (kind) => {
    const expected = allocations[allocationIndex++];
    const generated: unknown = hostCreateId(kind);
    const actual = typeof generated === "string" ? generated : "";
    if (!expected || expected.kind !== kind) {
      invalid = true;
    }
    if (
      actual.trim().length === 0
      || seen.has(actual)
      || (expected?.reserved ? actual !== expected.id : reservedIds.has(actual))
    ) {
      invalid = true;
    }
    if (actual.length > 0) {
      seen.add(actual);
    }
    return actual;
  };

  return {
    createId,
    finish() {
      return invalid || allocationIndex !== allocations.length
        ? invalidGeneratedIdsResult()
        : undefined;
    }
  };
}

function collectStructuredTableIds(workbook: WorkbookModel): Set<string> {
  const ids = new Set<string>();
  for (const table of workbook.tables) {
    ids.add(table.id);
    for (const column of table.columns) ids.add(column.id);
    for (const rowId of table.rowIds) ids.add(rowId);
  }
  return ids;
}

function collectCommandStringValues(command: WorkbookCommand): Set<string> {
  const strings = new Set<string>();
  const visited = new WeakSet<object>();
  visit(command);
  return strings;

  function visit(value: unknown): void {
    if (typeof value === "string") {
      strings.add(value);
      return;
    }
    if (!value || typeof value !== "object" || visited.has(value)) {
      return;
    }
    visited.add(value);
    if (Array.isArray(value)) {
      for (const entry of value) visit(entry);
      return;
    }
    for (const entry of Object.values(value as Record<string, unknown>)) {
      visit(entry);
    }
  }
}

function hasNestedIdReservations(commands: readonly WorkbookCommand[]): boolean {
  return commands.some((command) =>
    command.type === "transaction"
    && (
      command.idReservations !== undefined
      || hasNestedIdReservations(command.commands)
    )
  );
}

function isWorkbookIdReservation(value: unknown): value is WorkbookIdReservation {
  if (!value || typeof value !== "object") return false;
  const candidate = value as Partial<Record<keyof WorkbookIdReservation, unknown>>;
  return isIdKind(candidate.kind)
    && typeof candidate.occurrence === "number"
    && Number.isInteger(candidate.occurrence)
    && candidate.occurrence >= 0
    && typeof candidate.id === "string"
    && candidate.id.trim().length > 0;
}

function isIdKind(value: unknown): value is IdKind {
  return value === "table" || value === "table-column" || value === "table-row";
}

function idReservationSlot(kind: IdKind, occurrence: number): string {
  return `${kind}:${occurrence}`;
}

function invalidGeneratedIdsResult(): RejectedDraft {
  return {
    status: "rejected",
    reason: "validation",
    issues: [{
      code: "TABLE_GENERATED_ID_INVALID",
      message: "Generated structured-table IDs must be nonblank and unique"
    }]
  };
}

function rangesEqual(left: CellRange, right: CellRange): boolean {
  return left.start.row === right.start.row
    && left.start.column === right.start.column
    && left.end.row === right.end.row
    && left.end.column === right.end.column;
}

function persistenceEqual(left: PersistenceState, right: PersistenceState): boolean {
  return left.status === right.status
    && left.operation === right.operation
    && left.message === right.message;
}

function invokeSafely<T>(callback: ((value: T) => void) | undefined, value: T): void {
  if (!callback) {
    return;
  }
  try {
    callback(value);
  } catch {
    // Diagnostic sinks are isolated from the command lifecycle.
  }
}

function invalidCommandResult(): RejectedDraft {
  return {
    status: "rejected",
    reason: "unsupported",
    issues: [{ code: "command.invalid", message: "Command could not be applied" }]
  };
}

function destroyedResult(): WorkbookCommandResult {
  return {
    status: "rejected",
    reason: "unsupported",
    issues: [{ code: "session.destroyed", message: "Workbook session is destroyed" }]
  };
}

function projectionUnavailableResult(): WorkbookCommandResult {
  return {
    status: "rejected",
    reason: "unsupported",
    issues: [{ code: "projection.unavailable", message: "Formula projection is unavailable" }]
  };
}

function commandRequiresProjection(command: WorkbookCommand): boolean {
  if (command.type === "selection.set" || command.type === "persistence.status") {
    return false;
  }
  return command.type !== "transaction" || command.commands.some(commandRequiresProjection);
}

function createDefaultNow(): () => number {
  if (typeof globalThis.performance?.now === "function") {
    return () => globalThis.performance.now();
  }
  return () => 0;
}

function createDefaultCommandIdFactory(): () => string {
  const randomUUID = globalThis.crypto?.randomUUID?.bind(globalThis.crypto);
  if (randomUUID) {
    return randomUUID;
  }
  const sessionNonce = `${Date.now().toString(36)}-${Math.random().toString(36).slice(2)}`;
  let sequence = 0;
  return () => `${sessionNonce}-${++sequence}`;
}
