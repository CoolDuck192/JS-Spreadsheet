import type { CellRange, WorkbookHistory, WorkbookModel } from "../../types";
import {
  createFormulaEngine,
  type ComputedCellValue,
  type FormulaEngine
} from "../../lib/formulaEngine";
import type { CommandEnvelope, CommandResult, TableIssue } from "../commands/types";
import {
  applyWorkbookMutation,
  collectCommandDiagnostics,
  type WorkbookCommand,
  type WorkbookMutationContext
} from "./commands";
import {
  commitWorkbookHistory,
  createWorkbookHistory,
  redoWorkbookHistory,
  undoWorkbookHistory
} from "./history";

export type WorkbookSnapshot = {
  workbook: WorkbookModel;
  selection: CellRange;
  revision: string;
  canUndo: boolean;
  canRedo: boolean;
  persistence: { status: "idle" | "saving" | "failed"; message?: string };
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
  destroy(): void;
}

type PersistenceState = WorkbookSnapshot["persistence"];

type DraftState = {
  workbook: WorkbookModel;
  selection: CellRange;
  persistence: PersistenceState;
  revisionAffectingChange: boolean;
  resetHistory: boolean;
  historyOverride?: WorkbookHistory;
};

type AppliedDraft = { status: "applied"; state: DraftState };
type RejectedDraft = {
  status: "rejected";
  reason: "validation" | "permission" | "unsupported";
  issues?: readonly TableIssue[];
};
type DraftResult = AppliedDraft | RejectedDraft;

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
  const createCommandId = options.createCommandId ?? createDefaultCommandIdFactory();
  const now = options.now ?? createDefaultNow();

  let history = createWorkbookHistory(options.workbook, options.history);
  let selection = options.selection ?? DEFAULT_SELECTION;
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
      evaluateCell(workbook, sheetId, address) {
        return ensureScratchProjection(workbook).getComputedValue(sheetId, address);
      }
    };

    const initialState: DraftState = {
      workbook: history.present,
      selection,
      persistence,
      revisionAffectingChange: false,
      resetHistory: false
    };

    let applied: DraftResult;
    try {
      applied = reduce(command, initialState, mutationContext, () => {
        transactionDepth += 1;
      }, () => {
        transactionDepth -= 1;
      }, (workbook) => {
        if (transactionDepth > 0) {
          ensureScratchProjection(workbook);
        }
      });
    } catch {
      applied = invalidCommandResult();
    } finally {
      scratchEngine?.destroy();
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
        false
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
    if (nextState.historyOverride) {
      nextHistory = nextState.historyOverride.present === nextState.workbook
        ? nextState.historyOverride
        : commitWorkbookHistory(nextState.historyOverride, nextState.workbook);
    } else if (workbookChanged) {
      nextHistory = nextState.resetHistory
        ? createWorkbookHistory(nextState.workbook, history.limits)
        : commitWorkbookHistory(history, nextState.workbook);
    } else if (nextState.resetHistory) {
      nextHistory = createWorkbookHistory(nextState.workbook, history.limits);
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

  return {
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
    destroy() {
      if (destroyed) {
        return;
      }
      destroyed = true;
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
      const nextPersistence: PersistenceState = command.message === undefined
        ? { status: command.status }
        : { status: command.status, message: command.message };
      return {
        status: "applied",
        state: persistenceEqual(state.persistence, nextPersistence)
          ? state
          : { ...state, persistence: nextPersistence }
      };
    }

    if (command.type === "history.undo" || command.type === "history.redo") {
      const currentHistory = state.historyOverride ?? history;
      const baseHistory = state.workbook === currentHistory.present
        ? currentHistory
        : commitWorkbookHistory(currentHistory, state.workbook);
      const nextHistory = command.type === "history.undo"
        ? undoWorkbookHistory(baseHistory)
        : redoWorkbookHistory(baseHistory);
      if (nextHistory === baseHistory) {
        return { status: "applied", state };
      }
      projectTransactionWorkbook(nextHistory.present);
      return {
        status: "applied",
        state: {
          ...state,
          workbook: nextHistory.present,
          revisionAffectingChange: true,
          historyOverride: nextHistory
        }
      };
    }

    if (command.type === "workbook.replace") {
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
          resetHistory: command.history === "reset"
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
    projectTransactionWorkbook(mutation.workbook);
    return {
      status: "applied",
      state: {
        ...state,
        workbook: mutation.workbook,
        revisionAffectingChange: true
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
  changed: boolean
): WorkbookDiagnosticEvent {
  return {
    category,
    commandId,
    durationMs: Math.max(0, completedAt - startedAt),
    metadata: {
      commandType: command.type,
      outcome,
      changed,
      ...collectCommandDiagnostics(command)
    }
  };
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

function rangesEqual(left: CellRange, right: CellRange): boolean {
  return left.start.row === right.start.row
    && left.start.column === right.start.column
    && left.end.row === right.end.row
    && left.end.column === right.end.column;
}

function persistenceEqual(left: PersistenceState, right: PersistenceState): boolean {
  return left.status === right.status && left.message === right.message;
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
