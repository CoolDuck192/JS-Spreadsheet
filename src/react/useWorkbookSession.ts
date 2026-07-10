import {
  useEffect,
  useId,
  useLayoutEffect,
  useState,
  useSyncExternalStore
} from "react";
import type { CommandEnvelope } from "../core/commands/types";
import {
  createWorkbookSession,
  type WorkbookCommandResult,
  type WorkbookDiagnosticEvent,
  type WorkbookSession,
  type WorkbookSnapshot
} from "../core/workbook/WorkbookSession";
import type { WorkbookCommand } from "../core/workbook/commands";
import type { SpreadsheetServices } from "../core/workbook/services";
import { createBlankWorkbook } from "../lib/workbook";
import type { WorkbookModel } from "../types";
import type {
  SpreadsheetErrorEvent,
  WorkbookChangeEvent,
  WorkbookStorage
} from "../App";
import { getDefaultBrowserWorkbookStorage } from "./browserWorkbookStorage";

export type UseWorkbookSessionCommonOptions = Readonly<{
  storage?: WorkbookStorage | false;
  services?: SpreadsheetServices;
  onDiagnostic?: (event: WorkbookDiagnosticEvent) => void;
  onCommandResult?: (event: Readonly<{
    command: WorkbookCommand;
    result: WorkbookCommandResult;
  }>) => void;
  onWorkbookChangeEvent?: (event: WorkbookChangeEvent) => void;
  onError?: (event: SpreadsheetErrorEvent) => void;
}>;

export type UseWorkbookSessionOptions = UseWorkbookSessionCommonOptions & (
  | {
      workbook: WorkbookModel;
      onWorkbookChange(workbook: WorkbookModel): void;
      defaultWorkbook?: never;
    }
  | {
      defaultWorkbook?: WorkbookModel;
      workbook?: never;
      onWorkbookChange?: never;
    }
);

type CallbackState = {
  controlled: boolean;
  sourceGeneration: number;
  onWorkbookChange?: (workbook: WorkbookModel) => void;
  onDiagnostic?: (event: WorkbookDiagnosticEvent) => void;
  onCommandResult?: (event: Readonly<{
    command: WorkbookCommand;
    result: WorkbookCommandResult;
  }>) => void;
  onWorkbookChangeEvent?: (event: WorkbookChangeEvent) => void;
  onError?: (event: SpreadsheetErrorEvent) => void;
};

type SaveRequest = {
  revision: string;
  workbook: WorkbookModel;
};

type OwnedSessionState = {
  raw: WorkbookSession;
  facade: WorkbookSession;
  storage: WorkbookStorage | false;
  hydrated: boolean;
  hydrationStarted: boolean;
  lastQueuedRevision: string;
  lastQueuedWorkbook: WorkbookModel;
  pendingSave?: SaveRequest;
  saveInFlight: boolean;
  destroyed: boolean;
  everMounted: boolean;
  cleanupToken?: object;
  controlledWorkbook?: WorkbookModel;
  beginHydration(callbacks: CallbackState): void;
  queueSave(snapshot: WorkbookSnapshot): void;
  destroy(): void;
};

const renderSessionCache = new Map<string, OwnedSessionState>();

export function useWorkbookSession(options: UseWorkbookSessionOptions = {}): WorkbookSession {
  const instanceId = useId();
  const controlled = options.workbook !== undefined;
  const sourceWorkbook = controlled ? options.workbook : options.defaultWorkbook;
  const previousSource = useState<{ value: WorkbookModel | undefined; generation: number }>(() => ({
    value: sourceWorkbook,
    generation: 0
  }))[0];
  if (previousSource.value !== sourceWorkbook) {
    previousSource.value = sourceWorkbook;
    previousSource.generation += 1;
  }

  const callbacks = useState<{ current: CallbackState }>(() => ({
    current: createCallbackState(options, controlled, previousSource.generation)
  }))[0];
  callbacks.current = createCallbackState(options, controlled, previousSource.generation);

  const [state] = useState(() => acquireOwnedSession(instanceId, options, callbacks));
  const snapshot = useSyncExternalStore(
    state.facade.subscribe,
    state.facade.getSnapshot,
    state.facade.getSnapshot
  );

  useLayoutEffect(() => {
    if (!controlled || !options.workbook || state.destroyed) {
      return;
    }
    const previous = state.raw.getSnapshot();
    const sourceChanged = state.controlledWorkbook !== options.workbook;
    if (previous.workbook === options.workbook) {
      state.controlledWorkbook = options.workbook;
      return;
    }
    const command: WorkbookCommand = {
      type: "workbook.replace",
      workbook: options.workbook,
      history: "reset"
    };
    const result = state.raw.dispatch(command);
    const next = state.raw.getSnapshot();
    state.controlledWorkbook = options.workbook;
    if (sourceChanged) {
      publishWorkbookChange(callbacks.current, previous, next, result, "external");
    }
  }, [controlled, options.workbook, snapshot.revision, state, callbacks]);

  useEffect(() => {
    state.everMounted = true;
    state.destroyed = false;
    state.cleanupToken = undefined;
    state.beginHydration(callbacks.current);
    return () => {
      const token = {};
      state.cleanupToken = token;
      queueMicrotask(() => {
        if (state.cleanupToken === token) {
          state.destroy();
        }
      });
    };
  }, [state, callbacks]);

  useEffect(() => state.raw.subscribeDiagnostics((event) => {
    invokeSafely(callbacks.current.onDiagnostic, event);
  }), [state, callbacks]);

  useEffect(() => {
    state.queueSave(snapshot);
  }, [snapshot.revision, snapshot.workbook, state, callbacks]);

  return state.facade;
}

function acquireOwnedSession(
  instanceId: string,
  options: UseWorkbookSessionOptions,
  callbacks: { current: CallbackState }
): OwnedSessionState {
  const cached = renderSessionCache.get(instanceId);
  if (cached) {
    return cached;
  }
  const state = createOwnedSessionState(options, callbacks);
  renderSessionCache.set(instanceId, state);
  queueMicrotask(() => {
    if (renderSessionCache.get(instanceId) === state) {
      renderSessionCache.delete(instanceId);
    }
    if (!state.everMounted) {
      state.destroy();
    }
  });
  return state;
}

function createOwnedSessionState(
  options: UseWorkbookSessionOptions,
  callbacks: { current: CallbackState }
): OwnedSessionState {
  const controlled = options.workbook !== undefined;
  const defaultBrowserStorage = !controlled && options.storage === undefined
    ? getDefaultBrowserWorkbookStorage()
    : false;
  const storage = options.storage === false
    ? false
    : options.storage ?? defaultBrowserStorage;
  let initialWorkbook = options.workbook ?? options.defaultWorkbook ?? createBlankWorkbook();
  let prehydrated = false;

  if (defaultBrowserStorage) {
    try {
      const loaded = defaultBrowserStorage.load();
      if (!isPromiseLike(loaded) && loaded) {
        initialWorkbook = loaded;
        prehydrated = true;
      }
    } catch {
      prehydrated = true;
    }
  }

  const createCommandId = options.services?.createCommandId ?? createReactCommandIdFactory();
  const raw = createWorkbookSession({
    workbook: initialWorkbook,
    formulaEngineFactory: options.services?.formulaEngineFactory,
    now: options.services?.now,
    createCommandId
  });

  const state: OwnedSessionState = {
    raw,
    facade: undefined as unknown as WorkbookSession,
    storage,
    hydrated: prehydrated || storage === false,
    hydrationStarted: prehydrated || storage === false,
    lastQueuedRevision: raw.getSnapshot().revision,
    lastQueuedWorkbook: raw.getSnapshot().workbook,
    saveInFlight: false,
    destroyed: false,
    everMounted: false,
    controlledWorkbook: options.workbook,
    beginHydration(currentCallbacks) {
      if (state.hydrationStarted || state.destroyed || !state.storage) {
        return;
      }
      state.hydrationStarted = true;
      const initialRevision = raw.getSnapshot().revision;
      const sourceGeneration = currentCallbacks.sourceGeneration;
      let loaded: ReturnType<WorkbookStorage["load"]>;
      try {
        loaded = state.storage.load();
      } catch {
        finishHydrationFailure(state, callbacks.current);
        return;
      }
      Promise.resolve(loaded).then(
        (workbook) => {
          if (state.destroyed) {
            return;
          }
          const latestCallbacks = callbacks.current;
          if (
            workbook
            && !latestCallbacks.controlled
            && latestCallbacks.sourceGeneration === sourceGeneration
            && raw.getSnapshot().revision === initialRevision
          ) {
            const previous = raw.getSnapshot();
            const result = raw.dispatch({
              type: "workbook.replace",
              workbook,
              history: "reset"
            });
            const next = raw.getSnapshot();
            state.lastQueuedRevision = next.revision;
            state.lastQueuedWorkbook = next.workbook;
            publishWorkbookChange(latestCallbacks, previous, next, result, "storage");
          }
          state.hydrated = true;
          const current = raw.getSnapshot();
          if (
            latestCallbacks.controlled
            || current.revision !== initialRevision
          ) {
            state.queueSave(current);
          }
        },
        () => finishHydrationFailure(state, callbacks.current)
      );
    },
    queueSave(current) {
      if (
        state.destroyed
        || !state.storage
        || !state.hydrated
        || current.revision === state.lastQueuedRevision
        || current.workbook === state.lastQueuedWorkbook
      ) {
        return;
      }
      state.lastQueuedRevision = current.revision;
      state.lastQueuedWorkbook = current.workbook;
      state.pendingSave = { revision: current.revision, workbook: current.workbook };
      pumpSaveQueue(state, callbacks);
    },
    destroy() {
      if (state.destroyed) {
        return;
      }
      state.destroyed = true;
      state.pendingSave = undefined;
      raw.destroy();
    }
  };

  state.facade = {
    getSnapshot: raw.getSnapshot,
    getCellEvaluation: raw.getCellEvaluation,
    subscribe: raw.subscribe,
    subscribeDiagnostics: raw.subscribeDiagnostics,
    dispatch(commandOrEnvelope) {
      const envelope: CommandEnvelope<WorkbookCommand> = "intent" in commandOrEnvelope
        ? commandOrEnvelope
        : { id: createCommandId(), intent: commandOrEnvelope };
      const previous = raw.getSnapshot();
      const result = raw.dispatch(envelope);
      const next = raw.getSnapshot();
      invokeSafely(callbacks.current.onCommandResult, { command: envelope.intent, result });
      publishWorkbookChange(
        callbacks.current,
        previous,
        next,
        result,
        originForCommand(envelope.intent),
        envelope.id,
        true
      );
      return result;
    },
    replaceWorkbook(workbook, replaceOptions) {
      const command: WorkbookCommand = {
        type: "workbook.replace",
        workbook,
        history: replaceOptions?.history === "preserve" ? "commit" : "reset"
      };
      const envelope = { id: createCommandId(), intent: command };
      const previous = raw.getSnapshot();
      const result = raw.dispatch(envelope);
      const next = raw.getSnapshot();
      invokeSafely(callbacks.current.onCommandResult, { command, result });
      publishWorkbookChange(
        callbacks.current,
        previous,
        next,
        result,
        replaceOptions?.origin ?? "external",
        envelope.id,
        replaceOptions?.origin !== "external"
      );
      return result;
    },
    destroy: state.destroy
  };

  return state;
}

function pumpSaveQueue(
  state: OwnedSessionState,
  callbacks: { current: CallbackState }
): void {
  if (state.destroyed || state.saveInFlight || !state.storage || !state.pendingSave) {
    return;
  }
  const request = state.pendingSave;
  state.pendingSave = undefined;
  state.saveInFlight = true;
  state.raw.dispatch({ type: "persistence.status", status: "saving" });

  let saved: ReturnType<WorkbookStorage["save"]>;
  try {
    saved = state.storage.save(request.workbook);
  } catch {
    finishSave(state, callbacks, false);
    return;
  }
  Promise.resolve(saved).then(
    () => finishSave(state, callbacks, true),
    () => finishSave(state, callbacks, false)
  );
}

function finishSave(
  state: OwnedSessionState,
  callbacks: { current: CallbackState },
  succeeded: boolean
): void {
  if (state.destroyed) {
    return;
  }
  state.saveInFlight = false;
  if (succeeded) {
    state.raw.dispatch({ type: "persistence.status", status: "idle" });
  } else {
    state.raw.dispatch({
      type: "persistence.status",
      status: "failed",
      message: "Workbook could not be saved"
    });
    invokeSafely(callbacks.current.onError, {
      code: "storage.save.failed",
      message: "Workbook could not be saved",
      recoverable: true
    });
  }
  pumpSaveQueue(state, callbacks);
}

function finishHydrationFailure(
  state: OwnedSessionState,
  callbacks: CallbackState
): void {
  if (state.destroyed) {
    return;
  }
  state.hydrated = true;
  state.raw.dispatch({
    type: "persistence.status",
    status: "failed",
    message: "Workbook could not be loaded"
  });
  invokeSafely(callbacks.onError, {
    code: "storage.load.failed",
    message: "Workbook could not be loaded",
    recoverable: true
  });
  state.queueSave(state.raw.getSnapshot());
}

function publishWorkbookChange(
  callbacks: CallbackState,
  previous: WorkbookSnapshot,
  next: WorkbookSnapshot,
  result: WorkbookCommandResult,
  origin: WorkbookChangeEvent["origin"],
  commandId?: string,
  acknowledgeControlled = false
): void {
  if (
    result.status !== "committed"
    || previous.revision === next.revision
    || previous.workbook === next.workbook
  ) {
    return;
  }
  if (callbacks.controlled && acknowledgeControlled) {
    invokeSafely(callbacks.onWorkbookChange, next.workbook);
  }
  invokeSafely(callbacks.onWorkbookChangeEvent, {
    workbook: next.workbook,
    revision: next.revision,
    previousRevision: previous.revision,
    origin,
    ...(commandId ? { commandId } : {})
  });
}

function originForCommand(command: WorkbookCommand): WorkbookChangeEvent["origin"] {
  if (command.type === "history.undo") {
    return "undo";
  }
  if (command.type === "history.redo") {
    return "redo";
  }
  return "command";
}

function createCallbackState(
  options: UseWorkbookSessionOptions,
  controlled: boolean,
  sourceGeneration: number
): CallbackState {
  return {
    controlled,
    sourceGeneration,
    onWorkbookChange: options.onWorkbookChange,
    onDiagnostic: options.onDiagnostic,
    onCommandResult: options.onCommandResult,
    onWorkbookChangeEvent: options.onWorkbookChangeEvent,
    onError: options.onError
  };
}

function invokeSafely<T>(callback: ((value: T) => void) | undefined, value: T): void {
  if (!callback) {
    return;
  }
  try {
    callback(value);
  } catch {
    // Host callbacks are isolated from session publication and one another.
  }
}

function isPromiseLike<T>(value: T | Promise<T>): value is Promise<T> {
  return typeof (value as { then?: unknown })?.then === "function";
}

function createReactCommandIdFactory(): () => string {
  const randomUUID = globalThis.crypto?.randomUUID?.bind(globalThis.crypto);
  if (randomUUID) {
    return randomUUID;
  }
  const nonce = `${Date.now().toString(36)}-${Math.random().toString(36).slice(2)}`;
  let sequence = 0;
  return () => `${nonce}-${++sequence}`;
}
