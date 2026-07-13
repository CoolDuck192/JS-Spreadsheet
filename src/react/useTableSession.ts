import { useEffect, useLayoutEffect, useRef } from "react";
import {
  createLocalRecordTableSession,
  reviveLocalRecordTableSession,
  type LocalRecordTableSessionOptions,
  type RecordTableSession
} from "../table/local/RecordTableSession";
import {
  createRemoteTableSession,
  type RemoteTableSession,
  type RemoteTableSessionOptions
} from "../table/remote/RemoteTableSession";
import type { ChangeContext, TableStateUpdater } from "../table/core/types";
import type { ColumnDef } from "./tableTypes";

type OwnedTableSession<TRow> =
  | RecordTableSession<TRow, ColumnDef<TRow>>
  | RemoteTableSession<TRow, ColumnDef<TRow>>;

type TableSessionOptions<TRow> =
  | LocalRecordTableSessionOptions<TRow, ColumnDef<TRow>>
  | RemoteTableSessionOptions<TRow, ColumnDef<TRow>>;

type OwnedTableSessionState<TRow> = {
  kind: "local" | "remote";
  raw: OwnedTableSession<TRow>;
  facade: OwnedTableSession<TRow>;
  cleanupToken?: object;
  cleanupDestroy: boolean;
  destroyed: boolean;
  revivable: boolean;
};

export function useTableSession<TRow>(
  options: LocalRecordTableSessionOptions<TRow, ColumnDef<TRow>>
): RecordTableSession<TRow, ColumnDef<TRow>>;
export function useTableSession<TRow>(
  options: RemoteTableSessionOptions<TRow, ColumnDef<TRow>>
): RemoteTableSession<TRow, ColumnDef<TRow>>;
export function useTableSession<TRow>(
  options: TableSessionOptions<TRow>
): OwnedTableSession<TRow> {
  const sessionRef = useRef<OwnedTableSessionState<TRow> | null>(null);
  const renderingRef = useRef(false);
  const onStateChangeRef = useRef(options.onStateChange);
  const pendingCorrectionRef = useRef<{
    updater: TableStateUpdater;
    context: ChangeContext;
  } | null>(null);
  const deferredOnStateChange = useRef((updater: TableStateUpdater, context: ChangeContext) => {
    const onStateChange = onStateChangeRef.current;
    if (!onStateChange) return;
    if (renderingRef.current) {
      pendingCorrectionRef.current = { updater, context };
      return;
    }
    onStateChange(updater, context);
  }).current;
  onStateChangeRef.current = options.onStateChange;
  const kind = options.source.kind;
  const renderOptions = options.onStateChange
    ? { ...options, onStateChange: deferredOnStateChange } as TableSessionOptions<TRow>
    : options;

  renderingRef.current = true;
  try {
    if (!sessionRef.current || sessionRef.current.kind !== kind) {
      sessionRef.current = createOwnedTableSession(renderOptions);
    } else {
      const ownership = sessionRef.current;
      if (ownership.destroyed && ownership.revivable) {
        if (kind === "local") {
          const local = ownership.raw as RecordTableSession<TRow, ColumnDef<TRow>>;
          reviveLocalRecordTableSession(local);
          local.updateOptions(renderOptions as LocalRecordTableSessionOptions<TRow, ColumnDef<TRow>>);
        } else {
          ownership.raw = createRawTableSession(renderOptions);
        }
        ownership.destroyed = false;
        ownership.revivable = false;
        scheduleOwnedSessionDestroy(ownership, ownership.facade);
      } else if (kind === "remote") {
        (ownership.raw as RemoteTableSession<TRow, ColumnDef<TRow>>)
          .updateOptions(renderOptions as RemoteTableSessionOptions<TRow, ColumnDef<TRow>>);
      } else {
        (ownership.raw as RecordTableSession<TRow, ColumnDef<TRow>>)
          .updateOptions(renderOptions as LocalRecordTableSessionOptions<TRow, ColumnDef<TRow>>);
      }
    }
  } finally {
    renderingRef.current = false;
  }

  const ownership = sessionRef.current;
  const session = ownership.facade;

  useLayoutEffect(() => {
    ownership.cleanupToken = undefined;
  }, [ownership, session]);

  useEffect(() => {
    const correction = pendingCorrectionRef.current;
    pendingCorrectionRef.current = null;
    const onStateChange = onStateChangeRef.current;
    if (!correction || !onStateChange) return;
    try { onStateChange(correction.updater, correction.context); } catch { /* Host callbacks are isolated. */ }
  });

  useEffect(() => {
    ownership.cleanupToken = undefined;
    if (isRemoteSession(session)) session.start();
    return () => {
      if (isRemoteSession(session)) session.stop();
      scheduleOwnedSessionDestroy(ownership, session);
    };
  }, [ownership, session]);

  useEffect(() => {
    if (isRemoteSession(session)) session.start();
  });

  return session;
}

function createOwnedTableSession<TRow>(
  options: TableSessionOptions<TRow>
): OwnedTableSessionState<TRow> {
  const ownership: OwnedTableSessionState<TRow> = {
    kind: options.source.kind,
    raw: createRawTableSession(options),
    facade: undefined as unknown as OwnedTableSession<TRow>,
    cleanupDestroy: false,
    destroyed: false,
    revivable: false
  };
  const common = {
    getSnapshot: () => ownership.raw.getSnapshot(),
    subscribe: (listener: () => void) => ownership.raw.subscribe(listener),
    dispatch: (intent: Parameters<OwnedTableSession<TRow>["dispatch"]>[0]) => ownership.raw.dispatch(intent),
    refresh: () => ownership.raw.refresh(),
    undo: () => ownership.raw.undo(),
    redo: () => ownership.raw.redo(),
    export: (exportOptions: Parameters<OwnedTableSession<TRow>["export"]>[0]) => ownership.raw.export(exportOptions),
    destroy: () => {
      if (ownership.destroyed) return;
      ownership.raw.destroy();
      ownership.destroyed = true;
      ownership.revivable = ownership.cleanupDestroy;
    }
  };
  ownership.facade = ownership.kind === "remote"
    ? {
        ...common,
        updateOptions: (next: RemoteTableSessionOptions<TRow, ColumnDef<TRow>>) =>
          (ownership.raw as RemoteTableSession<TRow, ColumnDef<TRow>>).updateOptions(next),
        start: () => (ownership.raw as RemoteTableSession<TRow, ColumnDef<TRow>>).start(),
        stop: () => (ownership.raw as RemoteTableSession<TRow, ColumnDef<TRow>>).stop(),
        getDiagnostics: () =>
          (ownership.raw as RemoteTableSession<TRow, ColumnDef<TRow>>).getDiagnostics()
      } as RemoteTableSession<TRow, ColumnDef<TRow>>
    : {
        ...common,
        updateOptions: (next: LocalRecordTableSessionOptions<TRow, ColumnDef<TRow>>) =>
          (ownership.raw as RecordTableSession<TRow, ColumnDef<TRow>>).updateOptions(next)
      } as RecordTableSession<TRow, ColumnDef<TRow>>;
  return ownership;
}

function scheduleOwnedSessionDestroy<TRow>(
  ownership: OwnedTableSessionState<TRow>,
  session: OwnedTableSession<TRow>
): void {
  const token = {};
  ownership.cleanupToken = token;
  queueMicrotask(() => {
    if (ownership.cleanupToken !== token) return;
    ownership.cleanupToken = undefined;
    ownership.cleanupDestroy = true;
    session.destroy();
    ownership.cleanupDestroy = false;
  });
}

function createRawTableSession<TRow>(options: TableSessionOptions<TRow>): OwnedTableSession<TRow> {
  return options.source.kind === "remote"
    ? createRemoteTableSession(options as RemoteTableSessionOptions<TRow, ColumnDef<TRow>>)
    : createLocalRecordTableSession(options as LocalRecordTableSessionOptions<TRow, ColumnDef<TRow>>);
}

function isRemoteSession<TRow>(
  session: OwnedTableSession<TRow>
): session is RemoteTableSession<TRow, ColumnDef<TRow>> {
  return "start" in session && "stop" in session;
}
