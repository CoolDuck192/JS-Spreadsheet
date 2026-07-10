import { useEffect, useRef } from "react";
import {
  createLocalRecordTableSession,
  type LocalRecordTableSessionOptions,
  type RecordTableSession
} from "../table/local/RecordTableSession";
import {
  createRemoteTableSession,
  type RemoteTableSession,
  type RemoteTableSessionOptions
} from "../table/remote/RemoteTableSession";
import type { ColumnDef } from "./tableTypes";

type OwnedTableSession<TRow> =
  | RecordTableSession<TRow, ColumnDef<TRow>>
  | RemoteTableSession<TRow, ColumnDef<TRow>>;

type TableSessionOptions<TRow> =
  | LocalRecordTableSessionOptions<TRow, ColumnDef<TRow>>
  | RemoteTableSessionOptions<TRow, ColumnDef<TRow>>;

type OwnershipState = {
  cleanupTokens: WeakMap<object, object>;
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
  const sessionRef = useRef<{
    kind: "local" | "remote";
    session: OwnedTableSession<TRow>;
  } | null>(null);
  const ownershipRef = useRef<OwnershipState>({ cleanupTokens: new WeakMap() });
  const kind = options.source.kind;

  if (!sessionRef.current || sessionRef.current.kind !== kind) {
    sessionRef.current = {
      kind,
      session: kind === "remote"
        ? createRemoteTableSession(options as RemoteTableSessionOptions<TRow, ColumnDef<TRow>>)
        : createLocalRecordTableSession(options as LocalRecordTableSessionOptions<TRow, ColumnDef<TRow>>)
    };
  } else if (kind === "remote") {
    (sessionRef.current.session as RemoteTableSession<TRow, ColumnDef<TRow>>)
      .updateOptions(options as RemoteTableSessionOptions<TRow, ColumnDef<TRow>>);
  } else {
    (sessionRef.current.session as RecordTableSession<TRow, ColumnDef<TRow>>)
      .updateOptions(options as LocalRecordTableSessionOptions<TRow, ColumnDef<TRow>>);
  }

  const session = sessionRef.current.session;

  useEffect(() => {
    const ownership = ownershipRef.current;
    ownership.cleanupTokens.delete(session);
    if (isRemoteSession(session)) session.start();
    return () => {
      if (isRemoteSession(session)) session.stop();
      const token = {};
      ownership.cleanupTokens.set(session, token);
      queueMicrotask(() => {
        if (ownership.cleanupTokens.get(session) === token) {
          ownership.cleanupTokens.delete(session);
          session.destroy();
        }
      });
    };
  }, [session]);

  useEffect(() => {
    if (isRemoteSession(session)) session.start();
  });

  return session;
}

function isRemoteSession<TRow>(
  session: OwnedTableSession<TRow>
): session is RemoteTableSession<TRow, ColumnDef<TRow>> {
  return "start" in session && "stop" in session;
}
