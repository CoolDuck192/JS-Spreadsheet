import { useEffect, useRef } from "react";
import {
  createLocalRecordTableSession,
  type LocalRecordTableSessionOptions,
  type RecordTableSession
} from "../table/local/RecordTableSession";
import type { ColumnDef } from "./tableTypes";

type OwnershipState = {
  cleanupToken?: object;
};

export function useTableSession<TRow>(
  options: LocalRecordTableSessionOptions<TRow, ColumnDef<TRow>>
): RecordTableSession<TRow, ColumnDef<TRow>> {
  const sessionRef = useRef<RecordTableSession<TRow, ColumnDef<TRow>> | null>(null);
  const ownershipRef = useRef<OwnershipState>({});

  if (!sessionRef.current) {
    sessionRef.current = createLocalRecordTableSession(options);
  } else {
    sessionRef.current.updateOptions(options);
  }

  useEffect(() => {
    const ownership = ownershipRef.current;
    ownership.cleanupToken = undefined;
    return () => {
      const token = {};
      ownership.cleanupToken = token;
      queueMicrotask(() => {
        if (ownership.cleanupToken === token) {
          sessionRef.current?.destroy();
        }
      });
    };
  }, []);

  return sessionRef.current;
}
