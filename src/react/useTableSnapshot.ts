import { useSyncExternalStore } from "react";
import type { TableSession, TableViewSnapshot } from "../table/core/types";

export function useTableSnapshot<TRow, TColumn>(
  session: TableSession<TRow, TColumn>
): TableViewSnapshot<TRow, TColumn> {
  return useSyncExternalStore(
    session.subscribe.bind(session),
    session.getSnapshot.bind(session),
    session.getSnapshot.bind(session)
  );
}
