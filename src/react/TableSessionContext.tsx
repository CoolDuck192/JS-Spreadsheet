import { createContext, useContext, type ReactNode } from "react";
import type { TableSession } from "../table/core/types";
import type { ColumnDef } from "./tableTypes";

const TableSessionContext = createContext<TableSession<unknown, unknown> | null>(null);

export function TableSessionProvider<TRow>({
  session,
  children
}: {
  session: TableSession<TRow, ColumnDef<TRow>>;
  children: ReactNode;
}) {
  return (
    <TableSessionContext.Provider value={session as TableSession<unknown, unknown>}>
      {children}
    </TableSessionContext.Provider>
  );
}
export function useTableSessionContext<TRow>(): TableSession<TRow, ColumnDef<TRow>> {
  const session = useContext(TableSessionContext);
  if (!session) {
    throw new Error("useTableSessionContext must be used inside TableSessionProvider");
  }
  return session as TableSession<TRow, ColumnDef<TRow>>;
}
