import { createContext, useContext, useSyncExternalStore, type ReactNode } from "react";
import type { WorkbookSession } from "../core/workbook/WorkbookSession";

const WorkbookSessionContext = createContext<WorkbookSession | null>(null);

export function WorkbookSessionProvider({
  session,
  children
}: Readonly<{ session: WorkbookSession; children: ReactNode }>) {
  return (
    <WorkbookSessionContext.Provider value={session}>
      {children}
    </WorkbookSessionContext.Provider>
  );
}

export function useWorkbookSessionContext() {
  const session = useContext(WorkbookSessionContext);
  if (!session) {
    throw new Error("useWorkbookSessionContext must be used inside WorkbookSessionProvider");
  }
  const snapshot = useSyncExternalStore(
    session.subscribe,
    session.getSnapshot,
    session.getSnapshot
  );
  return { session, snapshot } as const;
}
