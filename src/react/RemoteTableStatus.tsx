import { useState, type ReactElement } from "react";
import type { ColumnDef, TableSession, TableViewSnapshot } from "../table/core/types";

export function RemoteTableStatus<TRow>({
  session,
  snapshot,
  onIssue
}: {
  session: TableSession<TRow, ColumnDef<TRow>>;
  snapshot: TableViewSnapshot<TRow, ColumnDef<TRow>>;
  onIssue(message: string): void;
}): ReactElement | null {
  const [retrying, setRetrying] = useState(false);

  async function retry() {
    if (retrying) return;
    setRetrying(true);
    const result = await session.dispatch({ type: "refresh" });
    if (result.status === "rejected") {
      onIssue(result.issues?.[0]?.message ?? "Could not reload rows");
    }
    setRetrying(false);
  }

  if (snapshot.status.phase === "loading") {
    return <div className="js-spreadsheet-data-table__status" role="status" aria-live="polite">Loading rows</div>;
  }
  if (snapshot.status.phase === "error") {
    return (
      <div className="js-spreadsheet-data-table__remote-error" role="alert">
        <span>{snapshot.status.message ?? "Could not load rows"}</span>
        <button type="button" disabled={retrying} onClick={() => void retry()}>
          Retry loading rows
        </button>
      </div>
    );
  }
  if (snapshot.pendingOperations.length > 0) {
    const count = snapshot.pendingOperations.length;
    return (
      <div className="js-spreadsheet-data-table__status" role="status" aria-live="polite">
        {count} table operation{count === 1 ? "" : "s"} pending
      </div>
    );
  }
  return null;
}
