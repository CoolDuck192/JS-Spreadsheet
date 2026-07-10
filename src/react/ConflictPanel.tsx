import { useState, type ReactElement } from "react";
import type {
  ColumnDef,
  TableConflict,
  TableSession
} from "../table/core/types";

export function ConflictPanel<TRow>({
  conflict,
  session,
  columnName,
  onIssue
}: {
  conflict: TableConflict<TRow>;
  session: TableSession<TRow, ColumnDef<TRow>>;
  columnName: string;
  onIssue(message: string): void;
}): ReactElement {
  const [busy, setBusy] = useState(false);
  const [pendingOperationId, setPendingOperationId] = useState<string | null>(null);
  const operationStillPending = pendingOperationId !== null;

  async function resolve(kind: "reload" | "retry") {
    if (busy || operationStillPending) return;
    setBusy(true);
    const result = await session.dispatch(kind === "reload"
      ? {
          type: "reload-authoritative",
          operationId: conflict.operationId,
          rowId: conflict.rowId
        }
      : {
          type: "retry-with-revision",
          operationId: conflict.operationId,
          rowId: conflict.rowId,
          expectedRevision: conflict.revision
        });
    if (result.status === "pending") setPendingOperationId(result.operationId);
    if (result.status === "rejected") {
      onIssue(result.issues?.[0]?.message ?? "This conflict is no longer current");
    }
    setBusy(false);
  }

  const disabled = busy || operationStillPending;
  return (
    <div className="js-spreadsheet-data-table__conflict" role="alert">
      <p>{columnName} changed on the server for row {conflict.rowId}.</p>
      <div>
        <button type="button" disabled={disabled} onClick={() => void resolve("reload")}>
          Reload server value
        </button>
        <button type="button" disabled={disabled} onClick={() => void resolve("retry")}>
          Retry my change
        </button>
      </div>
    </div>
  );
}
