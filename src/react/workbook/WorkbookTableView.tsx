import { useLayoutEffect, useRef, type ReactElement } from "react";
import type { TableCellRef, TableSession } from "../../table/core/types";
import type {
  WorkbookTableRow,
  WorkbookTableSession
} from "../../table/workbook/WorkbookTableSession";
import { DataTable } from "../DataTable";
import type { ColumnDef } from "../tableTypes";
import { useTableSnapshot } from "../useTableSnapshot";

export type WorkbookTableViewProps = {
  session: WorkbookTableSession;
  onClose(): void;
  onOpenInSpreadsheet?(cell?: TableCellRef): void;
};

export function WorkbookTableView({
  session,
  onClose,
  onOpenInSpreadsheet
}: WorkbookTableViewProps): ReactElement {
  const dialogRef = useRef<HTMLDialogElement>(null);
  const closeButtonRef = useRef<HTMLButtonElement>(null);
  const openerRef = useRef<HTMLElement | null>(
    typeof document !== "undefined" && document.activeElement instanceof HTMLElement
      ? document.activeElement
      : null
  );
  const snapshot = useTableSnapshot(session);
  const selectedCell = snapshot.selection?.focus
    ?? (snapshot.rows[0] && snapshot.columns[0]
      ? { rowId: snapshot.rows[0].id, columnId: snapshot.columns[0].id }
      : undefined);
  const unavailable = snapshot.status.phase === "error";

  function restoreOpenerFocus() {
    if (openerRef.current?.isConnected) {
      openerRef.current.focus({ preventScroll: true });
    }
  }

  function closeView() {
    const dialog = dialogRef.current;
    if (dialog?.open) {
      dialog.close();
    }
    restoreOpenerFocus();
    onClose();
  }

  useLayoutEffect(() => {
    const dialog = dialogRef.current;
    if (!dialog) return;
    if (!dialog.open) {
      dialog.showModal();
    }
    closeButtonRef.current?.focus({ preventScroll: true });

    return () => {
      if (dialog.open) {
        dialog.close();
      }
    };
  }, []);

  return (
    <dialog
      ref={dialogRef}
      className="js-spreadsheet-workbook-table-view"
      aria-modal="true"
      aria-label={`Workbook table ${session.tableId}`}
      onCancel={(event) => {
        event.preventDefault();
        closeView();
      }}
    >
      <header className="js-spreadsheet-workbook-table-view__header">
        <div>
          <h2>Table view</h2>
          <p>{snapshot.rowCount} visible row{snapshot.rowCount === 1 ? "" : "s"}</p>
        </div>
        <div className="js-spreadsheet-workbook-table-view__actions">
          {onOpenInSpreadsheet ? (
            <button
              type="button"
              disabled={unavailable || !selectedCell}
              onClick={() => onOpenInSpreadsheet(selectedCell)}
            >
              Open in Spreadsheet
            </button>
          ) : null}
          <button ref={closeButtonRef} type="button" onClick={closeView}>Close table view</button>
        </div>
      </header>
      <DataTable
        session={session as TableSession<WorkbookTableRow, ColumnDef<WorkbookTableRow>>}
        aria-label={`Table ${session.tableId}`}
        className="js-spreadsheet-workbook-table-view__table"
        layout="fluid"
        rowSelection="multiple"
        inlineFilters
        noDataMessage={unavailable ? "Table unavailable" : "No table rows"}
      />
    </dialog>
  );
}
