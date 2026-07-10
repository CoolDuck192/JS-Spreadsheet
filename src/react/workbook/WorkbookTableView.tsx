import type { ReactElement } from "react";
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
  const snapshot = useTableSnapshot(session);
  const selectedCell = snapshot.selection?.focus
    ?? (snapshot.rows[0] && snapshot.columns[0]
      ? { rowId: snapshot.rows[0].id, columnId: snapshot.columns[0].id }
      : undefined);
  const unavailable = snapshot.status.phase === "error";

  return (
    <section
      className="js-spreadsheet-workbook-table-view"
      role="dialog"
      aria-modal="true"
      aria-label={`Workbook table ${session.tableId}`}
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
          <button type="button" onClick={onClose}>Close table view</button>
        </div>
      </header>
      {unavailable ? (
        <div role="alert" className="js-spreadsheet-workbook-table-view__issue">
          {snapshot.status.message ?? "This table is no longer available"}
        </div>
      ) : null}
      <DataTable
        session={session as TableSession<WorkbookTableRow, ColumnDef<WorkbookTableRow>>}
        aria-label={`Table ${session.tableId}`}
        className="js-spreadsheet-workbook-table-view__table"
        layout="fluid"
        rowSelection="multiple"
        inlineFilters
        noDataMessage={unavailable ? "Table unavailable" : "No table rows"}
      />
    </section>
  );
}
