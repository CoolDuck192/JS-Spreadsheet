import { useId, useState } from "react";
import type {
  TableCellRef,
  TableSession,
  TableViewSnapshot
} from "../table/core/types";
import { downloadTableExport } from "./exportArtifact";
import { columnLabel } from "./DataTableCell";
import type { ColumnDef } from "./tableTypes";

export function DataTableToolbar<TRow>({
  session,
  snapshot,
  selectedCells,
  onShowColumn,
  onIssue
}: {
  session: TableSession<TRow, ColumnDef<TRow>>;
  snapshot: TableViewSnapshot<TRow, ColumnDef<TRow>>;
  selectedCells: readonly TableCellRef[];
  onShowColumn(columnId: string): void;
  onIssue(message: string): void;
}) {
  const [quickToolsOpen, setQuickToolsOpen] = useState(false);
  const [numberFormat, setNumberFormat] = useState("general");
  const [bold, setBold] = useState(false);
  const [fillColor, setFillColor] = useState("");
  const [validationList, setValidationList] = useState("");
  const [comment, setComment] = useState("");
  const [formula, setFormula] = useState("");
  const [readOnly, setReadOnly] = useState(false);
  const formulaState = snapshot.operationStates.formula;
  const metadataState = snapshot.operationStates.metadata;
  const validationState = snapshot.operationStates.validation;
  const exportState = snapshot.operationStates.export;
  const reasonPrefix = `js-spreadsheet-data-table-${useId().replace(/[^A-Za-z0-9_-]/g, "")}`;
  const formulaReasonId = `${reasonPrefix}-formula-reason`;
  const metadataReasonId = `${reasonPrefix}-metadata-reason`;
  const permissionReasonId = `${reasonPrefix}-permission-reason`;
  const validationReasonId = `${reasonPrefix}-validation-reason`;
  const exportReasonId = `${reasonPrefix}-export-reason`;
  const permissionDenied = selectedCells.some((cell) => {
    const snapshotCell = snapshot.getCell(cell.rowId, cell.columnId);
    return !snapshotCell.editable && snapshotCell.metadata.readOnly !== true;
  });
  const hiddenColumns = snapshot.columns.filter((column) => snapshot.state.columnVisibility[column.id] === false);
  const offsetPage = snapshot.state.pagination.kind === "offset"
    ? snapshot.state.pagination
    : null;

  async function run(command: Parameters<typeof session.dispatch>[0]) {
    const result = await session.dispatch(command);
    if (result.status === "rejected") {
      onIssue(result.issues?.[0]?.message ?? "Table operation was rejected");
    }
    return result;
  }

  async function applyQuickTools() {
    if (selectedCells.length === 0 || !metadataState.enabled || permissionDenied) return;
    const values = validationList.split(",").map((value) => value.trim()).filter(Boolean);
    await run({
      type: "update-cell-metadata",
      updates: selectedCells.map((cell) => ({
        ...cell,
        patch: {
          format: {
            numberFormat: numberFormat as "general" | "number" | "currency" | "percent" | "date" | "datetime",
            bold,
            ...(fillColor ? { backgroundColor: fillColor } : {})
          },
          ...(validationState.enabled && values.length > 0 ? { validation: { kind: "list" as const, values } } : {}),
          ...(comment ? { comment } : {}),
          ...(formulaState.enabled && formula ? { formula } : {}),
          readOnly
        }
      }))
    });
  }

  async function download(format: "csv" | "xlsx") {
    try {
      await downloadTableExport(session as TableSession<unknown>, {
        format,
        scope: exportState.scopeLabel === "Loaded rows" ? "currentView" : "completeDataset",
        fileName: `table.${format}`
      });
    } catch {
      onIssue(`Could not export ${format.toUpperCase()}`);
    }
  }

  return (
    <div className="js-spreadsheet-data-table__toolbar" role="toolbar" aria-label="Data table toolbar">
      <button type="button" aria-label="Undo" disabled={!snapshot.canUndo} onClick={() => void run({ type: "undo" })}>
        Undo
      </button>
      <button type="button" aria-label="Redo" disabled={!snapshot.canRedo} onClick={() => void run({ type: "redo" })}>
        Redo
      </button>
      <button type="button" aria-label="Refresh" onClick={() => void session.refresh()}>Refresh</button>
      <button type="button" aria-label="Export CSV" disabled={!exportState.enabled} aria-describedby={!exportState.enabled ? exportReasonId : undefined} onClick={() => void download("csv")}>Export CSV</button>
      <button type="button" aria-label="Export XLSX" disabled={!exportState.enabled} aria-describedby={!exportState.enabled ? exportReasonId : undefined} onClick={() => void download("xlsx")}>Export XLSX</button>
      {!exportState.enabled ? <span id={exportReasonId}>{exportState.reason}</span> : null}
      <span>{selectedCells.length === 1 ? "1 cell selected" : `${selectedCells.length} cells selected`}</span>
      <button type="button" aria-label="Quick tools" aria-expanded={quickToolsOpen} onClick={() => setQuickToolsOpen(!quickToolsOpen)}>
        Quick tools
      </button>
      {offsetPage ? (
        <span>
          <button
            type="button"
            aria-label="Previous page"
            disabled={offsetPage.offset === 0}
            onClick={() => void run({
              type: "set-pagination",
              pagination: { ...offsetPage, offset: Math.max(0, offsetPage.offset - offsetPage.limit) }
            })}
          >
            Previous
          </button>
          <button
            type="button"
            aria-label="Next page"
            disabled={snapshot.pageInfo.kind === "offset" && !snapshot.pageInfo.hasMore}
            onClick={() => void run({
              type: "set-pagination",
              pagination: { ...offsetPage, offset: offsetPage.offset + offsetPage.limit }
            })}
          >
            Next
          </button>
        </span>
      ) : null}
      {hiddenColumns.map((column) => (
        <button key={column.id} type="button" aria-label={`Show column ${columnLabel(column)}`} onClick={() => onShowColumn(column.id)}>
          Show {columnLabel(column)}
        </button>
      ))}
      {quickToolsOpen ? (
        <div className="js-spreadsheet-data-table__quick-tools" aria-label="Quick tools drawer">
          <label>
            Number format
            <select aria-label="Number format" value={numberFormat} onChange={(event) => setNumberFormat(event.currentTarget.value)}>
              {(["general", "number", "currency", "percent", "date", "datetime"] as const).map((value) => (
                <option key={value} value={value}>{value}</option>
              ))}
            </select>
          </label>
          <label><input type="checkbox" aria-label="Bold" checked={bold} onChange={(event) => setBold(event.currentTarget.checked)} /> Bold</label>
          <label>Fill color<input aria-label="Fill color" value={fillColor} onChange={(event) => setFillColor(event.currentTarget.value)} /></label>
          <label>Validation list<input aria-label="Validation list" disabled={!metadataState.enabled || !validationState.enabled || permissionDenied} aria-describedby={!validationState.enabled ? validationReasonId : undefined} value={validationList} onChange={(event) => setValidationList(event.currentTarget.value)} /></label>
          <label>Comment<input aria-label="Comment" value={comment} onChange={(event) => setComment(event.currentTarget.value)} /></label>
          <label>
            Formula
            <input
              aria-label="Formula"
              value={formula}
              disabled={!metadataState.enabled || !formulaState.enabled}
              aria-describedby={!formulaState.enabled ? formulaReasonId : !metadataState.enabled ? metadataReasonId : undefined}
              onChange={(event) => setFormula(event.currentTarget.value)}
            />
          </label>
          {!formulaState.enabled ? (
            <span id={formulaReasonId}>{formulaState.reason}</span>
          ) : null}
          {!validationState.enabled ? <span id={validationReasonId}>{validationState.reason}</span> : null}
          <label><input type="checkbox" aria-label="Read only" checked={readOnly} onChange={(event) => setReadOnly(event.currentTarget.checked)} /> Read only</label>
          {!metadataState.enabled ? <span id={metadataReasonId}>{metadataState.reason}</span> : null}
          {permissionDenied ? <span id={permissionReasonId}>Selection contains cells that do not permit metadata changes</span> : null}
          <button
            type="button"
            aria-label="Apply quick tools"
            aria-describedby={!metadataState.enabled ? metadataReasonId : permissionDenied ? permissionReasonId : undefined}
            onClick={() => void applyQuickTools()}
            disabled={selectedCells.length === 0 || !metadataState.enabled || permissionDenied}
          >
            Apply quick tools
          </button>
        </div>
      ) : null}
    </div>
  );
}
