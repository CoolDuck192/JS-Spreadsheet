import { useEffect, useId, useRef, useState } from "react";
import type {
  TableCellRef,
  TableSession,
  TableViewSnapshot
} from "../table/core/types";
import { downloadTableExport } from "./exportArtifact";
import { columnLabel } from "./DataTableCell";
import type { ColumnDef } from "./tableTypes";

type QuickToolField =
  | "numberFormat"
  | "bold"
  | "fillColor"
  | "validationList"
  | "comment"
  | "formula"
  | "readOnly";

type QuickToolValues = {
  numberFormat: "general" | "number" | "currency" | "percent" | "date" | "datetime";
  bold: boolean;
  fillColor: string;
  validationList: string;
  comment: string;
  formula: string;
  readOnly: boolean;
};

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
  const [quickToolValues, setQuickToolValues] = useState<QuickToolValues>(() =>
    quickToolValuesForSelection(snapshot, selectedCells)
  );
  const [touched, setTouched] = useState<Partial<Record<QuickToolField, true>>>({});
  const selectionKey = quickToolSelectionKey(selectedCells);
  const selectionKeyRef = useRef(selectionKey);
  selectionKeyRef.current = selectionKey;
  const { numberFormat, bold, fillColor, validationList, comment, formula, readOnly } = quickToolValues;
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
  const quickToolsId = `${reasonPrefix}-quick-tools`;
  const permissionDenied = selectedCells.some((cell) => {
    const snapshotCell = snapshot.getCell(cell.rowId, cell.columnId);
    return !snapshotCell.editable && snapshotCell.metadata.readOnly !== true;
  });
  const hiddenColumns = snapshot.columns.filter((column) => snapshot.state.columnVisibility[column.id] === false);
  const offsetPage = snapshot.state.pagination.kind === "offset"
    ? snapshot.state.pagination
    : null;
  const cursorPage = snapshot.state.pagination.kind === "cursor"
    ? snapshot.state.pagination
    : null;
  const infinitePage = snapshot.state.pagination.kind === "infinite"
    ? snapshot.state.pagination
    : null;

  useEffect(() => {
    resetQuickTools(snapshot, selectedCells);
  }, [selectionKey]);

  async function run(command: Parameters<typeof session.dispatch>[0]) {
    const result = await session.dispatch(command);
    if (result.status === "rejected") {
      onIssue(result.issues?.[0]?.message ?? "Table operation was rejected");
    }
    return result;
  }

  async function applyQuickTools() {
    if (selectedCells.length === 0 || !metadataState.enabled || permissionDenied) return;
    const appliedSelectionKey = selectionKey;
    const values = validationList.split(",").map((value) => value.trim()).filter(Boolean);
    const hasFormatUpdate = touched.numberFormat
      || touched.bold
      || touched.fillColor;
    const result = await run({
      type: "update-cell-metadata",
      updates: selectedCells.map((cell) => {
        const currentFormat = snapshot.getCell(cell.rowId, cell.columnId).metadata.format;
        const nextFormat = { ...currentFormat };
        if (touched.numberFormat) {
          nextFormat.numberFormat = numberFormat;
        }
        if (touched.bold) {
          nextFormat.bold = bold;
        }
        if (touched.fillColor) {
          if (fillColor) nextFormat.backgroundColor = fillColor;
          else delete nextFormat.backgroundColor;
        }
        return {
          ...cell,
          patch: {
            ...(hasFormatUpdate ? { format: nextFormat } : {}),
            ...(touched.validationList && validationState.enabled
              ? { validation: values.length > 0 ? { kind: "list" as const, values } : undefined }
              : {}),
            ...(touched.comment ? { comment: comment || undefined } : {}),
            ...(touched.formula && formulaState.enabled ? { formula: formula || undefined } : {}),
            ...(touched.readOnly ? { readOnly } : {})
          }
        };
      })
    });
    if (
      (result.status === "committed" || result.status === "pending")
      && selectionKeyRef.current === appliedSelectionKey
    ) {
      resetQuickTools(session.getSnapshot(), selectedCells);
    }
  }

  function resetQuickTools(
    sourceSnapshot: TableViewSnapshot<TRow, ColumnDef<TRow>>,
    sourceCells: readonly TableCellRef[]
  ) {
    setQuickToolValues(quickToolValuesForSelection(sourceSnapshot, sourceCells));
    setTouched({});
  }

  function toggleQuickTools() {
    if (!quickToolsOpen) {
      resetQuickTools(snapshot, selectedCells);
    }
    setQuickToolsOpen(!quickToolsOpen);
  }

  function updateQuickToolValue<TKey extends keyof QuickToolValues>(
    key: TKey,
    value: QuickToolValues[TKey]
  ) {
    setQuickToolValues((current) => ({ ...current, [key]: value }));
  }

  function markTouched(field: QuickToolField) {
    setTouched((current) => current[field] ? current : { ...current, [field]: true });
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
    <>
      <div className="js-spreadsheet-data-table__toolbar" role="toolbar" aria-label="Data table toolbar">
        <button className="js-spreadsheet-data-table__toolbar-icon-button" type="button" aria-label="Undo" disabled={!snapshot.canUndo} onClick={() => void run({ type: "undo" })}>
          Undo
        </button>
        <button className="js-spreadsheet-data-table__toolbar-icon-button" type="button" aria-label="Redo" disabled={!snapshot.canRedo} onClick={() => void run({ type: "redo" })}>
          Redo
        </button>
        <button className="js-spreadsheet-data-table__toolbar-icon-button" type="button" aria-label="Refresh" onClick={() => void session.refresh()}>Refresh</button>
        <button className="js-spreadsheet-data-table__toolbar-icon-button" type="button" aria-label="Export CSV" disabled={!exportState.enabled} aria-describedby={!exportState.enabled ? exportReasonId : undefined} onClick={() => void download("csv")}>Export CSV</button>
        <button className="js-spreadsheet-data-table__toolbar-icon-button" type="button" aria-label="Export XLSX" disabled={!exportState.enabled} aria-describedby={!exportState.enabled ? exportReasonId : undefined} onClick={() => void download("xlsx")}>Export XLSX</button>
        {!exportState.enabled ? <span className="js-spreadsheet-data-table__toolbar-summary" id={exportReasonId}>{exportState.reason}</span> : null}
        <span className="js-spreadsheet-data-table__toolbar-summary">{selectedCells.length === 1 ? "1 cell selected" : `${selectedCells.length} cells selected`}</span>
        <button className="js-spreadsheet-data-table__toolbar-icon-button" type="button" aria-label="Quick tools" aria-controls={quickToolsOpen ? quickToolsId : undefined} aria-expanded={quickToolsOpen} onClick={toggleQuickTools}>
          Quick tools
        </button>
        {offsetPage ? (
          <span className="js-spreadsheet-data-table__pagination">
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
        {cursorPage ? (
          <span className="js-spreadsheet-data-table__pagination">
            <button
              type="button"
              aria-label="Previous cursor page"
              disabled={snapshot.pageInfo.kind !== "cursor" || !snapshot.pageInfo.previousCursor}
              onClick={() => void run({
                type: "set-pagination",
                pagination: {
                  kind: "cursor",
                  cursor: snapshot.pageInfo.kind === "cursor" ? snapshot.pageInfo.previousCursor : undefined,
                  limit: cursorPage.limit
                }
              })}
            >
              Previous
            </button>
            <button
              type="button"
              aria-label="Next cursor page"
              disabled={snapshot.pageInfo.kind !== "cursor" || !snapshot.pageInfo.nextCursor}
              onClick={() => void run({
                type: "set-pagination",
                pagination: {
                  kind: "cursor",
                  cursor: snapshot.pageInfo.kind === "cursor" ? snapshot.pageInfo.nextCursor : undefined,
                  limit: cursorPage.limit
                }
              })}
            >
              Next
            </button>
          </span>
        ) : null}
        {infinitePage ? (
          <button
            type="button"
            aria-label="Load more rows"
            disabled={snapshot.pageInfo.kind !== "infinite" || !snapshot.pageInfo.nextCursor}
            onClick={() => void run({
              type: "set-pagination",
              pagination: {
                kind: "infinite",
                after: snapshot.pageInfo.kind === "infinite" ? snapshot.pageInfo.nextCursor : undefined,
                limit: infinitePage.limit
              }
            })}
          >
            Load more
          </button>
        ) : null}
        <span className="js-spreadsheet-data-table__toolbar-summary" aria-label="Table row total">
          {snapshot.totalRowCount.kind === "known"
            ? `${snapshot.totalRowCount.value} total rows`
            : "Total rows unknown"}
        </span>
        {hiddenColumns.map((column) => (
          <button key={column.id} type="button" aria-label={`Show column ${columnLabel(column)}`} onClick={() => onShowColumn(column.id)}>
            Show {columnLabel(column)}
          </button>
        ))}
      </div>
      {quickToolsOpen ? (
        <div id={quickToolsId} className="js-spreadsheet-data-table__quick-tools" aria-label="Quick tools drawer">
          <label>
            Number format
            <select aria-label="Number format" value={numberFormat} onChange={(event) => {
              markTouched("numberFormat");
              updateQuickToolValue("numberFormat", event.currentTarget.value as QuickToolValues["numberFormat"]);
            }}>
              {(["general", "number", "currency", "percent", "date", "datetime"] as const).map((value) => (
                <option key={value} value={value}>{value}</option>
              ))}
            </select>
          </label>
          <label><input type="checkbox" aria-label="Bold" checked={bold} onChange={(event) => { markTouched("bold"); updateQuickToolValue("bold", event.currentTarget.checked); }} /> Bold</label>
          <label>Fill color<input aria-label="Fill color" value={fillColor} onChange={(event) => { markTouched("fillColor"); updateQuickToolValue("fillColor", event.currentTarget.value); }} /></label>
          <label>Validation list<input aria-label="Validation list" disabled={!metadataState.enabled || !validationState.enabled || permissionDenied} aria-describedby={!validationState.enabled ? validationReasonId : undefined} value={validationList} onChange={(event) => { markTouched("validationList"); updateQuickToolValue("validationList", event.currentTarget.value); }} /></label>
          <label>Comment<input aria-label="Comment" value={comment} onChange={(event) => { markTouched("comment"); updateQuickToolValue("comment", event.currentTarget.value); }} /></label>
          <label>
            Formula
            <input
              aria-label="Formula"
              value={formula}
              disabled={!metadataState.enabled || !formulaState.enabled}
              aria-describedby={!formulaState.enabled ? formulaReasonId : !metadataState.enabled ? metadataReasonId : undefined}
              onChange={(event) => { markTouched("formula"); updateQuickToolValue("formula", event.currentTarget.value); }}
            />
          </label>
          {!formulaState.enabled ? (
            <span id={formulaReasonId}>{formulaState.reason}</span>
          ) : null}
          {!validationState.enabled ? <span id={validationReasonId}>{validationState.reason}</span> : null}
          <label><input type="checkbox" aria-label="Read only" checked={readOnly} onChange={(event) => { markTouched("readOnly"); updateQuickToolValue("readOnly", event.currentTarget.checked); }} /> Read only</label>
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
    </>
  );
}

function quickToolSelectionKey(selectedCells: readonly TableCellRef[]): string {
  return JSON.stringify(selectedCells.map((cell) => [cell.rowId, cell.columnId]));
}

function quickToolValuesForSelection<TRow>(
  snapshot: TableViewSnapshot<TRow, ColumnDef<TRow>>,
  selectedCells: readonly TableCellRef[]
): QuickToolValues {
  const focus = snapshot.selection?.focus;
  const target = (focus
    ? selectedCells.find((cell) => cell.rowId === focus.rowId && cell.columnId === focus.columnId)
    : undefined) ?? selectedCells[0];
  const cell = target ? snapshot.getCell(target.rowId, target.columnId) : null;
  const metadata = cell?.metadata;
  return {
    numberFormat: metadata?.format?.numberFormat ?? "general",
    bold: metadata?.format?.bold ?? false,
    fillColor: metadata?.format?.backgroundColor ?? "",
    validationList: metadata?.validation?.kind === "list" ? metadata.validation.values.join(", ") : "",
    comment: metadata?.comment ?? "",
    formula: metadata?.formula ?? cell?.formula ?? "",
    readOnly: metadata?.readOnly ?? false
  };
}
