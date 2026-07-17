import type { ReactElement } from "react";

type StatusBarProps = {
  status: string;
  persistence: {
    status: "idle" | "saving" | "failed";
    operation?: "load" | "save";
    message?: string;
  };
  activeAddress: string;
  selectedCount: number;
  selectionSummary: string;
  formulaFunctions: string;
  zoomLevel: number;
  onZoomIn: () => void;
  onZoomOut: () => void;
  onResetZoom: () => void;
};

export function StatusBar({
  status,
  persistence,
  activeAddress,
  selectedCount,
  selectionSummary,
  formulaFunctions,
  zoomLevel,
  onZoomIn,
  onZoomOut,
  onResetZoom
}: StatusBarProps): ReactElement {
  const persistenceMessage = persistence.status === "failed"
    ? persistence.message ?? "Workbook storage operation failed"
    : persistence.status === "saving"
      ? "Saving…"
      : null;
  return (
    <div className="js-spreadsheet-status-bar" aria-label="Status">
      <span>{status}</span>
      {persistenceMessage ? (
        <span
          className={`js-spreadsheet-status-persistence js-spreadsheet-status-persistence-${persistence.status}`}
          role={persistence.status === "failed" ? "alert" : "status"}
          aria-label={`Workbook storage status: ${persistenceMessage}`}
        >
          {persistenceMessage}
        </span>
      ) : null}
      <span>{activeAddress}</span>
      <span>{selectedCount} selected</span>
      <span>{selectionSummary}</span>
      <span>{formulaFunctions}</span>
      <div className="js-spreadsheet-status-zoom" aria-label="Worksheet zoom">
        <button type="button" aria-label="Zoom out" title="Zoom out" onClick={onZoomOut}>
          -
        </button>
        <button type="button" aria-label="Reset zoom" title="Reset zoom" onClick={onResetZoom}>
          {zoomLevel}%
        </button>
        <button type="button" aria-label="Zoom in" title="Zoom in" onClick={onZoomIn}>
          +
        </button>
      </div>
    </div>
  );
}
