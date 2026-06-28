type StatusBarProps = {
  status: string;
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
  activeAddress,
  selectedCount,
  selectionSummary,
  formulaFunctions,
  zoomLevel,
  onZoomIn,
  onZoomOut,
  onResetZoom
}: StatusBarProps) {
  return (
    <div className="status-bar" aria-label="Status">
      <span>{status}</span>
      <span>{activeAddress}</span>
      <span>{selectedCount} selected</span>
      <span>{selectionSummary}</span>
      <span>{formulaFunctions}</span>
      <div className="status-zoom" aria-label="Worksheet zoom">
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
