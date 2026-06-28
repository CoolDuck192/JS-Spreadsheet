import { X } from "lucide-react";
import type { NamedRange, SheetModel } from "../types";
import { formatRangeAddress } from "../lib/autoSum";

type GoToPanelProps = {
  referenceValue: string;
  namedRanges: NamedRange[];
  sheets: SheetModel[];
  onReferenceChange: (value: string) => void;
  onGoToReference: () => void;
  onSelectNamedRange: (namedRange: NamedRange) => void;
  onClose: () => void;
};

export function GoToPanel({
  referenceValue,
  namedRanges,
  sheets,
  onReferenceChange,
  onGoToReference,
  onSelectNamedRange,
  onClose
}: GoToPanelProps) {
  return (
    <aside className="go-to-panel" aria-label="Go to">
      <div className="go-to-panel-header">
        <strong>Go To</strong>
        <button type="button" aria-label="Close go to" onClick={onClose}>
          <X />
        </button>
      </div>
      <form
        className="go-to-form"
        onSubmit={(event) => {
          event.preventDefault();
          onGoToReference();
        }}
      >
        <label>
          <span>Reference</span>
          <input
            aria-label="Go to reference"
            autoFocus
            value={referenceValue}
            onChange={(event) => onReferenceChange(event.currentTarget.value)}
          />
        </label>
        <button type="submit" className="primary-panel-button">
          Go
        </button>
      </form>
      <div className="go-to-list" role="list" aria-label="Go to named ranges">
        {namedRanges.length > 0 ? (
          namedRanges.map((namedRange) => (
            <article key={namedRange.name} className="go-to-item" role="listitem">
              <div>
                <strong>{namedRange.name}</strong>
                <code>{formatNamedRangeReference(namedRange, sheets)}</code>
              </div>
              <button type="button" aria-label={`Go to ${namedRange.name}`} onClick={() => onSelectNamedRange(namedRange)}>
                Go
              </button>
            </article>
          ))
        ) : (
          <p className="go-to-empty">No named ranges</p>
        )}
      </div>
    </aside>
  );
}

function formatNamedRangeReference(namedRange: NamedRange, sheets: SheetModel[]): string {
  const sheetName = sheets.find((sheet) => sheet.id === namedRange.sheetId)?.name ?? "Unknown";
  return `${sheetName}!${formatRangeAddress(namedRange.range)}`;
}
