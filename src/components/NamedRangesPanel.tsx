import { X } from "lucide-react";
import type { NamedRange, SheetModel } from "../types";
import { formatRangeAddress } from "../lib/autoSum";

type NamedRangesPanelProps = {
  isOpen: boolean;
  namedRanges: NamedRange[];
  sheets: SheetModel[];
  onClose: () => void;
  onSelect: (namedRange: NamedRange) => void;
  onDelete: (name: string) => void;
};

export function NamedRangesPanel({
  isOpen,
  namedRanges,
  sheets,
  onClose,
  onSelect,
  onDelete
}: NamedRangesPanelProps) {
  if (!isOpen) {
    return null;
  }

  return (
    <aside className="named-ranges-panel" aria-label="Named ranges">
      <div className="named-ranges-panel-header">
        <strong>Named Ranges</strong>
        <button type="button" aria-label="Close named ranges" onClick={onClose}>
          <X />
        </button>
      </div>
      <p className="named-ranges-summary">{namedRanges.length === 1 ? "1 named range" : `${namedRanges.length} named ranges`}</p>
      <div className="named-ranges-list" role="list" aria-label="Named range list">
        {namedRanges.length > 0 ? (
          namedRanges.map((namedRange) => (
            <article key={namedRange.name} className="named-range-item" role="listitem">
              <div>
                <strong>{namedRange.name}</strong>
                <code>{formatNamedRangeReference(namedRange, sheets)}</code>
              </div>
              <div className="named-range-actions">
                <button type="button" aria-label={`Select ${namedRange.name}`} onClick={() => onSelect(namedRange)}>
                  Select
                </button>
                <button type="button" aria-label={`Delete ${namedRange.name}`} onClick={() => onDelete(namedRange.name)}>
                  Delete
                </button>
              </div>
            </article>
          ))
        ) : (
          <p className="named-ranges-empty">No named ranges</p>
        )}
      </div>
    </aside>
  );
}

function formatNamedRangeReference(namedRange: NamedRange, sheets: SheetModel[]): string {
  const sheetName = sheets.find((sheet) => sheet.id === namedRange.sheetId)?.name ?? "Unknown";
  return `${sheetName}!${formatRangeAddress(namedRange.range)}`;
}
