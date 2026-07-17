import type { ReactElement } from "react";
import { X } from "lucide-react";

type FindReplacePanelProps = {
  findValue: string;
  replaceValue: string;
  onFindChange: (value: string) => void;
  onReplaceChange: (value: string) => void;
  onFindNext: () => void;
  onReplace: () => void;
  onReplaceAll: () => void;
  onClose: () => void;
};

export function FindReplacePanel({
  findValue,
  replaceValue,
  onFindChange,
  onReplaceChange,
  onFindNext,
  onReplace,
  onReplaceAll,
  onClose
}: FindReplacePanelProps): ReactElement {
  return (
    <aside className="js-spreadsheet-find-panel" aria-label="Find and replace panel">
      <div className="js-spreadsheet-find-panel-header">
        <strong>Find and Replace</strong>
        <button type="button" aria-label="Close find and replace" onClick={onClose}>
          <X />
        </button>
      </div>
      <label>
        <span>Find</span>
        <input
          aria-label="Find text"
          autoFocus
          value={findValue}
          onChange={(event) => onFindChange(event.currentTarget.value)}
          onKeyDown={(event) => {
            if (event.key === "Enter") {
              event.preventDefault();
              onFindNext();
            }
          }}
        />
      </label>
      <label>
        <span>Replace</span>
        <input
          aria-label="Replace text"
          value={replaceValue}
          onChange={(event) => onReplaceChange(event.currentTarget.value)}
        />
      </label>
      <div className="js-spreadsheet-find-panel-actions">
        <button type="button" onClick={onFindNext}>
          Find next
        </button>
        <button type="button" onClick={onReplace}>
          Replace
        </button>
        <button type="button" className="js-spreadsheet-primary-panel-button" onClick={onReplaceAll}>
          Replace all
        </button>
      </div>
    </aside>
  );
}
