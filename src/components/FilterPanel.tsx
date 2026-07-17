import { useState } from "react";
import { X } from "lucide-react";
import type { FilterOperator } from "../types";

type FilterPanelProps = {
  isOpen: boolean;
  onClose: () => void;
  onApply: (filter: { operator: FilterOperator; value: string }) => void;
  onClear: () => void;
};

export function FilterPanel({ isOpen, onClose, onApply, onClear }: FilterPanelProps) {
  const [operator, setOperator] = useState<FilterOperator>("contains");
  const [value, setValue] = useState("");

  if (!isOpen) {
    return null;
  }

  return (
    <aside className="js-spreadsheet-filter-panel" aria-label="Filter">
      <div className="js-spreadsheet-filter-panel-header">
        <strong>Filter</strong>
        <button type="button" aria-label="Close filter" onClick={onClose}>
          <X />
        </button>
      </div>
      <label>
        <span>Operator</span>
        <select
          aria-label="Filter operator"
          value={operator}
          onChange={(event) => setOperator(event.currentTarget.value as FilterOperator)}
        >
          <option value="contains">Contains</option>
          <option value="equals">Equals</option>
          <option value="greaterThan">Greater than</option>
          <option value="lessThan">Less than</option>
        </select>
      </label>
      <label>
        <span>Value</span>
        <input
          aria-label="Filter value"
          autoFocus
          value={value}
          onChange={(event) => setValue(event.currentTarget.value)}
          onKeyDown={(event) => {
            if (event.key === "Enter" && value.trim() !== "") {
              event.preventDefault();
              onApply({ operator, value: value.trim() });
            }
          }}
        />
      </label>
      <div className="js-spreadsheet-filter-panel-actions">
        <button
          type="button"
          className="js-spreadsheet-primary-panel-button"
          disabled={value.trim() === ""}
          onClick={() => onApply({ operator, value: value.trim() })}
        >
          Apply filter
        </button>
        <button type="button" onClick={onClear}>
          Clear filters
        </button>
      </div>
    </aside>
  );
}
