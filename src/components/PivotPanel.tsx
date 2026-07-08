import { useEffect, useMemo, useState } from "react";
import { X } from "lucide-react";
import type { PivotAggregator, PivotConfig } from "../lib/pivot";

type PivotPanelProps = {
  headers: string[];
  sourceLabel?: string;
  sourceRowCount?: number;
  isOpen: boolean;
  onClose: () => void;
  onCreate: (config: PivotConfig) => void;
};

const AGGREGATORS: Array<{ value: PivotAggregator; label: string }> = [
  { value: "SUM", label: "Sum" },
  { value: "COUNT", label: "Count" },
  { value: "COUNTNUMS", label: "Count numbers" },
  { value: "AVERAGE", label: "Average" },
  { value: "MIN", label: "Min" },
  { value: "MAX", label: "Max" },
  { value: "PRODUCT", label: "Product" }
];

export function PivotPanel({ headers, sourceLabel, sourceRowCount, isOpen, onClose, onCreate }: PivotPanelProps) {
  // Key the defaults on header CONTENT, not array identity: the headers array is
  // recomputed on every selection change while the panel is open, and resetting
  // the user's field choices on an unrelated click would wipe their setup.
  const headersKey = headers.join("␟");
  // eslint-disable-next-line react-hooks/exhaustive-deps
  const defaults = useMemo(() => createDefaultConfig(headers), [headersKey]);
  const [rowField, setRowField] = useState(defaults.rowFields[0] ?? "");
  const [secondaryRowField, setSecondaryRowField] = useState("");
  const [columnField, setColumnField] = useState(defaults.columnField ?? "");
  const [valueField, setValueField] = useState(defaults.valueField);
  const [aggregator, setAggregator] = useState<PivotAggregator>(defaults.aggregator);

  useEffect(() => {
    setRowField(defaults.rowFields[0] ?? "");
    setSecondaryRowField("");
    setColumnField(defaults.columnField ?? "");
    setValueField(defaults.valueField);
    setAggregator(defaults.aggregator);
  }, [defaults]);

  if (!isOpen) {
    return null;
  }

  const rowFields = [rowField, secondaryRowField].filter(Boolean);
  const canCreate = headers.length > 0 && rowField && valueField && secondaryRowField !== rowField;

  return (
    <aside className="pivot-panel" aria-label="Pivot table builder">
      <div className="pivot-panel-header">
        <strong>Pivot Table</strong>
        <button type="button" aria-label="Close pivot table builder" onClick={onClose}>
          <X />
        </button>
      </div>
      {headers.length === 0 ? (
        <p>Select a source range with a header row first.</p>
      ) : (
        <>
          {sourceLabel ? (
            <p className="pivot-panel-source">
              Source <strong>{sourceLabel}</strong>
              {typeof sourceRowCount === "number" && sourceRowCount > 0
                ? ` · ${sourceRowCount} ${sourceRowCount === 1 ? "row" : "rows"}`
                : ""}
            </p>
          ) : null}
          <label>
            Rows
            <select
              aria-label="Pivot rows"
              value={rowField}
              onChange={(event) => {
                const nextRowField = event.currentTarget.value;
                setRowField(nextRowField);
                if (secondaryRowField === nextRowField) {
                  setSecondaryRowField("");
                }
              }}
            >
              {headers.map((header) => (
                <option key={header} value={header}>
                  {header}
                </option>
              ))}
            </select>
          </label>
          <label>
            Detail rows
            <select
              aria-label="Pivot row detail"
              value={secondaryRowField}
              onChange={(event) => setSecondaryRowField(event.currentTarget.value)}
            >
              <option value="">None</option>
              {headers.map((header) => (
                <option key={header} value={header} disabled={header === rowField}>
                  {header}
                </option>
              ))}
            </select>
          </label>
          <label>
            Columns
            <select
              aria-label="Pivot columns"
              value={columnField}
              onChange={(event) => setColumnField(event.currentTarget.value)}
            >
              <option value="">None</option>
              {headers.map((header) => (
                <option key={header} value={header}>
                  {header}
                </option>
              ))}
            </select>
          </label>
          <label>
            Values
            <select
              aria-label="Pivot values"
              value={valueField}
              onChange={(event) => setValueField(event.currentTarget.value)}
            >
              {headers.map((header) => (
                <option key={header} value={header}>
                  {header}
                </option>
              ))}
            </select>
          </label>
          <label>
            Aggregator
            <select
              aria-label="Pivot aggregator"
              value={aggregator}
              onChange={(event) => setAggregator(event.currentTarget.value as PivotAggregator)}
            >
              {AGGREGATORS.map((option) => (
                <option key={option.value} value={option.value}>
                  {option.label}
                </option>
              ))}
            </select>
          </label>
          <button
            type="button"
            className="primary-panel-button"
            disabled={!canCreate}
            onClick={() =>
              onCreate({
                rowFields,
                columnField: columnField || undefined,
                valueField,
                aggregator
              })
            }
          >
            Create pivot table
          </button>
          <p className="pivot-panel-hint">Double-click a pivot value later to drill into its source rows.</p>
        </>
      )}
    </aside>
  );
}

function createDefaultConfig(headers: string[]): PivotConfig {
  return {
    rowFields: headers[0] ? [headers[0]] : [],
    columnField: headers[1],
    valueField: headers[2] ?? headers.at(-1) ?? "",
    aggregator: "SUM"
  };
}
