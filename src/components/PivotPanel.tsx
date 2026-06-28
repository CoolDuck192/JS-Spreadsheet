import { useEffect, useMemo, useState } from "react";
import { X } from "lucide-react";
import type { PivotAggregator, PivotConfig } from "../lib/pivot";

type PivotPanelProps = {
  headers: string[];
  isOpen: boolean;
  onClose: () => void;
  onCreate: (config: PivotConfig) => void;
};

const AGGREGATORS: PivotAggregator[] = ["SUM", "COUNT", "AVERAGE", "MIN", "MAX"];

export function PivotPanel({ headers, isOpen, onClose, onCreate }: PivotPanelProps) {
  const defaults = useMemo(() => createDefaultConfig(headers), [headers]);
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
                <option key={option} value={option}>
                  {option}
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
