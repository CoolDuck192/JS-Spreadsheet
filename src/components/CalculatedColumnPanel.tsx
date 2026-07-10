import { useEffect, useId, useRef, useState } from "react";
import type { StructuredTableColumn } from "../types";

export type CalculatedColumnPanelProps = {
  columns: readonly StructuredTableColumn[];
  activeColumnId?: string;
  issue?: string;
  onApply(columnId: string, formula?: string): void;
  onClose?(): void;
};

export function CalculatedColumnPanel({
  columns,
  activeColumnId,
  issue,
  onApply,
  onClose
}: CalculatedColumnPanelProps) {
  const fallbackColumnId = activeColumnId && columns.some((column) => column.id === activeColumnId)
    ? activeColumnId
    : columns[0]?.id ?? "";
  const [columnId, setColumnId] = useState(fallbackColumnId);
  const selectedFormula = columns.find((column) => column.id === columnId)?.calculatedFormula ?? "";
  const [formulaDraft, setFormulaDraft] = useState(
    columns.find((column) => column.id === fallbackColumnId)?.calculatedFormula ?? ""
  );
  const [error, setError] = useState("");
  const formulaRef = useRef<HTMLInputElement>(null);
  const errorId = `calculated-column-error-${useId()}`;
  const displayedError = error || issue;

  useEffect(() => {
    setColumnId(fallbackColumnId);
  }, [fallbackColumnId]);

  useEffect(() => {
    setFormulaDraft(selectedFormula);
    setError("");
  }, [columnId, selectedFormula]);

  function applyFormula() {
    const nextFormula = formulaDraft.trim();
    if (!nextFormula.startsWith("=")) {
      setError("Formula must start with =");
      formulaRef.current?.focus();
      return;
    }
    if (!columnId) {
      setError("Choose a table column");
      return;
    }

    setError("");
    onApply(columnId, nextFormula);
  }

  return (
    <section className="structured-table-popover" aria-label="Calculated column">
      <label className="structured-table-field">
        <span>Calculated column</span>
        <select
          aria-label="Calculated column"
          value={columnId}
          onChange={(event) => setColumnId(event.currentTarget.value)}
        >
          {columns.map((column) => (
            <option key={column.id} value={column.id}>
              {column.name}
            </option>
          ))}
        </select>
      </label>
      <label className="structured-table-field">
        <span>Formula</span>
        <input
          ref={formulaRef}
          aria-label="Calculated column formula"
          aria-describedby={displayedError ? errorId : undefined}
          value={formulaDraft}
          onChange={(event) => {
            setFormulaDraft(event.currentTarget.value);
            setError("");
          }}
          onKeyDown={(event) => {
            if (event.key === "Enter") {
              event.preventDefault();
              applyFormula();
            }
          }}
          placeholder="=[@Quantity]*[@Price]"
        />
      </label>
      {displayedError ? (
        <p id={errorId} className="structured-table-field-error" role="alert">
          {displayedError}
        </p>
      ) : null}
      <div className="structured-table-actions">
        <button type="button" onClick={applyFormula} disabled={!columnId}>
          Apply calculated column
        </button>
        <button
          type="button"
          onClick={() => {
            if (columnId) {
              onApply(columnId, undefined);
              setFormulaDraft("");
              setError("");
            }
          }}
          disabled={!columnId}
        >
          Clear calculated column
        </button>
        {onClose ? (
          <button type="button" onClick={onClose}>
            Close calculated column
          </button>
        ) : null}
      </div>
    </section>
  );
}
