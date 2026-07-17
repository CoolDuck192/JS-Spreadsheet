import { useState } from "react";
import { X } from "lucide-react";
import type { DataValidationRule } from "../types";
import { formatRangeAddress } from "../lib/autoSum";
import type { DataValidationSummary } from "../lib/dataValidationSummary";

type DataValidationPanelProps = {
  isOpen: boolean;
  onClose: () => void;
  onApply: (rule: DataValidationRule) => void;
  onClear: () => void;
  validationRules: DataValidationSummary[];
  onDeleteRule: (summary: DataValidationSummary) => void;
};

export function DataValidationPanel({
  isOpen,
  onClose,
  onApply,
  onClear,
  validationRules,
  onDeleteRule
}: DataValidationPanelProps) {
  const [type, setType] = useState<DataValidationRule["type"]>("list");
  const [listValues, setListValues] = useState("Open, Closed");
  const [minimum, setMinimum] = useState("");
  const [maximum, setMaximum] = useState("");

  if (!isOpen) {
    return null;
  }

  const listOptions = listValues
    .split(",")
    .map((value) => value.trim())
    .filter(Boolean);
  const parsedMinimum = parseOptionalNumber(minimum);
  const parsedMaximum = parseOptionalNumber(maximum);
  const hasValidBounds =
    parsedMinimum.isValid &&
    parsedMaximum.isValid &&
    (parsedMinimum.value === undefined || parsedMaximum.value === undefined || parsedMinimum.value <= parsedMaximum.value);
  const canApply = type === "number" || type === "textLength" ? hasValidBounds : listOptions.length > 0;

  return (
    <aside className="js-spreadsheet-validation-panel" aria-label="Data validation">
      <div className="js-spreadsheet-validation-panel-header">
        <strong>Data Validation</strong>
        <button type="button" aria-label="Close data validation" onClick={onClose}>
          <X />
        </button>
      </div>
      <label>
        <span>Type</span>
        <select
          aria-label="Validation type"
          value={type}
          onChange={(event) => setType(event.currentTarget.value as DataValidationRule["type"])}
        >
          <option value="list">List</option>
          <option value="number">Number</option>
          <option value="textLength">Text length</option>
        </select>
      </label>
      {type === "list" ? (
        <label>
          <span>List values</span>
          <input
            aria-label="List values"
            value={listValues}
            onChange={(event) => setListValues(event.currentTarget.value)}
          />
        </label>
      ) : (
        <div className="js-spreadsheet-validation-number-fields">
          <label>
            <span>{type === "textLength" ? "Min length" : "Min"}</span>
            <input
              aria-label={type === "textLength" ? "Minimum length" : "Minimum"}
              inputMode="decimal"
              value={minimum}
              onChange={(event) => setMinimum(event.currentTarget.value)}
            />
          </label>
          <label>
            <span>{type === "textLength" ? "Max length" : "Max"}</span>
            <input
              aria-label={type === "textLength" ? "Maximum length" : "Maximum"}
              inputMode="decimal"
              value={maximum}
              onChange={(event) => setMaximum(event.currentTarget.value)}
            />
          </label>
        </div>
      )}
      <div className="js-spreadsheet-validation-panel-actions">
        <button
          type="button"
          className="js-spreadsheet-primary-panel-button"
          disabled={!canApply}
          onClick={() => {
            if (type === "list") {
              onApply({ type: "list", values: listOptions });
              return;
            }

            if (type === "textLength") {
              onApply({
                type: "textLength",
                min: optionalNumberValue(parsedMinimum),
                max: optionalNumberValue(parsedMaximum)
              });
              return;
            }

            onApply({
              type: "number",
              min: optionalNumberValue(parsedMinimum),
              max: optionalNumberValue(parsedMaximum)
            });
          }}
        >
          Apply validation
        </button>
        <button type="button" onClick={onClear}>
          Clear validation
        </button>
      </div>
      <section className="js-spreadsheet-validation-rules-section" aria-label="Data validation rules">
        <h3>Existing rules</h3>
        <div className="js-spreadsheet-validation-rules-list" role="list">
          {validationRules.length > 0 ? (
            validationRules.map((summary) => {
              const rangeLabel = formatRangeAddress(summary.range);
              const ruleLabel = formatValidationRule(summary.rule);
              return (
                <article key={summary.id} className="js-spreadsheet-validation-rule-item" role="listitem">
                  <div>
                    <strong>{rangeLabel}</strong>
                    <span>{ruleLabel}</span>
                  </div>
                  <button
                    type="button"
                    aria-label={`Delete data validation rule ${rangeLabel} ${ruleLabel}`}
                    onClick={() => onDeleteRule(summary)}
                  >
                    Delete
                  </button>
                </article>
              );
            })
          ) : (
            <p className="js-spreadsheet-validation-rules-empty">No data validation rules</p>
          )}
        </div>
      </section>
    </aside>
  );
}

function formatValidationRule(rule: DataValidationRule): string {
  if (rule.type === "list") {
    return `List: ${rule.values.join(", ")}`;
  }

  if (rule.type === "textLength") {
    if (rule.min !== undefined && rule.max !== undefined) {
      return `Text length: ${rule.min} to ${rule.max}`;
    }

    if (rule.min !== undefined) {
      return `Text length: at least ${rule.min}`;
    }

    if (rule.max !== undefined) {
      return `Text length: at most ${rule.max}`;
    }

    return "Text length";
  }

  if (rule.min !== undefined && rule.max !== undefined) {
    return `Number: ${rule.min} to ${rule.max}`;
  }

  if (rule.min !== undefined) {
    return `Number: at least ${rule.min}`;
  }

  if (rule.max !== undefined) {
    return `Number: at most ${rule.max}`;
  }

  return "Number";
}

function parseOptionalNumber(value: string): { isValid: true; value: number | undefined } | { isValid: false } {
  const trimmedValue = value.trim();
  if (trimmedValue === "") {
    return { isValid: true, value: undefined };
  }

  const numberValue = Number(trimmedValue);
  return Number.isFinite(numberValue) ? { isValid: true, value: numberValue } : { isValid: false };
}

function optionalNumberValue(parsed: ReturnType<typeof parseOptionalNumber>): number | undefined {
  return parsed.isValid ? parsed.value : undefined;
}
