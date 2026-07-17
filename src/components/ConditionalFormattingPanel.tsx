import { useState } from "react";
import { X } from "lucide-react";
import type { CellFormat, ConditionalFormatCondition, ConditionalFormatRule } from "../types";
import { formatRangeAddress } from "../lib/autoSum";

type ConditionalRuleType = ConditionalFormatCondition["type"];

type ConditionalFormattingPanelProps = {
  isOpen: boolean;
  onClose: () => void;
  onApply: (rule: {
    condition: ConditionalFormatCondition;
    format: CellFormat;
  }) => void;
  onClear: () => void;
  rules: ConditionalFormatRule[];
  onDeleteRule: (ruleId: string) => void;
};

export function ConditionalFormattingPanel({
  isOpen,
  onClose,
  onApply,
  onClear,
  rules,
  onDeleteRule
}: ConditionalFormattingPanelProps) {
  const [ruleType, setRuleType] = useState<ConditionalRuleType>("greaterThan");
  const [value, setValue] = useState("10");
  const [secondValue, setSecondValue] = useState("20");
  const [rank, setRank] = useState("10");
  const [fillColor, setFillColor] = useState("#fff1d6");
  const [textColor, setTextColor] = useState("#8a4b00");
  const [dataBarColor, setDataBarColor] = useState("#2f7d9f");
  const [colorScaleMinColor, setColorScaleMinColor] = useState("#f8fafc");
  const [colorScaleMaxColor, setColorScaleMaxColor] = useState("#0f766e");
  const [bold, setBold] = useState(true);

  if (!isOpen) {
    return null;
  }

  const conditionNeedsValue =
    ruleType !== "blank" &&
    ruleType !== "notBlank" &&
    ruleType !== "duplicate" &&
    ruleType !== "unique" &&
    ruleType !== "top" &&
    ruleType !== "bottom" &&
    ruleType !== "dataBar" &&
    ruleType !== "colorScale";
  const conditionNeedsRank = ruleType === "top" || ruleType === "bottom";
  const isDataBarRule = ruleType === "dataBar";
  const isColorScaleRule = ruleType === "colorScale";
  const rankValue = Number(rank);
  const isValidRank = Number.isInteger(rankValue) && rankValue > 0;
  const canApply =
    isColorScaleRule ||
    isDataBarRule ||
    (conditionNeedsRank && isValidRank) ||
    (!conditionNeedsRank && (!conditionNeedsValue || (value.trim() !== "" && (ruleType !== "between" || secondValue.trim() !== ""))));

  return (
    <aside className="js-spreadsheet-conditional-panel" aria-label="Conditional formatting">
      <div className="js-spreadsheet-conditional-panel-header">
        <strong>Conditional Formatting</strong>
        <button type="button" aria-label="Close conditional formatting" onClick={onClose}>
          <X />
        </button>
      </div>
      <label>
        <span>Rule</span>
        <select
          aria-label="Conditional rule"
          value={ruleType}
          onChange={(event) => setRuleType(event.currentTarget.value as ConditionalRuleType)}
        >
          <option value="greaterThan">Greater than</option>
          <option value="lessThan">Less than</option>
          <option value="equalTo">Equal to</option>
          <option value="between">Between</option>
          <option value="textContains">Text contains</option>
          <option value="blank">Blank</option>
          <option value="notBlank">Not blank</option>
          <option value="duplicate">Duplicate values</option>
          <option value="unique">Unique values</option>
          <option value="top">Top values</option>
          <option value="bottom">Bottom values</option>
          <option value="dataBar">Data bars</option>
          <option value="colorScale">Color scale</option>
        </select>
      </label>
      {conditionNeedsRank ? (
        <div className="js-spreadsheet-conditional-value-fields">
          <label>
            <span>Rank</span>
            <input
              aria-label="Conditional rank"
              type="number"
              min="1"
              step="1"
              value={rank}
              onChange={(event) => setRank(event.currentTarget.value)}
            />
          </label>
        </div>
      ) : null}
      {conditionNeedsValue ? (
        <div className="js-spreadsheet-conditional-value-fields">
          <label>
            <span>Value</span>
            <input
              aria-label="Conditional value"
              value={value}
              onChange={(event) => setValue(event.currentTarget.value)}
            />
          </label>
          {ruleType === "between" ? (
            <label>
              <span>And</span>
              <input
                aria-label="Second conditional value"
                value={secondValue}
                onChange={(event) => setSecondValue(event.currentTarget.value)}
              />
            </label>
          ) : null}
        </div>
      ) : null}
      {isDataBarRule ? (
        <div className="js-spreadsheet-conditional-style-fields">
          <label>
            <span>Bar</span>
            <input
              type="color"
              aria-label="Data bar color"
              value={dataBarColor}
              onInput={(event) => setDataBarColor(event.currentTarget.value)}
            />
          </label>
        </div>
      ) : isColorScaleRule ? (
        <div className="js-spreadsheet-conditional-style-fields">
          <label>
            <span>Minimum</span>
            <input
              type="color"
              aria-label="Minimum color"
              value={colorScaleMinColor}
              onInput={(event) => setColorScaleMinColor(event.currentTarget.value)}
            />
          </label>
          <label>
            <span>Maximum</span>
            <input
              type="color"
              aria-label="Maximum color"
              value={colorScaleMaxColor}
              onInput={(event) => setColorScaleMaxColor(event.currentTarget.value)}
            />
          </label>
        </div>
      ) : (
        <div className="js-spreadsheet-conditional-style-fields">
          <label>
            <span>Fill</span>
            <input
              type="color"
              aria-label="Conditional fill color"
              value={fillColor}
              onInput={(event) => setFillColor(event.currentTarget.value)}
            />
          </label>
          <label>
            <span>Text</span>
            <input
              type="color"
              aria-label="Conditional text color"
              value={textColor}
              onInput={(event) => setTextColor(event.currentTarget.value)}
            />
          </label>
          <label className="js-spreadsheet-conditional-check">
            <input
              type="checkbox"
              aria-label="Conditional bold"
              checked={bold}
              onChange={(event) => setBold(event.currentTarget.checked)}
            />
            <span>Bold</span>
          </label>
        </div>
      )}
      <div className="js-spreadsheet-conditional-panel-actions">
        <button
          type="button"
          className="js-spreadsheet-primary-panel-button"
          disabled={!canApply}
          onClick={() => {
            onApply({
              condition:
                ruleType === "dataBar"
                  ? { type: "dataBar", color: dataBarColor }
                  : ruleType === "colorScale"
                  ? { type: "colorScale", minColor: colorScaleMinColor, maxColor: colorScaleMaxColor }
                  : ruleType === "top" || ruleType === "bottom"
                  ? { type: ruleType, count: rankValue }
                  : ruleType === "blank" || ruleType === "notBlank" || ruleType === "duplicate" || ruleType === "unique"
                  ? { type: ruleType }
                  : ruleType === "between"
                  ? { type: "between", value: value.trim(), secondValue: secondValue.trim() }
                  : { type: ruleType, value: value.trim() },
              format: isDataBarRule || isColorScaleRule
                ? {}
                : {
                    backgroundColor: fillColor,
                    textColor,
                    bold
                  }
            });
          }}
        >
          Apply conditional format
        </button>
        <button type="button" onClick={onClear}>
          Clear conditional formats
        </button>
      </div>
      <section className="js-spreadsheet-conditional-rules-section" aria-label="Conditional format rules">
        <h3>Existing rules</h3>
        <div className="js-spreadsheet-conditional-rules-list" role="list">
          {rules.length > 0 ? (
            rules.map((rule) => {
              const rangeLabel = formatRangeAddress(rule.range);
              const conditionLabel = formatCondition(rule.condition);
              const previewBackground =
                rule.condition.type === "dataBar"
                  ? rule.condition.color
                  : rule.condition.type === "colorScale"
                  ? `linear-gradient(90deg, ${rule.condition.minColor}, ${rule.condition.maxColor})`
                  : rule.format.backgroundColor;
              const previewText = rule.condition.type === "dataBar" ? "Bar" : rule.condition.type === "colorScale" ? "Scale" : "Aa";
              return (
                <article key={rule.id} className="js-spreadsheet-conditional-rule-item" role="listitem">
                  <div>
                    <strong>{rangeLabel}</strong>
                    <span>{conditionLabel}</span>
                  </div>
                  <div className="js-spreadsheet-conditional-rule-preview" aria-hidden="true">
                    <span
                      style={{
                        backgroundColor: previewBackground,
                        color: rule.format.textColor,
                        fontWeight: rule.format.bold ? 700 : 500
                      }}
                    >
                      {previewText}
                    </span>
                  </div>
                  <button
                    type="button"
                    aria-label={`Delete conditional format rule ${rangeLabel} ${conditionLabel}`}
                    onClick={() => onDeleteRule(rule.id)}
                  >
                    Delete
                  </button>
                </article>
              );
            })
          ) : (
            <p className="js-spreadsheet-conditional-rules-empty">No conditional format rules</p>
          )}
        </div>
      </section>
    </aside>
  );
}

function formatCondition(condition: ConditionalFormatCondition): string {
  switch (condition.type) {
    case "greaterThan":
      return `Greater than ${condition.value}`;
    case "lessThan":
      return `Less than ${condition.value}`;
    case "equalTo":
      return `Equal to ${condition.value}`;
    case "between":
      return `Between ${condition.value} and ${condition.secondValue}`;
    case "textContains":
      return `Text contains ${condition.value}`;
    case "blank":
      return "Blank";
    case "notBlank":
      return "Not blank";
    case "duplicate":
      return "Duplicate values";
    case "unique":
      return "Unique values";
    case "top":
      return `Top ${condition.count} values`;
    case "bottom":
      return `Bottom ${condition.count} values`;
    case "dataBar":
      return "Data bars";
    case "colorScale":
      return "Color scale";
  }
}
