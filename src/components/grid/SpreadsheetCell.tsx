import { useEffect, useMemo, useState, type ReactNode } from "react";
import { ChevronDown, ListFilter } from "lucide-react";
import type { CellFormat, DataValidationRule } from "../../types";
import { foldDeterministicText } from "../../lib/filters";
import { getFormulaSuggestions, insertFormulaSuggestion } from "../../lib/formulaSuggestions";
import type { GridViewportRenderContext } from "../../react/viewport/types";
import { FormulaSuggestions, formulaSuggestionOptionId } from "../FormulaSuggestions";
import type { StructuredTableCellProjection } from "../Grid";

const AUTO_FILTER_DISPLAY_LIMIT = 200;

export type SpreadsheetMergeInfo = {
  role: "anchor" | "covered";
  columnSpan: number;
  rowSpan: number;
};

export type SpreadsheetAutoFilterChoice = {
  key: string;
  label: string;
  value: string;
};

export type SpreadsheetAutoFilter = {
  column: number;
  label: string;
  filtered: boolean;
  activeValues: readonly string[];
  loadChoices(): SpreadsheetAutoFilterChoice[];
  onApply(values: string[]): void;
  onClear(): void;
  onSort(direction: "asc" | "desc"): void;
};

export type SpreadsheetCellProps = {
  context: GridViewportRenderContext;
  mode: "display" | "editor";
  address: string;
  visibleValue: string;
  hyperlink: string;
  conditionalDataBar?: { percent: number; color: string } | null;
  format?: CellFormat;
  validation?: DataValidationRule | null;
  validationValue: string;
  mergeInfo?: SpreadsheetMergeInfo | null;
  ruleValues?: readonly (readonly string[])[];
  showValidationDropdown: boolean;
  editorValue?: string;
  autoFilter?: SpreadsheetAutoFilter | null;
  structuredTableCell?: StructuredTableCellProjection | null;
  interactionResetKey: unknown;
  onEditValueChange(value: string): void;
  onCommitEdit(address: string, value: string, move?: "down" | "up" | "right" | "left"): void;
  onCancelEdit(): void;
};

export function SpreadsheetCell(props: SpreadsheetCellProps): ReactNode {
  const {
    context,
    mode,
    address,
    visibleValue,
    hyperlink,
    conditionalDataBar,
    validation,
    validationValue,
    showValidationDropdown,
    editorValue = "",
    autoFilter,
    structuredTableCell,
    interactionResetKey,
    onEditValueChange,
    onCommitEdit,
    onCancelEdit
  } = props;
  const [activeSuggestionIndex, setActiveSuggestionIndex] = useState(0);
  const [dismissedSuggestion, setDismissedSuggestion] = useState<{ address: string; value: string } | null>(null);
  const [validationDropdownOpen, setValidationDropdownOpen] = useState(false);
  const [autoFilterOpen, setAutoFilterOpen] = useState(false);
  const [autoFilterDraft, setAutoFilterDraft] = useState<string[]>([]);
  const [autoFilterSearch, setAutoFilterSearch] = useState("");
  const [autoFilterChoices, setAutoFilterChoices] = useState<SpreadsheetAutoFilterChoice[]>([]);
  const suggestions = useMemo(
    () => (mode === "editor" ? getFormulaSuggestions(editorValue) : []),
    [editorValue, mode]
  );
  const suggestionKey = suggestions.map((suggestion) => suggestion.name).join("|");
  const visibleSuggestions =
    dismissedSuggestion?.address === address && dismissedSuggestion.value === editorValue ? [] : suggestions;
  const suggestionListId = `cell-editor-${address.toLowerCase()}-formula-suggestions`;
  const activeSuggestion = visibleSuggestions[wrapSuggestionIndex(activeSuggestionIndex, visibleSuggestions.length)];
  const matchingAutoFilterChoices = autoFilterSearch.trim()
    ? autoFilterChoices.filter((choice) =>
        foldDeterministicText(choice.label).includes(foldDeterministicText(autoFilterSearch.trim()))
      )
    : autoFilterChoices;
  const displayedAutoFilterChoices = matchingAutoFilterChoices.slice(0, AUTO_FILTER_DISPLAY_LIMIT);
  const autoFilterResultSummary = autoFilterSearch.trim()
    ? `Showing ${displayedAutoFilterChoices.length} of ${matchingAutoFilterChoices.length} matching values (${autoFilterChoices.length} total)`
    : `Showing ${displayedAutoFilterChoices.length} of ${autoFilterChoices.length} values`;

  useEffect(() => {
    setActiveSuggestionIndex(0);
  }, [address, editorValue, suggestionKey]);

  useEffect(() => {
    setDismissedSuggestion(null);
  }, [address]);

  useEffect(() => {
    setValidationDropdownOpen(false);
    setAutoFilterOpen(false);
    setAutoFilterDraft([]);
    setAutoFilterSearch("");
    setAutoFilterChoices([]);
  }, [interactionResetKey]);

  if (mode === "editor") {
    return (
      <div className="js-spreadsheet-cell-editor-shell" onMouseDown={(event) => event.stopPropagation()}>
        {validation?.type === "list" ? (
          <select
            className="js-spreadsheet-cell-editor"
            aria-label={`Cell editor ${address}`}
            autoFocus
            value={editorValue}
            onChange={(event) => onEditValueChange(event.currentTarget.value)}
            onBlur={() => onCommitEdit(address, editorValue)}
            onKeyDown={(event) => {
              if (event.key === "Enter") {
                event.preventDefault();
                onCommitEdit(address, editorValue, event.shiftKey ? "up" : "down");
              }
              if (event.key === "Escape") {
                event.preventDefault();
                onCancelEdit();
              }
            }}
          >
            {validation.allowBlank === false ? null : <option value="">(None)</option>}
            {validation.values.map((value) => (
              <option key={value} value={value}>
                {value}
              </option>
            ))}
          </select>
        ) : (
          <>
            <input
              className="js-spreadsheet-cell-editor"
              role="combobox"
              aria-label={`Cell editor ${address}`}
              aria-autocomplete="list"
              aria-controls={visibleSuggestions.length > 0 ? suggestionListId : undefined}
              aria-expanded={visibleSuggestions.length > 0}
              aria-activedescendant={
                activeSuggestion ? formulaSuggestionOptionId(suggestionListId, activeSuggestion.name) : undefined
              }
              autoFocus
              value={editorValue}
              onChange={(event) => {
                setDismissedSuggestion(null);
                onEditValueChange(event.currentTarget.value);
              }}
              onBlur={() => onCommitEdit(address, editorValue)}
              onKeyDown={(event) => {
                if (event.key === "ArrowDown" && visibleSuggestions.length > 0) {
                  event.preventDefault();
                  setActiveSuggestionIndex((current) =>
                    wrapSuggestionIndex(current + 1, visibleSuggestions.length)
                  );
                  return;
                }
                if (event.key === "ArrowUp" && visibleSuggestions.length > 0) {
                  event.preventDefault();
                  setActiveSuggestionIndex((current) =>
                    wrapSuggestionIndex(current - 1, visibleSuggestions.length)
                  );
                  return;
                }
                if (event.key === "Tab" && visibleSuggestions.length > 0) {
                  event.preventDefault();
                  const suggestion =
                    visibleSuggestions[wrapSuggestionIndex(activeSuggestionIndex, visibleSuggestions.length)];
                  if (suggestion) {
                    setDismissedSuggestion(null);
                    onEditValueChange(insertFormulaSuggestion(editorValue, suggestion.name));
                  }
                  return;
                }
                if (event.key === "Tab") {
                  event.preventDefault();
                  onCommitEdit(address, editorValue, event.shiftKey ? "left" : "right");
                  return;
                }
                if (event.key === "Enter") {
                  event.preventDefault();
                  onCommitEdit(address, editorValue, event.shiftKey ? "up" : "down");
                }
                if (event.key === "Escape") {
                  event.preventDefault();
                  if (visibleSuggestions.length > 0) {
                    setDismissedSuggestion({ address, value: editorValue });
                    return;
                  }
                  onCancelEdit();
                }
              }}
            />
            <div
              className="js-spreadsheet-cell-suggestion-layer"
              onMouseDown={(event) => event.stopPropagation()}
              onClick={(event) => event.stopPropagation()}
            >
              <FormulaSuggestions
                suggestions={visibleSuggestions}
                activeIndex={activeSuggestionIndex}
                listId={suggestionListId}
                onActiveIndexChange={setActiveSuggestionIndex}
                onSelect={(name) => {
                  setDismissedSuggestion(null);
                  onEditValueChange(insertFormulaSuggestion(editorValue, name));
                }}
              />
            </div>
          </>
        )}
      </div>
    );
  }

  return (
    <>
      {conditionalDataBar ? (
        <span
          className="js-spreadsheet-cell-data-bar"
          aria-hidden="true"
          style={{ width: `${conditionalDataBar.percent}%`, backgroundColor: conditionalDataBar.color }}
        />
      ) : null}
      {hyperlink ? (
        <a
          href={hyperlink}
          target="_blank"
          rel="noreferrer"
          onMouseDown={(event) => event.stopPropagation()}
          onClick={(event) => event.stopPropagation()}
        >
          {visibleValue}
        </a>
      ) : (
        <span>{visibleValue}</span>
      )}
      {autoFilter ? (
        <>
          <button
            type="button"
            className="js-spreadsheet-auto-filter-toggle"
            aria-label={`Open AutoFilter menu for ${autoFilter.label}`}
            aria-haspopup="menu"
            aria-expanded={autoFilterOpen}
            title={`AutoFilter menu for ${autoFilter.label}`}
            onMouseDown={(event) => {
              event.preventDefault();
              event.stopPropagation();
            }}
            onClick={(event) => {
              event.stopPropagation();
              if (autoFilterOpen) {
                closeAutoFilterMenu();
                return;
              }
              setAutoFilterChoices(autoFilter.loadChoices());
              setAutoFilterDraft([...autoFilter.activeValues]);
              setAutoFilterSearch("");
              setAutoFilterOpen(true);
            }}
            onKeyDown={(event) => {
              if (event.key === "Escape") {
                event.preventDefault();
                closeAutoFilterMenu();
              }
            }}
          >
            <ListFilter aria-hidden="true" />
          </button>
          {autoFilterOpen ? (
            <div
              className="js-spreadsheet-auto-filter-menu"
              role="menu"
              aria-label={`AutoFilter menu for ${autoFilter.label}`}
              onMouseDown={(event) => event.stopPropagation()}
              onClick={(event) => event.stopPropagation()}
              onKeyDown={(event) => {
                if (event.key === "Escape") {
                  event.preventDefault();
                  closeAutoFilterMenu();
                }
              }}
            >
              <button
                type="button"
                role="menuitem"
                onClick={() => {
                  autoFilter.onSort("asc");
                  closeAutoFilterMenu();
                }}
              >
                Sort A to Z
              </button>
              <button
                type="button"
                role="menuitem"
                onClick={() => {
                  autoFilter.onSort("desc");
                  closeAutoFilterMenu();
                }}
              >
                Sort Z to A
              </button>
              <div className="js-spreadsheet-auto-filter-menu-divider" />
              <button
                type="button"
                role="menuitem"
                disabled={!autoFilter.filtered}
                onClick={() => {
                  autoFilter.onClear();
                  closeAutoFilterMenu();
                }}
              >
                Clear filter from {autoFilter.label}
              </button>
              <div className="js-spreadsheet-auto-filter-menu-divider" />
              <input
                type="search"
                aria-label={`Search ${autoFilter.label} filter values`}
                value={autoFilterSearch}
                onChange={(event) => setAutoFilterSearch(event.currentTarget.value)}
              />
              <div role="status" aria-live="polite">
                {autoFilterResultSummary}
              </div>
              <div className="js-spreadsheet-auto-filter-choice-list">
                {displayedAutoFilterChoices.map((choice) => (
                  <button
                    key={choice.key}
                    type="button"
                    role="menuitemcheckbox"
                    aria-checked={autoFilterDraft.includes(choice.value)}
                    onClick={() => setAutoFilterDraft((current) => toggleFilterValue(current, choice.value))}
                  >
                    {choice.label}
                  </button>
                ))}
              </div>
              <button
                type="button"
                role="menuitem"
                onClick={() => {
                  autoFilter.onApply(autoFilterDraft);
                  closeAutoFilterMenu();
                }}
              >
                Apply selected values
              </button>
            </div>
          ) : null}
        </>
      ) : null}
      {structuredTableCell?.role === "header" && !autoFilter ? (
        <span
          className="js-spreadsheet-structured-table-filter-affordance"
          data-testid="structured-table-filter-affordance"
          aria-hidden="true"
          title="Table filter available from the Table tab"
        >
          <ListFilter />
        </span>
      ) : null}
      {showValidationDropdown && validation?.type === "list" ? (
        <>
          <button
            type="button"
            className="js-spreadsheet-validation-dropdown-toggle"
            aria-label={`Open validation choices for ${address}`}
            aria-haspopup="listbox"
            aria-expanded={validationDropdownOpen}
            title={`Validation choices for ${address}`}
            onMouseDown={(event) => {
              event.preventDefault();
              event.stopPropagation();
            }}
            onClick={(event) => {
              event.stopPropagation();
              setValidationDropdownOpen((current) => !current);
            }}
            onKeyDown={(event) => {
              if (event.key === "Escape") {
                event.preventDefault();
                setValidationDropdownOpen(false);
              }
            }}
          >
            <ChevronDown aria-hidden="true" />
          </button>
          {validationDropdownOpen ? (
            <div
              className="js-spreadsheet-validation-dropdown"
              role="listbox"
              aria-label={`Validation choices for ${address}`}
              onMouseDown={(event) => event.stopPropagation()}
              onClick={(event) => event.stopPropagation()}
              onKeyDown={(event) => {
                if (event.key === "Escape") {
                  event.preventDefault();
                  setValidationDropdownOpen(false);
                }
              }}
            >
              {validation.values.map((value, index) => (
                <button
                  key={`${value}-${index}`}
                  type="button"
                  role="option"
                  aria-selected={validationValue === value}
                  onClick={() => {
                    onCommitEdit(address, value);
                    setValidationDropdownOpen(false);
                  }}
                >
                  {value}
                </button>
              ))}
            </div>
          ) : null}
        </>
      ) : null}
    </>
  );

  function closeAutoFilterMenu() {
    setAutoFilterOpen(false);
    setAutoFilterDraft([]);
    setAutoFilterSearch("");
    setAutoFilterChoices([]);
  }
}

function wrapSuggestionIndex(index: number, suggestionCount: number): number {
  if (suggestionCount <= 0) {
    return 0;
  }
  return ((index % suggestionCount) + suggestionCount) % suggestionCount;
}

function toggleFilterValue(values: readonly string[], value: string): string[] {
  return values.includes(value) ? values.filter((current) => current !== value) : [...values, value];
}
