import { useEffect, useMemo, useState } from "react";
import { FormulaSuggestions, formulaSuggestionOptionId } from "./FormulaSuggestions";
import type { FormulaSuggestion } from "../lib/formulaSuggestions";

const FORMULA_BAR_SUGGESTION_LIST_ID = "formula-bar-suggestions";

type FormulaBarProps = {
  nameBoxValue: string;
  onNameBoxChange: (value: string) => void;
  onNameBoxCommit: () => void;
  onNameBoxCancel: () => void;
  value: string;
  onChange: (value: string) => void;
  onCommit: () => void;
  suggestions: FormulaSuggestion[];
  onSelectSuggestion: (name: string) => void;
};

export function FormulaBar({
  nameBoxValue,
  onNameBoxChange,
  onNameBoxCommit,
  onNameBoxCancel,
  value,
  onChange,
  onCommit,
  suggestions,
  onSelectSuggestion
}: FormulaBarProps) {
  const [activeSuggestionIndex, setActiveSuggestionIndex] = useState(0);
  const [dismissedSuggestionValue, setDismissedSuggestionValue] = useState<string | null>(null);
  const [isFormulaInputFocused, setIsFormulaInputFocused] = useState(false);
  const suggestionKey = useMemo(() => suggestions.map((suggestion) => suggestion.name).join("|"), [suggestions]);
  const visibleSuggestions = isFormulaInputFocused && dismissedSuggestionValue !== value ? suggestions : [];
  const activeSuggestion = visibleSuggestions[wrapSuggestionIndex(activeSuggestionIndex, visibleSuggestions.length)];

  useEffect(() => {
    setActiveSuggestionIndex(0);
  }, [suggestionKey, value]);

  function moveActiveSuggestion(delta: number) {
    if (visibleSuggestions.length === 0) {
      return;
    }

    setActiveSuggestionIndex((current) => wrapSuggestionIndex(current + delta, visibleSuggestions.length));
  }

  function commitActiveSuggestion() {
    const suggestion = visibleSuggestions[wrapSuggestionIndex(activeSuggestionIndex, visibleSuggestions.length)];
    if (suggestion) {
      setDismissedSuggestionValue(null);
      onSelectSuggestion(suggestion.name);
    }
  }

  return (
    // role="group" exposes the aria-label (ignored on generic divs) and keeps
    // the region distinguishable from the View tab's "Formula bar" toggle.
    <div className="js-spreadsheet-formula-bar" role="group" aria-label="Formula bar">
      <input
        className="js-spreadsheet-name-box"
        aria-label="Name box"
        title="Name box"
        value={nameBoxValue}
        onChange={(event) => onNameBoxChange(event.currentTarget.value)}
        onBlur={onNameBoxCancel}
        onKeyDown={(event) => {
          if (event.key === "Enter") {
            event.preventDefault();
            onNameBoxCommit();
          }
        }}
      />
      <div className="js-spreadsheet-formula-input-wrap">
        <input
          aria-label="Formula input"
          aria-autocomplete="list"
          aria-controls={visibleSuggestions.length > 0 ? FORMULA_BAR_SUGGESTION_LIST_ID : undefined}
          aria-expanded={visibleSuggestions.length > 0}
          aria-activedescendant={
            activeSuggestion ? formulaSuggestionOptionId(FORMULA_BAR_SUGGESTION_LIST_ID, activeSuggestion.name) : undefined
          }
          value={value}
          placeholder="Enter a value or formula"
          onFocus={() => setIsFormulaInputFocused(true)}
          onChange={(event) => {
            setDismissedSuggestionValue(null);
            onChange(event.currentTarget.value);
          }}
          onBlur={() => {
            setIsFormulaInputFocused(false);
            onCommit();
          }}
          onKeyDown={(event) => {
            // Excel semantics: Up/Down navigate suggestions, Tab accepts, Enter always
            // commits the input. Left/Right keep moving the text caret.
            if (event.key === "ArrowDown" && visibleSuggestions.length > 0) {
              event.preventDefault();
              moveActiveSuggestion(1);
              return;
            }

            if (event.key === "ArrowUp" && visibleSuggestions.length > 0) {
              event.preventDefault();
              moveActiveSuggestion(-1);
              return;
            }

            if (event.key === "Tab" && visibleSuggestions.length > 0) {
              event.preventDefault();
              commitActiveSuggestion();
              return;
            }

            if (event.key === "Escape" && visibleSuggestions.length > 0) {
              event.preventDefault();
              setDismissedSuggestionValue(value);
              return;
            }

            if (event.key === "Enter") {
              event.preventDefault();
              event.currentTarget.blur();
            }
          }}
        />
        <FormulaSuggestions
          suggestions={visibleSuggestions}
          activeIndex={activeSuggestionIndex}
          listId={FORMULA_BAR_SUGGESTION_LIST_ID}
          onActiveIndexChange={setActiveSuggestionIndex}
          onSelect={onSelectSuggestion}
        />
      </div>
    </div>
  );
}

function wrapSuggestionIndex(index: number, suggestionCount: number): number {
  if (suggestionCount <= 0) {
    return 0;
  }

  return ((index % suggestionCount) + suggestionCount) % suggestionCount;
}
