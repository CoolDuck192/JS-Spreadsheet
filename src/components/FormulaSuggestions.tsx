import type { ReactElement } from "react";
import type { FormulaSuggestion } from "../lib/formulaSuggestions";

type FormulaSuggestionsProps = {
  suggestions: FormulaSuggestion[];
  activeIndex?: number;
  listId?: string;
  onActiveIndexChange?: (index: number) => void;
  onSelect: (name: string) => void;
};

export function FormulaSuggestions({
  suggestions,
  activeIndex = 0,
  listId,
  onActiveIndexChange,
  onSelect
}: FormulaSuggestionsProps): ReactElement | null {
  if (suggestions.length === 0) {
    return null;
  }

  const selectedIndex = Math.min(Math.max(activeIndex, 0), suggestions.length - 1);

  return (
    <div
      id={listId}
      className="js-spreadsheet-formula-suggestions js-spreadsheet-formula-suggestions--excel"
      role="listbox"
      aria-label="Formula suggestions"
      aria-orientation="vertical"
    >
      {suggestions.map((suggestion, index) => {
        const isSelected = selectedIndex === index;
        const optionClassName = [
          "js-spreadsheet-formula-suggestion-option",
          isSelected ? "js-spreadsheet-formula-suggestion-option--active" : "js-spreadsheet-formula-suggestion-option--idle"
        ].join(" ");

        return (
          <button
            key={suggestion.name}
            id={listId ? formulaSuggestionOptionId(listId, suggestion.name) : undefined}
            className={optionClassName}
            type="button"
            role="option"
            aria-selected={isSelected}
            aria-label={suggestion.name}
            title={`${suggestion.syntax} - ${suggestion.description}`}
            onFocus={() => onActiveIndexChange?.(index)}
            onMouseMove={() => onActiveIndexChange?.(index)}
            onMouseDown={(event) => event.preventDefault()}
            onClick={() => onSelect(suggestion.name)}
          >
            <span className="js-spreadsheet-formula-suggestion-main">
              <strong>{suggestion.name}</strong>
              <code>{suggestion.syntax}</code>
            </span>
            <span className="js-spreadsheet-formula-suggestion-description">{suggestion.description}</span>
          </button>
        );
      })}
    </div>
  );
}

export function formulaSuggestionOptionId(listId: string, suggestionName: string): string {
  return `${listId}-option-${suggestionName.toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/^-|-$/g, "")}`;
}
