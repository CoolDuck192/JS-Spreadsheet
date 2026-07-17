import { useMemo, useState } from "react";
import { X } from "lucide-react";
import {
  FORMULA_SUGGESTIONS,
  searchFormulaCatalog,
  type FormulaSuggestion
} from "../lib/formulaSuggestions";

type FunctionLibraryPanelProps = {
  isOpen: boolean;
  onClose: () => void;
  onInsert: (name: string) => void;
};

const FUNCTION_LIBRARY_RESULT_LIMIT = 60;

export function FunctionLibraryPanel({ isOpen, onClose, onInsert }: FunctionLibraryPanelProps) {
  const [query, setQuery] = useState("");
  const matches = useMemo(() => searchFormulaCatalog(query, FUNCTION_LIBRARY_RESULT_LIMIT), [query]);

  if (!isOpen) {
    return null;
  }

  return (
    <aside className="js-spreadsheet-function-library-panel" aria-label="Function library">
      <div className="js-spreadsheet-function-library-panel-header">
        <strong>Function Library</strong>
        <button type="button" aria-label="Close function library" onClick={onClose}>
          <X />
        </button>
      </div>
      <label>
        <span>Search</span>
        <input
          aria-label="Search functions"
          autoFocus
          value={query}
          onChange={(event) => setQuery(event.currentTarget.value)}
        />
      </label>
      <p className="js-spreadsheet-function-library-summary">
        {FORMULA_SUGGESTIONS.length} functions, {matches.length} shown
      </p>
      <div className="js-spreadsheet-function-library-results" role="list" aria-label="Function results">
        {matches.length > 0 ? (
          matches.map((suggestion) => (
            <FunctionLibraryItem key={suggestion.name} suggestion={suggestion} onInsert={onInsert} />
          ))
        ) : (
          <p className="js-spreadsheet-function-library-empty">No functions found</p>
        )}
      </div>
    </aside>
  );
}

function FunctionLibraryItem({
  suggestion,
  onInsert
}: {
  suggestion: FormulaSuggestion;
  onInsert: (name: string) => void;
}) {
  return (
    <article className="js-spreadsheet-function-library-item" role="listitem">
      <div>
        <strong>{suggestion.name}</strong>
        <code>{suggestion.syntax}</code>
        <p>{suggestion.description}</p>
      </div>
      <button type="button" aria-label={`Insert ${suggestion.name}`} onClick={() => onInsert(suggestion.name)}>
        Insert
      </button>
    </article>
  );
}
