import { HyperFormula } from "hyperformula";

export type FormulaSuggestion = {
  name: string;
  syntax: string;
  description: string;
};

const MAX_FORMULA_SUGGESTIONS = 6;

const CURATED_FORMULA_DETAILS: Record<string, Omit<FormulaSuggestion, "name">> = {
  SUM: { syntax: "SUM(number1, [number2], ...)", description: "Adds numbers or ranges." },
  AVERAGE: { syntax: "AVERAGE(number1, [number2], ...)", description: "Returns the arithmetic mean." },
  COUNT: { syntax: "COUNT(value1, [value2], ...)", description: "Counts numeric values." },
  COUNTA: { syntax: "COUNTA(value1, [value2], ...)", description: "Counts non-empty values." },
  MIN: { syntax: "MIN(number1, [number2], ...)", description: "Returns the smallest number." },
  MAX: { syntax: "MAX(number1, [number2], ...)", description: "Returns the largest number." },
  IF: { syntax: "IF(logical_test, value_if_true, value_if_false)", description: "Returns one value when true and another when false." },
  VLOOKUP: { syntax: "VLOOKUP(search_key, range, index, [is_sorted])", description: "Finds a value in the first column of a range." },
  XLOOKUP: { syntax: "XLOOKUP(search_key, lookup_range, return_range, ...)", description: "Finds a value and returns a matching result." },
  NETWORKDAYS: { syntax: "NETWORKDAYS(start_date, end_date, [holidays])", description: "Counts workdays between two dates." },
  "NETWORKDAYS.INTL": { syntax: "NETWORKDAYS.INTL(start_date, end_date, [weekend], [holidays])", description: "Counts workdays with custom weekends." },
  CONCATENATE: { syntax: "CONCATENATE(text1, [text2], ...)", description: "Joins text values." },
  ROUND: { syntax: "ROUND(value, places)", description: "Rounds a number to a number of places." },
  ABS: { syntax: "ABS(value)", description: "Returns absolute value." },
  TODAY: { syntax: "TODAY()", description: "Returns the current date." }
};

const CURATED_PRIORITY = [
  "SUM",
  "AVERAGE",
  "COUNT",
  "COUNTA",
  "MIN",
  "MAX",
  "IF",
  "VLOOKUP",
  "XLOOKUP",
  "ROUND",
  "ABS",
  "TODAY"
];

export const FORMULA_SUGGESTIONS: FormulaSuggestion[] = createFormulaSuggestionCatalog();

export function getFormulaSuggestions(input: string): FormulaSuggestion[] {
  if (!input.startsWith("=")) {
    return [];
  }

  const formulaBody = input.slice(1);
  if (/^\$?[A-Za-z]+\$?\d/.test(formulaBody) || /^[A-Za-z.]+\(/.test(formulaBody)) {
    return [];
  }

  const prefix = getFormulaPrefix(input);
  if (prefix.length === 0) {
    return FORMULA_SUGGESTIONS.slice(0, MAX_FORMULA_SUGGESTIONS);
  }

  return FORMULA_SUGGESTIONS.filter((suggestion) => suggestion.name.startsWith(prefix)).slice(0, MAX_FORMULA_SUGGESTIONS);
}

export function searchFormulaCatalog(query: string, limit = 60): FormulaSuggestion[] {
  const normalizedQuery = query.trim().toUpperCase();
  const boundedLimit = Math.max(1, Math.floor(limit));
  if (!normalizedQuery) {
    return FORMULA_SUGGESTIONS.slice(0, boundedLimit);
  }

  return FORMULA_SUGGESTIONS.filter((suggestion) => formulaSuggestionMatches(suggestion, normalizedQuery)).slice(
    0,
    boundedLimit
  );
}

export function insertFormulaSuggestion(input: string, functionName: string): string {
  const normalizedName = functionName.toUpperCase();
  const matched = FORMULA_SUGGESTIONS.find((suggestion) => suggestion.name === normalizedName);
  if (!input.startsWith("=") || !matched) {
    return input;
  }

  return `=${matched.name}(`;
}

function getFormulaPrefix(input: string): string {
  const formulaBody = input.slice(1);
  const match = formulaBody.match(/^[A-Za-z.]*/);
  return (match?.[0] ?? "").toUpperCase();
}

function createFormulaSuggestionCatalog(): FormulaSuggestion[] {
  const registeredNames = new Set(HyperFormula.getRegisteredFunctionNames("enGB"));
  return [...registeredNames]
    .sort((left, right) => functionPriority(left) - functionPriority(right) || left.localeCompare(right))
    .map((name) => ({ name, ...formulaDetails(name) }));
}

function functionPriority(name: string): number {
  const priorityIndex = CURATED_PRIORITY.indexOf(name);
  return priorityIndex < 0 ? CURATED_PRIORITY.length : priorityIndex;
}

function formulaDetails(name: string): Omit<FormulaSuggestion, "name"> {
  return CURATED_FORMULA_DETAILS[name] ?? { syntax: `${name}(...)`, description: "HyperFormula function." };
}

function formulaSuggestionMatches(suggestion: FormulaSuggestion, query: string): boolean {
  return (
    suggestion.name.includes(query) ||
    suggestion.syntax.toUpperCase().includes(query) ||
    suggestion.description.toUpperCase().includes(query)
  );
}
