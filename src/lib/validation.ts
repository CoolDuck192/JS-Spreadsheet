import type { CellContent, DataValidationRule } from "../types";
import type { ComputedCellValue } from "./formulaEngine";

export type ValidationResult = { valid: true } | { valid: false; message: string };
export type ValidationCandidate = {
  raw: string;
  parsed: CellContent;
  evaluated: ComputedCellValue;
  formula?: string;
};

export function validateCellValue(value: string, rule: DataValidationRule | null | undefined): ValidationResult {
  return validateCellCandidate({ raw: value, parsed: value, evaluated: value }, rule);
}

export function validateCellCandidate(
  candidate: ValidationCandidate,
  rule: DataValidationRule | null | undefined
): ValidationResult {
  if (!rule) {
    return { valid: true };
  }

  if (candidate.formula && isComputedError(candidate.evaluated)) {
    return { valid: false, message: `Formula evaluates to ${candidate.evaluated.code}` };
  }

  const parsedText = candidate.parsed === null ? "" : String(candidate.parsed);
  const trimmedValue = parsedText.trim();
  if (trimmedValue === "" && !candidate.formula && rule.allowBlank !== false) {
    return { valid: true };
  }

  if (rule.type === "list") {
    if (rule.values.includes(parsedText) || rule.values.includes(trimmedValue)) {
      return { valid: true };
    }
    return { valid: false, message: `Choose one of: ${rule.values.join(", ")}` };
  }

  if (rule.type === "textLength") {
    const length = parsedText.length;
    if ((rule.min !== undefined && length < rule.min) || (rule.max !== undefined && length > rule.max)) {
      return { valid: false, message: textLengthValidationMessage(rule) };
    }
    return { valid: true };
  }

  const numericValue = numericCandidateValue(candidate, trimmedValue);
  if (numericValue === null) {
    return { valid: false, message: numberValidationMessage(rule) };
  }

  if ((rule.min !== undefined && numericValue < rule.min) || (rule.max !== undefined && numericValue > rule.max)) {
    return { valid: false, message: numberValidationMessage(rule) };
  }

  return { valid: true };
}

function numericCandidateValue(candidate: ValidationCandidate, parsedText: string): number | null {
  const value = candidate.formula ? candidate.evaluated : candidate.parsed;
  if (typeof value === "number") {
    return Number.isFinite(value) ? value : null;
  }
  if (typeof value !== "string" || value.trim() === "") {
    return null;
  }

  const parsed = Number(parsedText.replace(/[$,%]/g, "").replace(/,/g, ""));
  return Number.isFinite(parsed) ? parsed : null;
}

function isComputedError(value: ComputedCellValue): value is { kind: "error"; code: string } {
  return typeof value === "object" && value !== null && "kind" in value && value.kind === "error";
}

function textLengthValidationMessage(rule: Extract<DataValidationRule, { type: "textLength" }>): string {
  if (rule.min !== undefined && rule.max !== undefined) {
    return `Enter text between ${rule.min} and ${rule.max} characters`;
  }
  if (rule.min !== undefined) {
    return `Enter text with at least ${rule.min} characters`;
  }
  if (rule.max !== undefined) {
    return `Enter text with at most ${rule.max} characters`;
  }
  return "Enter text with a valid length";
}

function numberValidationMessage(rule: Extract<DataValidationRule, { type: "number" }>): string {
  if (rule.min !== undefined && rule.max !== undefined) {
    return `Enter a number between ${rule.min} and ${rule.max}`;
  }
  if (rule.min !== undefined) {
    return `Enter a number greater than or equal to ${rule.min}`;
  }
  if (rule.max !== undefined) {
    return `Enter a number less than or equal to ${rule.max}`;
  }
  return "Enter a valid number";
}
