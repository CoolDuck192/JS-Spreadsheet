import type { DataValidationRule } from "../types";

export type ValidationResult = { valid: true } | { valid: false; message: string };

export function validateCellValue(value: string, rule: DataValidationRule | null | undefined): ValidationResult {
  if (!rule) {
    return { valid: true };
  }

  const trimmedValue = value.trim();
  if (trimmedValue === "" && rule.allowBlank !== false) {
    return { valid: true };
  }

  if (rule.type === "list") {
    if (rule.values.includes(value) || rule.values.includes(trimmedValue)) {
      return { valid: true };
    }
    return { valid: false, message: `Choose one of: ${rule.values.join(", ")}` };
  }

  if (rule.type === "textLength") {
    const length = value.length;
    if ((rule.min !== undefined && length < rule.min) || (rule.max !== undefined && length > rule.max)) {
      return { valid: false, message: textLengthValidationMessage(rule) };
    }
    return { valid: true };
  }

  const numericValue = Number(trimmedValue.replace(/[$,%]/g, "").replace(/,/g, ""));
  if (!Number.isFinite(numericValue)) {
    return { valid: false, message: numberValidationMessage(rule) };
  }

  if ((rule.min !== undefined && numericValue < rule.min) || (rule.max !== undefined && numericValue > rule.max)) {
    return { valid: false, message: numberValidationMessage(rule) };
  }

  return { valid: true };
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
