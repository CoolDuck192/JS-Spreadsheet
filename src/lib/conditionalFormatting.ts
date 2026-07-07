import type { CellFormat, ConditionalFormatRule } from "../types";

export type ConditionalFormatEvaluationOptions = {
  getRuleValues?: (rule: ConditionalFormatRule) => readonly string[];
};

export type ConditionalDataBar = {
  color: string;
  percent: number;
};

export function getConditionalFormatForValue(
  value: string,
  rules: readonly ConditionalFormatRule[],
  options: ConditionalFormatEvaluationOptions = {}
): CellFormat | null {
  const matchingFormat = rules.reduce<CellFormat | null>((format, rule) => {
    const ruleFormat = getConditionalRuleFormatForValue(value, rule, options);
    if (!ruleFormat) {
      return format;
    }
    return {
      ...(format ?? {}),
      ...ruleFormat
    };
  }, null);

  return matchingFormat && Object.keys(matchingFormat).length > 0 ? matchingFormat : null;
}

function getConditionalRuleFormatForValue(
  value: string,
  rule: ConditionalFormatRule,
  options: ConditionalFormatEvaluationOptions
): CellFormat | null {
  if (rule.condition.type === "colorScale") {
    return getColorScaleFormatForValue(value, rule, options);
  }

  return matchesConditionalRule(value, rule, options) ? rule.format : null;
}

export function getConditionalDataBarForValue(
  value: string,
  rules: readonly ConditionalFormatRule[],
  options: ConditionalFormatEvaluationOptions = {}
): ConditionalDataBar | null {
  const numericValue = parseNumericValue(value);
  if (numericValue === null) {
    return null;
  }

  return rules.reduce<ConditionalDataBar | null>((dataBar, rule) => {
    if (rule.condition.type !== "dataBar") {
      return dataBar;
    }

    const ruleValues = options.getRuleValues?.(rule) ?? [];
    let maxValue = Number.NEGATIVE_INFINITY;
    for (const candidate of ruleValues) {
      const parsed = parseNumericValue(candidate);
      if (parsed !== null && parsed > maxValue) {
        maxValue = parsed;
      }
    }

    if (!Number.isFinite(maxValue) || maxValue <= 0) {
      return dataBar;
    }

    return {
      color: rule.condition.color,
      percent: clampPercent(Math.round((numericValue / maxValue) * 100))
    };
  }, null);
}

export function matchesConditionalRule(
  value: string,
  rule: ConditionalFormatRule,
  options: ConditionalFormatEvaluationOptions = {}
): boolean {
  const trimmedValue = value.trim();

  switch (rule.condition.type) {
    case "blank":
      return trimmedValue === "";
    case "notBlank":
      return trimmedValue !== "";
    case "duplicate":
      return countComparableValues(trimmedValue, options.getRuleValues?.(rule) ?? []) > 1;
    case "unique":
      return countComparableValues(trimmedValue, options.getRuleValues?.(rule) ?? []) === 1;
    case "top":
      return matchesRankedNumericValue(trimmedValue, options.getRuleValues?.(rule) ?? [], rule.condition.count, "top");
    case "bottom":
      return matchesRankedNumericValue(trimmedValue, options.getRuleValues?.(rule) ?? [], rule.condition.count, "bottom");
    case "greaterThan": {
      const numericValue = parseNumericValue(trimmedValue);
      const firstRuleValue = parseNumericValue(rule.condition.value);
      return numericValue !== null && firstRuleValue !== null && numericValue > firstRuleValue;
    }
    case "lessThan": {
      const numericValue = parseNumericValue(trimmedValue);
      const firstRuleValue = parseNumericValue(rule.condition.value);
      return numericValue !== null && firstRuleValue !== null && numericValue < firstRuleValue;
    }
    case "equalTo": {
      const numericValue = parseNumericValue(trimmedValue);
      const firstRuleValue = parseNumericValue(rule.condition.value);
      if (numericValue !== null && firstRuleValue !== null) {
        return numericValue === firstRuleValue;
      }
      return trimmedValue.toLocaleLowerCase() === rule.condition.value.trim().toLocaleLowerCase();
    }
    case "between": {
      const numericValue = parseNumericValue(trimmedValue);
      const firstRuleValue = parseNumericValue(rule.condition.value);
      const secondRuleValue = parseNumericValue(rule.condition.secondValue);
      if (numericValue === null || firstRuleValue === null || secondRuleValue === null) {
        return false;
      }
      const min = Math.min(firstRuleValue, secondRuleValue);
      const max = Math.max(firstRuleValue, secondRuleValue);
      return numericValue >= min && numericValue <= max;
    }
    case "textContains":
      return trimmedValue.toLocaleLowerCase().includes(rule.condition.value.trim().toLocaleLowerCase());
    case "dataBar":
    case "colorScale":
      return parseNumericValue(trimmedValue) !== null;
  }
}

function parseNumericValue(value: string): number | null {
  const normalizedValue = value.trim();
  if (normalizedValue === "") {
    return null;
  }

  const parsed = Number(normalizedValue.replace(/[$,%]/g, "").replace(/,/g, ""));
  return Number.isFinite(parsed) ? parsed : null;
}

function countComparableValues(value: string, values: readonly string[]): number {
  const comparableValue = normalizeComparableValue(value);
  if (comparableValue === "") {
    return 0;
  }

  return values.filter((candidate) => normalizeComparableValue(candidate) === comparableValue).length;
}

function normalizeComparableValue(value: string): string {
  return value.trim().toLocaleLowerCase();
}

function matchesRankedNumericValue(
  value: string,
  values: readonly string[],
  count: number,
  direction: "top" | "bottom"
): boolean {
  const numericValue = parseNumericValue(value);
  const normalizedCount = Math.max(0, Math.floor(count));
  if (numericValue === null || normalizedCount < 1) {
    return false;
  }

  const sortedValues = values
    .map((candidate) => parseNumericValue(candidate.trim()))
    .filter((candidate): candidate is number => candidate !== null)
    .sort((left, right) => (direction === "top" ? right - left : left - right));

  if (sortedValues.length === 0) {
    return false;
  }

  const threshold = sortedValues[Math.min(normalizedCount, sortedValues.length) - 1];
  return direction === "top" ? numericValue >= threshold : numericValue <= threshold;
}

function clampPercent(percent: number): number {
  return Math.min(100, Math.max(0, percent));
}

function getColorScaleFormatForValue(
  value: string,
  rule: ConditionalFormatRule,
  options: ConditionalFormatEvaluationOptions
): CellFormat | null {
  if (rule.condition.type !== "colorScale") {
    return null;
  }

  const numericValue = parseNumericValue(value);
  const minColor = parseHexColor(rule.condition.minColor);
  const maxColor = parseHexColor(rule.condition.maxColor);
  if (numericValue === null || !minColor || !maxColor) {
    return null;
  }

  const numericValues = (options.getRuleValues?.(rule) ?? [])
    .map((candidate) => parseNumericValue(candidate))
    .filter((candidate): candidate is number => candidate !== null);

  if (numericValues.length === 0) {
    return null;
  }

  const minValue = Math.min(...numericValues);
  const maxValue = Math.max(...numericValues);
  const ratio = maxValue === minValue ? 1 : clampRatio((numericValue - minValue) / (maxValue - minValue));

  return {
    backgroundColor: formatHexColor({
      red: interpolateChannel(minColor.red, maxColor.red, ratio),
      green: interpolateChannel(minColor.green, maxColor.green, ratio),
      blue: interpolateChannel(minColor.blue, maxColor.blue, ratio)
    })
  };
}

function parseHexColor(color: string): { red: number; green: number; blue: number } | null {
  const normalized = color.trim();
  const match = normalized.match(/^#([0-9a-f]{6})$/i);
  if (!match) {
    return null;
  }

  const hex = match[1];
  return {
    red: Number.parseInt(hex.slice(0, 2), 16),
    green: Number.parseInt(hex.slice(2, 4), 16),
    blue: Number.parseInt(hex.slice(4, 6), 16)
  };
}

function formatHexColor(color: { red: number; green: number; blue: number }): string {
  return `#${toHexChannel(color.red)}${toHexChannel(color.green)}${toHexChannel(color.blue)}`;
}

function toHexChannel(value: number): string {
  return value.toString(16).padStart(2, "0");
}

function interpolateChannel(min: number, max: number, ratio: number): number {
  return Math.round(min + (max - min) * ratio);
}

function clampRatio(ratio: number): number {
  return Math.min(1, Math.max(0, ratio));
}
