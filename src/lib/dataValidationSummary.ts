import type { CellRange, DataValidationRule } from "../types";
import { formatCellAddress, normalizeRange, parseCellAddress } from "./addressing";
import { formatRangeAddress } from "./autoSum";

export type DataValidationSummary = {
  id: string;
  range: CellRange;
  rule: DataValidationRule;
};

export function summarizeDataValidationRules(validations: Record<string, DataValidationRule>): DataValidationSummary[] {
  const entries = Object.entries(validations)
    .map(([address, rule]) => ({ address, coord: parseCellAddress(address), rule }))
    .sort((left, right) => left.coord.row - right.coord.row || left.coord.column - right.coord.column);
  const byAddress = new Map(entries.map((entry) => [entry.address, entry.rule]));
  const visited = new Set<string>();
  const summaries: DataValidationSummary[] = [];

  for (const entry of entries) {
    if (visited.has(entry.address)) {
      continue;
    }

    const range = expandValidationRectangle(entry.coord, entry.rule, byAddress, visited);
    for (const address of rectangleAddresses(range)) {
      visited.add(address);
    }
    summaries.push({
      id: `${formatRangeAddress(range)}|${validationRuleKey(entry.rule)}`,
      range,
      rule: cloneValidationRule(entry.rule)
    });
  }

  return summaries;
}

function expandValidationRectangle(
  start: { row: number; column: number },
  rule: DataValidationRule,
  byAddress: Map<string, DataValidationRule>,
  visited: Set<string>
): CellRange {
  let endColumn = start.column;
  while (hasSameUnvisitedRule({ row: start.row, column: endColumn + 1 }, rule, byAddress, visited)) {
    endColumn += 1;
  }

  let endRow = start.row;
  while (canExtendRow(endRow + 1, start.column, endColumn, rule, byAddress, visited)) {
    endRow += 1;
  }

  return normalizeRange({ start, end: { row: endRow, column: endColumn } });
}

function canExtendRow(
  row: number,
  startColumn: number,
  endColumn: number,
  rule: DataValidationRule,
  byAddress: Map<string, DataValidationRule>,
  visited: Set<string>
): boolean {
  for (let column = startColumn; column <= endColumn; column += 1) {
    if (!hasSameUnvisitedRule({ row, column }, rule, byAddress, visited)) {
      return false;
    }
  }

  return true;
}

function hasSameUnvisitedRule(
  coord: { row: number; column: number },
  rule: DataValidationRule,
  byAddress: Map<string, DataValidationRule>,
  visited: Set<string>
): boolean {
  const address = formatCellAddress(coord);
  return !visited.has(address) && validationRulesEqual(byAddress.get(address), rule);
}

function rectangleAddresses(range: CellRange): string[] {
  const normalized = normalizeRange(range);
  const addresses: string[] = [];
  for (let row = normalized.start.row; row <= normalized.end.row; row += 1) {
    for (let column = normalized.start.column; column <= normalized.end.column; column += 1) {
      addresses.push(formatCellAddress({ row, column }));
    }
  }
  return addresses;
}

function validationRulesEqual(left: DataValidationRule | undefined, right: DataValidationRule): boolean {
  if (!left || left.type !== right.type || left.allowBlank !== right.allowBlank) {
    return false;
  }

  if (left.type === "list" && right.type === "list") {
    return left.values.length === right.values.length && left.values.every((value, index) => value === right.values[index]);
  }

  if (left.type === "number" && right.type === "number") {
    return left.min === right.min && left.max === right.max;
  }

  if (left.type === "textLength" && right.type === "textLength") {
    return left.min === right.min && left.max === right.max;
  }

  return false;
}

function validationRuleKey(rule: DataValidationRule): string {
  if (rule.type === "list") {
    return `list|${rule.values.join(",")}`;
  }

  if (rule.type === "textLength") {
    return `textLength|${rule.min ?? ""}|${rule.max ?? ""}`;
  }

  return `number|${rule.min ?? ""}|${rule.max ?? ""}`;
}

function cloneValidationRule(rule: DataValidationRule): DataValidationRule {
  return rule.type === "list" ? { ...rule, values: [...rule.values] } : { ...rule };
}
