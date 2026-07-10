import type { ColumnDef } from "../core/types";
import type { FilterExpression, QueryScalar } from "../core/query";

type EvaluationError = { kind: "error"; code: string };

export function matchesFilter<TRow>(
  row: TRow,
  rowId: string,
  expression: FilterExpression,
  columns: ReadonlyMap<string, ColumnDef<TRow, any>>,
  getValue: (row: TRow, rowId: string, columnId: string) => unknown
): boolean {
  if (expression.kind === "logical") {
    return expression.operator === "and"
      ? expression.operands.every((operand) => matchesFilter(row, rowId, operand, columns, getValue))
      : expression.operands.some((operand) => matchesFilter(row, rowId, operand, columns, getValue));
  }
  if (expression.kind === "not") {
    return !matchesFilter(row, rowId, expression.operand, columns, getValue);
  }

  const column = columns.get(expression.columnId);
  if (!column) {
    throw new Error(`Unknown column id: ${expression.columnId}`);
  }
  const value = getValue(row, rowId, expression.columnId);

  if (expression.kind === "blank") {
    const matched = expression.operator === "isNull" || expression.operator === "isNotNull"
      ? value === null
      : expression.operator === "isEmpty" || expression.operator === "isNotEmpty"
        ? value === ""
        : value === null || value === "";
    return expression.operator === "isNotNull"
      || expression.operator === "isNotEmpty"
      || expression.operator === "isNotBlank"
      ? !matched
      : matched;
  }

  if (expression.kind === "comparison") {
    if (expression.operator === "contains") {
      return stringOperation(value, expression.value, (left, right) => left.includes(right));
    }
    if (expression.operator === "startsWith") {
      return stringOperation(value, expression.value, (left, right) => left.startsWith(right));
    }
    if (expression.operator === "endsWith") {
      return stringOperation(value, expression.value, (left, right) => left.endsWith(right));
    }
    const comparison = compareToScalar(value, expression.value, column);
    if (comparison === null) return false;
    switch (expression.operator) {
      case "eq": return comparison === 0;
      case "neq": return comparison !== 0;
      case "gt": return comparison > 0;
      case "gte": return comparison >= 0;
      case "lt": return comparison < 0;
      case "lte": return comparison <= 0;
    }
  }

  if (expression.kind === "set") {
    const comparisons = expression.values.map((candidate) => compareToScalar(value, candidate, column));
    if (!comparisons.some((comparison) => comparison !== null)) return false;
    const matched = comparisons.some((comparison) => comparison === 0);
    return expression.operator === "in" ? matched : !matched;
  }

  const lower = compareToScalar(value, expression.lower, column);
  const upper = compareToScalar(value, expression.upper, column);
  if (lower === null || upper === null) return false;
  const matched = lower >= 0 && upper <= 0;
  return expression.operator === "between" ? matched : !matched;
}

function stringOperation(
  value: unknown,
  scalar: QueryScalar,
  compare: (left: string, right: string) => boolean
): boolean {
  return typeof value === "string"
    && scalar.type === "string"
    && compare(value.toLowerCase(), scalar.value.toLowerCase());
}

function compareToScalar<TRow>(
  value: unknown,
  scalar: QueryScalar,
  column: ColumnDef<TRow, any>
): number | null {
  if (scalar.type === "null") {
    return value === null ? 0 : null;
  }
  if (scalar.type === "error") {
    return isEvaluationError(value) && value.code === scalar.value ? 0 : null;
  }
  if (scalar.type === "number") {
    return typeof value === "number" && Number.isFinite(value) ? comparePrimitive(value, scalar.value) : null;
  }
  if (scalar.type === "boolean") {
    return typeof value === "boolean" ? comparePrimitive(Number(value), Number(scalar.value)) : null;
  }
  if (scalar.type === "date" || scalar.type === "datetime") {
    return typeof value === "string" && column.dataType === scalar.type
      ? comparePrimitive(value, scalar.value)
      : null;
  }
  return typeof value === "string" ? comparePrimitive(value.toLowerCase(), scalar.value.toLowerCase()) : null;
}

function comparePrimitive(left: number | string, right: number | string): number {
  if (left === right) return 0;
  return left < right ? -1 : 1;
}

function isEvaluationError(value: unknown): value is EvaluationError {
  return typeof value === "object" && value !== null && "kind" in value && value.kind === "error" && "code" in value;
}
