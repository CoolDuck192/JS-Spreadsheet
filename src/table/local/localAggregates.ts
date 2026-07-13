import type { TableAggregateRequest } from "../core/query";
import type { ColumnDef } from "../core/types";

type AggregateRow<TRow> = { row: TRow; rowId: string };

export function calculateLocalAggregates<TRow>(
  rows: readonly AggregateRow<TRow>[],
  requests: readonly TableAggregateRequest[],
  columns: ReadonlyMap<string, ColumnDef<TRow, any>>,
  getValue: (row: TRow, rowId: string, columnId: string) => unknown
): Readonly<Record<string, unknown>> {
  const result: Record<string, unknown> = {};
  for (const request of requests) {
    if (!columns.has(request.columnId)) {
      throw new Error(`Unknown column id: ${request.columnId}`);
    }
    const values = rows
      .map(({ row, rowId }) => getValue(row, rowId, request.columnId))
      .filter((value) => value !== null && value !== undefined && !isEvaluationError(value));
    switch (request.function) {
      case "count":
        result[request.id] = values.length;
        break;
      case "sum": {
        const numbers = values.filter((value): value is number => typeof value === "number" && Number.isFinite(value));
        result[request.id] = numbers.reduce((sum, value) => sum + value, 0);
        break;
      }
      case "average": {
        const numbers = values.filter((value): value is number => typeof value === "number" && Number.isFinite(value));
        result[request.id] = numbers.length === 0
          ? null
          : numbers.reduce((sum, value) => sum + value, 0) / numbers.length;
        break;
      }
      case "min":
        result[request.id] = extreme(values, "min");
        break;
      case "max":
        result[request.id] = extreme(values, "max");
        break;
    }
  }
  return result;
}

function extreme(values: readonly unknown[], direction: "min" | "max"): unknown {
  const comparable = values.filter(isComparable);
  if (comparable.length === 0) return null;
  let selected = comparable[0];
  for (let index = 1; index < comparable.length; index += 1) {
    const candidate = comparable[index];
    const comparison = compareComparable(candidate, selected);
    if ((direction === "min" && comparison < 0) || (direction === "max" && comparison > 0)) {
      selected = candidate;
    }
  }
  return selected;
}

function compareComparable(
  left: number | string | boolean,
  right: number | string | boolean
): number {
  if (typeof left !== typeof right) return comparableTypeOrder(left) - comparableTypeOrder(right);
  if (left === right) return 0;
  if (typeof left === "boolean" && typeof right === "boolean") {
    return Number(left) - Number(right);
  }
  return left < right ? -1 : 1;
}

function comparableTypeOrder(value: number | string | boolean): number {
  if (typeof value === "boolean") return 0;
  return typeof value === "number" ? 1 : 2;
}

function isComparable(value: unknown): value is number | string | boolean {
  return (typeof value === "number" && Number.isFinite(value)) || typeof value === "string" || typeof value === "boolean";
}

function isEvaluationError(value: unknown): value is { kind: "error"; code: string } {
  return typeof value === "object" && value !== null && "kind" in value && value.kind === "error";
}
