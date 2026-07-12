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
    if ((direction === "min" && candidate < selected) || (direction === "max" && candidate > selected)) {
      selected = candidate;
    }
  }
  return selected;
}

function isComparable(value: unknown): value is number | string | boolean {
  return (typeof value === "number" && Number.isFinite(value)) || typeof value === "string" || typeof value === "boolean";
}

function isEvaluationError(value: unknown): value is { kind: "error"; code: string } {
  return typeof value === "object" && value !== null && "kind" in value && value.kind === "error";
}
