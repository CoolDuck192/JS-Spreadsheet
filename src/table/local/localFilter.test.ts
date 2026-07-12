import { describe, expect, it } from "vitest";
import { createColumnHelper } from "../core/columnHelper";
import type { FilterExpression } from "../core/query";
import { matchesFilter } from "./localFilter";

type TemporalRow = { id: string; day: string; instant: string };

const helper = createColumnHelper<TemporalRow>();
const columns = [
  helper.accessor("day", { id: "day", header: "Day", dataType: "date" }),
  helper.accessor("instant", { id: "instant", header: "Instant", dataType: "datetime" })
] as const;
const columnMap = new Map(columns.map((column) => [column.id, column] as const));

function matches(row: TemporalRow, expression: FilterExpression): boolean {
  return matchesFilter(
    row,
    row.id,
    expression,
    columnMap,
    (current, _rowId, columnId) => columnId === "day" ? current.day : current.instant
  );
}

describe("matchesFilter temporal comparisons", () => {
  const row: TemporalRow = {
    id: "edited",
    day: "2026-03-05",
    instant: "2026-03-05T12:30:00.000Z"
  };

  it("matches equivalent datetime forms through comparison, set, and range filters", () => {
    const equivalentFilters: readonly FilterExpression[] = [
      {
        kind: "comparison",
        columnId: "instant",
        operator: "eq",
        value: { type: "datetime", value: "2026-03-05T12:30" }
      },
      {
        kind: "set",
        columnId: "instant",
        operator: "in",
        values: [{ type: "datetime", value: "2026-03-05T12:30:00Z" }]
      },
      {
        kind: "range",
        columnId: "instant",
        operator: "between",
        lower: { type: "datetime", value: "2026-03-05T07:30:00-05:00" },
        upper: { type: "datetime", value: "2026-03-05T07:30:00-05:00" }
      }
    ];

    for (const filter of equivalentFilters) {
      expect(matches(row, filter)).toBe(true);
    }
  });

  it("orders datetimes chronologically after normalization", () => {
    expect(matches(row, {
      kind: "comparison",
      columnId: "instant",
      operator: "gt",
      value: { type: "datetime", value: "2026-03-05T12:29:59Z" }
    })).toBe(true);
    expect(matches(row, {
      kind: "comparison",
      columnId: "instant",
      operator: "lt",
      value: { type: "datetime", value: "2026-03-05T07:31:00-05:00" }
    })).toBe(true);
  });

  it("rejects invalid values and parsed temporal-kind mismatches", () => {
    expect(matches(row, {
      kind: "comparison",
      columnId: "instant",
      operator: "eq",
      value: { type: "datetime", value: "not-a-datetime" }
    })).toBe(false);
    expect(matches({ ...row, instant: "not-a-datetime" }, {
      kind: "comparison",
      columnId: "instant",
      operator: "eq",
      value: { type: "datetime", value: "2026-03-05T12:30" }
    })).toBe(false);
    expect(matches(row, {
      kind: "comparison",
      columnId: "day",
      operator: "eq",
      value: { type: "date", value: "2026-03-05T12:30" }
    })).toBe(false);
    expect(matches(row, {
      kind: "comparison",
      columnId: "instant",
      operator: "eq",
      value: { type: "datetime", value: "2026-03-05" }
    })).toBe(false);
  });
});
