import { describe, expect, it } from "vitest";
import {
  deserializeQueryRequest,
  serializeQueryRequest,
  type FilterExpression,
  type QueryRequest,
  type QueryResult
} from "./query";

describe("table query AST", () => {
  it("round-trips a nested, typed, serializable query", () => {
    const filter: FilterExpression = {
      kind: "logical",
      operator: "and",
      operands: [
        {
          kind: "comparison",
          columnId: "salary",
          operator: "gte",
          value: { type: "number", value: 50_000 }
        },
        {
          kind: "not",
          operand: {
            kind: "blank",
            columnId: "started-at",
            operator: "isBlank"
          }
        }
      ]
    };
    const request: QueryRequest = {
      sorting: [{ columnId: "name", direction: "asc", nulls: "last" }],
      filter,
      grouping: [{ columnId: "department", direction: "asc" }],
      aggregates: [{ id: "salary-sum", columnId: "salary", function: "sum" }],
      pagination: { kind: "none" }
    };

    const encoded = serializeQueryRequest(request);

    expect(JSON.parse(encoded)).toEqual(request);
    expect(deserializeQueryRequest(encoded)).toEqual(request);
  });

  it("preserves explicit null, empty string, zero, and false literals", () => {
    const request: QueryRequest = {
      sorting: [],
      filter: {
        kind: "set",
        columnId: "value",
        operator: "in",
        values: [
          { type: "null" },
          { type: "string", value: "" },
          { type: "number", value: 0 },
          { type: "boolean", value: false },
          { type: "error", value: "#CYCLE!" }
        ]
      },
      grouping: [],
      aggregates: [],
      pagination: { kind: "none" }
    };

    expect(deserializeQueryRequest(serializeQueryRequest(request))).toEqual(request);
  });

  it("rejects malformed query JSON", () => {
    expect(() => deserializeQueryRequest('{"sorting":"wrong"}')).toThrow("Invalid table query");
  });

  it("carries renderer-neutral server group and aggregate projections", () => {
    const result: QueryResult<{ id: string; department: string }> = {
      items: [
        {
          kind: "group",
          id: "group-finance",
          depth: 0,
          columnId: "department",
          key: { type: "string", value: "Finance" },
          count: 2,
          aggregates: { "salary-sum": 300 }
        },
        { kind: "aggregate", id: "grand-total", depth: 0, aggregates: { "salary-sum": 300 } }
      ],
      revision: "r1",
      completeness: "completeDataset",
      pageInfo: { kind: "none", total: { kind: "known", value: 2 } }
    };
    expect(result.items.map((item) => item.kind)).toEqual(["group", "aggregate"]);
  });

  it("rejects ambiguous grouping combined with pagination", () => {
    const request: QueryRequest = {
      sorting: [],
      filter: null,
      grouping: [{ columnId: "department" }],
      aggregates: [],
      pagination: { kind: "offset", offset: 0, limit: 25 }
    };
    expect(() => deserializeQueryRequest(JSON.stringify(request))).toThrow(
      "Grouping requires pagination kind none"
    );
  });
});
