import { describe, expect, it } from "vitest";
import { createColumnHelper } from "../core/columnHelper";
import type { QueryRequest } from "../core/query";
import { buildLocalRowModel, type LocalValueResolver } from "./localRowModel";

type Row = {
  id: string;
  department: string;
  salary: number | null;
  active: boolean;
  note: string | null;
};

const rows: readonly Row[] = [
  { id: "1", department: "Finance", salary: 0, active: false, note: "" },
  { id: "2", department: "Finance", salary: 100, active: true, note: null },
  { id: "3", department: "IT", salary: 100, active: false, note: "Ada" }
];
const helper = createColumnHelper<Row>();
const columns = [
  helper.accessor("department", { id: "department", header: "Department" }),
  helper.accessor("salary", { id: "salary", header: "Salary", dataType: "number" }),
  helper.accessor("active", { id: "active", header: "Active", dataType: "boolean" }),
  helper.accessor("note", { id: "note", header: "Note" })
];

function query(overrides: Partial<QueryRequest>): QueryRequest {
  return {
    sorting: [], filter: null, grouping: [], aggregates: [], pagination: { kind: "none" }, ...overrides
  };
}

describe("buildLocalRowModel", () => {
  it("distinguishes null, empty string, zero, and false", () => {
    const blank = buildLocalRowModel(rows, columns, query({
      filter: { kind: "blank", columnId: "note", operator: "isBlank" }
    }), (row) => row.id);
    const zero = buildLocalRowModel(rows, columns, query({
      filter: { kind: "comparison", columnId: "salary", operator: "eq", value: { type: "number", value: 0 } }
    }), (row) => row.id);
    const inactive = buildLocalRowModel(rows, columns, query({
      filter: { kind: "comparison", columnId: "active", operator: "eq", value: { type: "boolean", value: false } }
    }), (row) => row.id);

    expect(blank.items.map((row) => row.id)).toEqual(["1", "2"]);
    expect(zero.items.map((row) => row.id)).toEqual(["1"]);
    expect(inactive.items.map((row) => row.id)).toEqual(["1", "3"]);
  });

  it("treats undefined accessor values as blanks in filters, aggregates, and groups", () => {
    type OptionalRow = { id: string; bonus?: number };
    const optionalHelper = createColumnHelper<OptionalRow>();
    const optionalColumns = [
      optionalHelper.accessor("bonus", { id: "bonus", header: "Bonus", dataType: "number" })
    ];
    const optionalRows: readonly OptionalRow[] = [
      { id: "missing-1" },
      { id: "paid", bonus: 10 },
      { id: "missing-2" }
    ];

    const blank = buildLocalRowModel(optionalRows, optionalColumns, query({
      filter: { kind: "blank", columnId: "bonus", operator: "isBlank" }
    }), (row) => row.id);
    const counted = buildLocalRowModel(optionalRows, optionalColumns, query({
      aggregates: [{ id: "bonus-count", columnId: "bonus", function: "count" }]
    }), (row) => row.id);
    const grouped = buildLocalRowModel(optionalRows, optionalColumns, query({
      grouping: [{ columnId: "bonus" }]
    }), (row) => row.id);

    expect(blank.items.map((row) => row.id)).toEqual(["missing-1", "missing-2"]);
    expect(counted.items.at(-1)).toMatchObject({
      kind: "aggregate",
      aggregates: { "bonus-count": 1 }
    });
    expect(grouped.items).toContainEqual(expect.objectContaining({
      kind: "group",
      key: { type: "null" },
      count: 2
    }));
  });

  it("sorts stably, groups, and aggregates the complete dataset", () => {
    const model = buildLocalRowModel(rows, columns, query({
      sorting: [{ columnId: "salary", direction: "desc", nulls: "last" }],
      grouping: [{ columnId: "department" }],
      aggregates: [{ id: "salary-sum", columnId: "salary", function: "sum" }],
      pagination: { kind: "none" }
    }), (row) => row.id);

    expect(model.totalDataRowCount).toBe(3);
    expect(model.pageInfo).toEqual({ kind: "none", total: { kind: "known", value: 3 } });
    expect(model.items[0]).toMatchObject({ kind: "group", count: 2, aggregates: { "salary-sum": 100 } });
  });

  it("paginates ungrouped data rows with record-count totals", () => {
    const model = buildLocalRowModel(rows, columns, query({
      pagination: { kind: "offset", offset: 1, limit: 1 }
    }), (row) => row.id);
    expect(model.items).toHaveLength(1);
    expect(model.pageInfo).toEqual({
      kind: "offset",
      offset: 1,
      limit: 1,
      total: { kind: "known", value: 3 },
      hasMore: true
    });
  });

  it("rejects grouping combined with offset pagination", () => {
    expect(() => buildLocalRowModel(rows, columns, query({
      grouping: [{ columnId: "department" }],
      pagination: { kind: "offset", offset: 0, limit: 2 }
    }), (row) => row.id)).toThrow("Grouping requires pagination kind none");
  });

  it("rejects blank and duplicate host row ids", () => {
    expect(() => buildLocalRowModel([{ ...rows[0], id: "" }], columns, query({}), (row) => row.id))
      .toThrow("Row id must not be blank");
    expect(() => buildLocalRowModel([rows[0], { ...rows[1], id: "1" }], columns, query({}), (row) => row.id))
      .toThrow("Duplicate row id: 1");
  });

  it("evaluates nested and/or/not filters", () => {
    const model = buildLocalRowModel(rows, columns, query({
      filter: {
        kind: "logical",
        operator: "and",
        operands: [
          { kind: "not", operand: { kind: "comparison", columnId: "salary", operator: "eq", value: { type: "number", value: 0 } } },
          {
            kind: "logical",
            operator: "or",
            operands: [
              { kind: "comparison", columnId: "note", operator: "contains", value: { type: "string", value: "ad" } },
              { kind: "comparison", columnId: "active", operator: "eq", value: { type: "boolean", value: true } }
            ]
          }
        ]
      }
    }), (row) => row.id);

    expect(model.items.map((item) => item.id)).toEqual(["2", "3"]);
  });

  it("does not let incompatible typed values satisfy negative filters", () => {
    const filters: QueryRequest["filter"][] = [
      { kind: "comparison", columnId: "salary", operator: "neq", value: { type: "string", value: "100" } },
      { kind: "set", columnId: "salary", operator: "notIn", values: [{ type: "string", value: "100" }] },
      {
        kind: "range",
        columnId: "salary",
        operator: "notBetween",
        lower: { type: "string", value: "0" },
        upper: { type: "string", value: "100" }
      }
    ];
    for (const filter of filters) {
      const model = buildLocalRowModel(rows, columns, query({ filter }), (row) => row.id);
      expect(model.items).toEqual([]);
    }
  });

  it("orders typed dates and datetimes without string coercion", () => {
    type Temporal = { id: string; day: string; instant: string };
    const temporalHelper = createColumnHelper<Temporal>();
    const temporalColumns = [
      temporalHelper.accessor("day", { id: "day", header: "Day", dataType: "date" }),
      temporalHelper.accessor("instant", { id: "instant", header: "Instant", dataType: "datetime" })
    ];
    const temporalRows: readonly Temporal[] = [
      { id: "late", day: "2026-02-01", instant: "2026-02-01T12:00:00Z" },
      { id: "early", day: "2026-01-01", instant: "2026-01-01T12:00:00Z" }
    ];
    const model = buildLocalRowModel(temporalRows, temporalColumns, query({
      sorting: [{ columnId: "instant", direction: "asc" }],
      filter: { kind: "comparison", columnId: "day", operator: "gte", value: { type: "date", value: "2026-01-01" } }
    }), (row) => row.id);
    expect(model.items.map((item) => item.id)).toEqual(["early", "late"]);
  });

  it("uses original source order as the stable sort tie-breaker", () => {
    const model = buildLocalRowModel(rows, columns, query({
      sorting: [{ columnId: "salary", direction: "desc" }]
    }), (row) => row.id);
    expect(model.items.map((item) => item.id)).toEqual(["2", "3", "1"]);
  });

  it("creates collision-proof nested group ids", () => {
    type Grouped = { id: string; region: string; team: string; value: number };
    const groupedHelper = createColumnHelper<Grouped>();
    const groupedRows: readonly Grouped[] = [
      { id: "a", region: "North", team: "Shared", value: 1 },
      { id: "b", region: "South", team: "Shared", value: 2 }
    ];
    const groupedColumns = [
      groupedHelper.accessor("region", { id: "region", header: "Region" }),
      groupedHelper.accessor("team", { id: "team", header: "Team" }),
      groupedHelper.accessor("value", { id: "value", header: "Value", dataType: "number" })
    ];
    const model = buildLocalRowModel(groupedRows, groupedColumns, query({
      grouping: [{ columnId: "region" }, { columnId: "team" }]
    }), (row) => row.id);
    const sharedGroups = model.items.filter((item) => item.kind === "group" && item.columnId === "team");
    expect(sharedGroups).toHaveLength(2);
    expect(new Set(sharedGroups.map((item) => item.id)).size).toBe(2);
  });

  it("expands and collapses trees while retaining all rows in the total", () => {
    type TreeRow = { id: string; label: string; children?: readonly TreeRow[] };
    const treeHelper = createColumnHelper<TreeRow>();
    const treeColumns = [treeHelper.accessor("label", { id: "label", header: "Label" })];
    const treeRows: readonly TreeRow[] = [{
      id: "root",
      label: "Root",
      children: [{ id: "child", label: "Child" }]
    }];
    const collapsed = buildLocalRowModel(treeRows, treeColumns, query({ tree: { expandedRowIds: [] } }), (row) => row.id, undefined, (row) => row.children);
    const expanded = buildLocalRowModel(treeRows, treeColumns, query({ tree: { expandedRowIds: ["root"] } }), (row) => row.id, undefined, (row) => row.children);
    expect(collapsed.items.map((item) => item.id)).toEqual(["root"]);
    expect(expanded.items).toMatchObject([
      { id: "root", depth: 0, hasChildren: true, expanded: true },
      { id: "child", depth: 1, parentId: "root" }
    ]);
    expect(collapsed.totalDataRowCount).toBe(2);
  });

  it("retains ancestors of matching tree descendants and sorts siblings", () => {
    type TreeRow = { id: string; label: string; children?: readonly TreeRow[] };
    const treeHelper = createColumnHelper<TreeRow>();
    const treeColumns = [treeHelper.accessor("label", { id: "label", header: "Label" })];
    const treeRows: readonly TreeRow[] = [{
      id: "root",
      label: "Parent",
      children: [{ id: "z", label: "Zulu match" }, { id: "a", label: "Alpha match" }]
    }];
    const model = buildLocalRowModel(treeRows, treeColumns, query({
      tree: { expandedRowIds: ["root"] },
      sorting: [{ columnId: "label", direction: "asc" }],
      filter: { kind: "comparison", columnId: "label", operator: "contains", value: { type: "string", value: "match" } }
    }), (row) => row.id, undefined, (row) => row.children);
    expect(model.items.map((item) => item.id)).toEqual(["root", "a", "z"]);
  });

  it("rejects duplicate and cyclic tree ids", () => {
    type TreeRow = { id: string; children?: readonly TreeRow[] };
    const treeHelper = createColumnHelper<TreeRow>();
    const treeColumns = [treeHelper.accessor("id", { id: "id", header: "Id" })];
    const shared: TreeRow = { id: "shared" };
    expect(() => buildLocalRowModel([
      { id: "first", children: [shared] },
      { id: "second", children: [shared] }
    ], treeColumns, query({ tree: { expandedRowIds: [] } }), (row) => row.id, undefined, (row) => row.children))
      .toThrow("Duplicate row id: shared");

    const cycle: { id: string; children: TreeRow[] } = { id: "cycle", children: [] };
    cycle.children.push(cycle);
    expect(() => buildLocalRowModel([cycle], treeColumns, query({ tree: { expandedRowIds: [] } }), (row) => row.id, undefined, (row) => row.children))
      .toThrow("Duplicate row id: cycle");
  });

  it("rejects tree grouping and pagination", () => {
    type TreeRow = { id: string; label: string; children?: readonly TreeRow[] };
    const treeHelper = createColumnHelper<TreeRow>();
    const treeColumns = [treeHelper.accessor("label", { id: "label", header: "Label" })];
    const treeRows: readonly TreeRow[] = [{ id: "root", label: "Root" }];
    expect(() => buildLocalRowModel(treeRows, treeColumns, query({
      tree: { expandedRowIds: [] }, grouping: [{ columnId: "label" }]
    }), (row) => row.id, undefined, (row) => row.children)).toThrow("Tree expansion cannot be combined with grouping");
    expect(() => buildLocalRowModel(treeRows, treeColumns, query({
      tree: { expandedRowIds: [] }, pagination: { kind: "offset", offset: 0, limit: 1 }
    }), (row) => row.id, undefined, (row) => row.children)).toThrow("Tree expansion requires pagination kind none");
  });

  it("computes all five aggregates over non-null typed values", () => {
    const model = buildLocalRowModel(rows, columns, query({
      grouping: [{ columnId: "department" }],
      aggregates: [
        { id: "sum", columnId: "salary", function: "sum" },
        { id: "average", columnId: "salary", function: "average" },
        { id: "count", columnId: "salary", function: "count" },
        { id: "min", columnId: "salary", function: "min" },
        { id: "max", columnId: "salary", function: "max" }
      ]
    }), (row) => row.id);
    expect(model.items[0]).toMatchObject({
      kind: "group",
      aggregates: { sum: 100, average: 50, count: 2, min: 0, max: 100 }
    });
  });

  it("rejects unknown aggregate columns and unsupported cursor pagination", () => {
    expect(() => buildLocalRowModel(rows, columns, query({
      aggregates: [{ id: "missing", columnId: "missing", function: "sum" }]
    }), (row) => row.id)).toThrow("Unknown column id: missing");
    expect(() => buildLocalRowModel(rows, columns, query({
      pagination: { kind: "cursor", limit: 10 }
    }), (row) => row.id)).toThrow("Local tables do not support cursor pagination");
  });

  it("uses one injected resolver for sort, filter, group, and aggregate", () => {
    const resolver: LocalValueResolver<Row> = ({ row, columnId }) => {
      if (columnId === "salary") {
        return { kind: "value", value: row.id === "1" ? 10 : 2 };
      }
      return { kind: "value", value: row[columnId as keyof Row] };
    };
    const model = buildLocalRowModel(rows, columns, query({
      sorting: [{ columnId: "salary", direction: "desc" }],
      filter: { kind: "comparison", columnId: "salary", operator: "gte", value: { type: "number", value: 2 } },
      grouping: [{ columnId: "salary" }],
      aggregates: [{ id: "sum", columnId: "salary", function: "sum" }]
    }), (row) => row.id, resolver);
    expect(model.items[0]).toMatchObject({ kind: "group", key: { type: "number", value: 10 }, aggregates: { sum: 10 } });
  });

  it("supports calculated dependencies and reports cycles", () => {
    const calculatedColumns = [
      ...columns,
      helper.computed<number>({
        id: "double",
        header: "Double",
        dataType: "number",
        calculate: ({ getValue }) => Number(getValue("salary")) * 2
      })
    ];
    const calculated = buildLocalRowModel(rows, calculatedColumns, query({
      sorting: [{ columnId: "double", direction: "desc" }]
    }), (row) => row.id);
    expect(calculated.items.map((item) => item.id)).toEqual(["2", "3", "1"]);

    const cycleColumns = [
      helper.computed<number>({ id: "a", header: "A", calculate: ({ getValue }) => Number(getValue("b")) }),
      helper.computed<number>({ id: "b", header: "B", calculate: ({ getValue }) => Number(getValue("a")) })
    ];
    const cycled = buildLocalRowModel(rows, cycleColumns, query({ sorting: [{ columnId: "a", direction: "asc" }] }), (row) => row.id);
    expect(cycled.issues).toContainEqual(expect.objectContaining({ code: "TABLE_VALUE_CYCLE", rowId: "1", columnId: "a" }));
  });

  it("sorts formula errors after values and de-duplicates addressed issues", () => {
    const resolver: LocalValueResolver<Row> = ({ row, columnId }) =>
      columnId === "salary" && row.id === "2"
        ? { kind: "error", code: "#FAIL!" }
        : { kind: "value", value: row[columnId as keyof Row] };
    const model = buildLocalRowModel(rows, columns, query({
      sorting: [{ columnId: "salary", direction: "asc" }],
      aggregates: [{ id: "sum", columnId: "salary", function: "sum" }]
    }), (row) => row.id, resolver);
    expect(model.items.filter((item) => item.kind === "data").map((item) => item.id)).toEqual(["1", "3", "2"]);
    expect(model.items.at(-1)).toMatchObject({ kind: "aggregate", aggregates: { sum: 100 } });
    expect(model.issues.filter((issue) => issue.rowId === "2" && issue.columnId === "salary")).toHaveLength(1);
    expect(model.issues).toContainEqual(expect.objectContaining({ code: "TABLE_VALUE_ERROR", rowId: "2", columnId: "salary" }));

    const grouped = buildLocalRowModel(rows, columns, query({
      grouping: [{ columnId: "salary" }]
    }), (row) => row.id, resolver);
    expect(grouped.items).toContainEqual(expect.objectContaining({
      kind: "group",
      key: { type: "error", value: "#FAIL!" }
    }));
  });

  it("does not mutate frozen input rows or arrays", () => {
    const frozenRows = Object.freeze(rows.map((row) => Object.freeze({ ...row })));
    const before = JSON.stringify(frozenRows);
    buildLocalRowModel(frozenRows, columns, query({ sorting: [{ columnId: "salary", direction: "desc" }] }), (row) => row.id);
    expect(JSON.stringify(frozenRows)).toBe(before);
  });
});
