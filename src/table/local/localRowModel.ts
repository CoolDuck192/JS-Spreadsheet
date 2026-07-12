import { normalizeColumns } from "../core/columnHelper";
import type { QueryRequest, QueryResult, QueryRow, QueryScalar, TableGrouping } from "../core/query";
import { safeInvokeTableExtension } from "../core/safeInvoke";
import type { ColumnDef, TableCellIssue } from "../core/types";
import { calculateLocalAggregates } from "./localAggregates";
import { matchesFilter } from "./localFilter";

export type LocalRowModel<TRow> = {
  items: readonly QueryRow<TRow>[];
  totalDataRowCount: number;
  pageInfo: QueryResult<TRow>["pageInfo"];
  dataRowsById: ReadonlyMap<string, TRow>;
  orderedDataRowIds: readonly string[];
  issues: readonly TableCellIssue[];
};

export type LocalEvaluatedValue =
  | { kind: "value"; value: unknown }
  | { kind: "error"; code: string };
export type LocalValueResolver<TRow> = (context: {
  row: TRow;
  rowId: string;
  columnId: string;
  getValue(columnId: string): LocalEvaluatedValue;
}) => LocalEvaluatedValue;

type AnyColumn<TRow> = ColumnDef<TRow, any>;
type DataNode<TRow> = {
  row: TRow;
  id: string;
  depth: number;
  parentId?: string;
  sourceIndex: number;
  children: readonly DataNode<TRow>[];
};

export function buildLocalRowModel<TRow>(
  rows: readonly TRow[],
  columns: readonly AnyColumn<TRow>[],
  request: QueryRequest,
  getRowId: (row: TRow) => string,
  resolveValue?: LocalValueResolver<TRow>,
  getSubRows?: (row: TRow) => readonly TRow[] | undefined
): LocalRowModel<TRow> {
  if (request.grouping.length > 0 && request.pagination.kind !== "none") {
    throw new Error("Grouping requires pagination kind none");
  }
  if (getSubRows && request.grouping.length > 0) {
    throw new Error("Tree expansion cannot be combined with grouping");
  }
  if (getSubRows && request.pagination.kind !== "none") {
    throw new Error("Tree expansion requires pagination kind none");
  }
  if (request.pagination.kind === "cursor") {
    throw new Error("Local tables do not support cursor pagination");
  }
  if (request.pagination.kind === "infinite") {
    throw new Error("Local tables do not support infinite pagination");
  }

  const normalizedColumns = normalizeColumns(columns) as readonly AnyColumn<TRow>[];
  const columnsById = new Map(normalizedColumns.map((column) => [column.id, column]));
  validateRequestColumns(request, columnsById);

  const dataRowsById = new Map<string, TRow>();
  const roots = normalizeRows(rows, undefined, 0);
  const issues = new Map<string, TableCellIssue>();
  const evaluationCache = new Map<string, LocalEvaluatedValue>();
  const evaluationStack = new Set<string>();

  function normalizeRows(source: readonly TRow[], parentId: string | undefined, depth: number): DataNode<TRow>[] {
    return source.map((row, sourceIndex) => {
      let rawId: string;
      try {
        rawId = getRowId(row);
      } catch {
        throw new Error("Host getRowId extension failed");
      }
      const id = typeof rawId === "string" ? rawId.trim() : "";
      if (!id) throw new Error("Row id must not be blank");
      if (dataRowsById.has(id)) throw new Error(`Duplicate row id: ${id}`);
      dataRowsById.set(id, row);

      let childRows: readonly TRow[] = [];
      if (getSubRows) {
        try {
          childRows = getSubRows(row) ?? [];
        } catch {
          throw new Error("Host getSubRows extension failed");
        }
        if (!Array.isArray(childRows)) throw new Error("Host getSubRows extension failed");
      }
      const children = normalizeRows(childRows, id, depth + 1);
      return { row, id, depth, ...(parentId ? { parentId } : {}), sourceIndex, children };
    });
  }

  function resolve(row: TRow, rowId: string, columnId: string): LocalEvaluatedValue {
    const key = JSON.stringify([rowId, columnId]);
    const cached = evaluationCache.get(key);
    if (cached) return cached;
    const column = columnsById.get(columnId);
    if (!column) throw new Error(`Unknown column id: ${columnId}`);
    if (evaluationStack.has(key)) {
      const cycle = { kind: "error", code: "#CYCLE!" } as const;
      recordValueIssue(rowId, columnId, cycle.code);
      return cycle;
    }

    evaluationStack.add(key);
    let evaluated: LocalEvaluatedValue;
    if (resolveValue) {
      const result = safeInvokeTableExtension("formula", () => resolveValue({
        row,
        rowId,
        columnId,
        getValue: (dependency) => resolve(row, rowId, dependency)
      }));
      evaluated = result.ok ? result.value : { kind: "error", code: "#ERROR!" };
      if (!result.ok) recordIssue({ ...result.issue, rowId, columnId });
    } else {
      evaluated = resolveColumn(column, row, rowId, columnId);
    }
    evaluationStack.delete(key);
    evaluationCache.set(key, evaluated);
    if (evaluated.kind === "error") recordValueIssue(rowId, columnId, evaluated.code);
    return evaluated;
  }

  function resolveColumn(column: AnyColumn<TRow>, row: TRow, rowId: string, columnId: string): LocalEvaluatedValue {
    if (column.kind === "computed") {
      let dependencyError: LocalEvaluatedValue | null = null;
      const result = safeInvokeTableExtension("calculate", () => column.calculate({
        row,
        rowId,
        columnId,
        getValue(dependency) {
          const value = resolve(row, rowId, dependency);
          if (value.kind === "error") dependencyError = value;
          return value.kind === "value" ? value.value : value;
        }
      }));
      if (!result.ok) {
        recordIssue({ ...result.issue, rowId, columnId });
        return { kind: "error", code: "#ERROR!" };
      }
      return dependencyError ?? { kind: "value", value: result.value };
    }
    if ("accessor" in column && typeof column.accessor === "function") {
      const result = safeInvokeTableExtension("accessor", () => column.accessor(row));
      if (!result.ok) {
        recordIssue({ ...result.issue, rowId, columnId });
        return { kind: "error", code: "#ERROR!" };
      }
      return { kind: "value", value: result.value };
    }
    return { kind: "value", value: null };
  }

  function recordValueIssue(rowId: string, columnId: string, code: string): void {
    recordIssue({
      code: code === "#CYCLE!" ? "TABLE_VALUE_CYCLE" : "TABLE_VALUE_ERROR",
      message: code === "#CYCLE!" ? "Calculated column cycle" : "Cell evaluation failed",
      rowId,
      columnId
    });
  }

  function recordIssue(issue: TableCellIssue): void {
    const key = JSON.stringify([issue.code, issue.rowId ?? "", issue.columnId ?? ""]);
    if (!issues.has(key)) issues.set(key, issue);
  }

  const getValue = (row: TRow, rowId: string, columnId: string): unknown => {
    const value = resolve(row, rowId, columnId);
    return value.kind === "value" ? value.value : value;
  };
  const filteredRoots = filterNodes(roots);
  const sortedRoots = sortNodes(filteredRoots);
  const orderedNodes = getSubRows ? flattenAll(sortedRoots) : sortedRoots;
  const totalDataRowCount = orderedNodes.length;

  let items: QueryRow<TRow>[];
  let pageInfo: QueryResult<TRow>["pageInfo"];
  if (request.grouping.length > 0) {
    items = buildGroups(sortedRoots, 0, []);
    pageInfo = { kind: "none", total: { kind: "known", value: totalDataRowCount } };
  } else if (getSubRows) {
    items = flattenVisible(sortedRoots, new Set(request.tree?.expandedRowIds ?? []));
    if (request.aggregates.length > 0) items.push(aggregateRow(orderedNodes));
    pageInfo = { kind: "none", total: { kind: "known", value: totalDataRowCount } };
  } else if (request.pagination.kind === "offset") {
    const { offset, limit } = request.pagination;
    const page = sortedRoots.slice(offset, offset + limit);
    items = page.map((node) => dataQueryRow(node, 0));
    if (request.aggregates.length > 0) items.push(aggregateRow(sortedRoots));
    pageInfo = {
      kind: "offset",
      offset,
      limit,
      total: { kind: "known", value: totalDataRowCount },
      hasMore: offset + limit < totalDataRowCount
    };
  } else {
    items = sortedRoots.map((node) => dataQueryRow(node, 0));
    if (request.aggregates.length > 0) items.push(aggregateRow(sortedRoots));
    pageInfo = { kind: "none", total: { kind: "known", value: totalDataRowCount } };
  }

  return {
    items,
    totalDataRowCount,
    pageInfo,
    dataRowsById,
    orderedDataRowIds: orderedNodes.map((node) => node.id),
    issues: [...issues.values()]
  };

  function filterNodes(nodes: readonly DataNode<TRow>[]): DataNode<TRow>[] {
    return nodes.flatMap((node) => {
      const children = filterNodes(node.children);
      const selfMatches = request.filter === null
        || matchesFilter(node.row, node.id, request.filter, columnsById, getValue);
      return selfMatches || children.length > 0 ? [{ ...node, children }] : [];
    });
  }

  function sortNodes(nodes: readonly DataNode<TRow>[]): DataNode<TRow>[] {
    return nodes
      .map((node) => ({ ...node, children: sortNodes(node.children) }))
      .sort(compareNodes);
  }

  function compareNodes(left: DataNode<TRow>, right: DataNode<TRow>): number {
    for (const sort of request.sorting) {
      const column = columnsById.get(sort.columnId)!;
      const leftValue = getValue(left.row, left.id, sort.columnId);
      const rightValue = getValue(right.row, right.id, sort.columnId);
      const category = compareCategories(leftValue, rightValue, sort.nulls);
      if (category !== 0) return category;
      if (isOrdinary(leftValue) && isOrdinary(rightValue)) {
        let comparison: number;
        if (column.compare) {
          const result = safeInvokeTableExtension("calculate", () => column.compare!(leftValue, rightValue));
          if (!result.ok) {
            recordIssue({ ...result.issue, columnId: sort.columnId });
            comparison = 0;
          } else {
            comparison = result.value;
          }
        } else {
          comparison = compareOrdinary(leftValue, rightValue);
        }
        if (comparison !== 0) return sort.direction === "desc" ? -comparison : comparison;
      }
    }
    return left.sourceIndex - right.sourceIndex;
  }

  function buildGroups(
    nodes: readonly DataNode<TRow>[],
    level: number,
    ancestors: readonly [string, QueryScalar][]
  ): QueryRow<TRow>[] {
    const grouping = request.grouping[level];
    const grouped = new Map<string, { scalar: QueryScalar; nodes: DataNode<TRow>[] }>();
    for (const node of nodes) {
      const scalar = toQueryScalar(getValue(node.row, node.id, grouping.columnId), columnsById.get(grouping.columnId)!);
      const key = JSON.stringify(scalar);
      const bucket = grouped.get(key) ?? { scalar, nodes: [] };
      bucket.nodes.push(node);
      grouped.set(key, bucket);
    }
    let groups = [...grouped.values()];
    if (grouping.direction) {
      groups = groups.sort((left, right) => {
        const comparison = compareScalars(left.scalar, right.scalar);
        return grouping.direction === "desc" ? -comparison : comparison;
      });
    }
    return groups.flatMap((group) => {
      const path = [...ancestors, [grouping.columnId, group.scalar] as [string, QueryScalar]];
      const groupRow: QueryRow<TRow> = {
        kind: "group",
        id: JSON.stringify(["group", ...path]),
        depth: level,
        columnId: grouping.columnId,
        key: group.scalar,
        count: group.nodes.length,
        aggregates: aggregates(group.nodes)
      };
      const descendants = level + 1 < request.grouping.length
        ? buildGroups(group.nodes, level + 1, path)
        : group.nodes.map((node) => dataQueryRow(node, request.grouping.length));
      return [groupRow, ...descendants];
    });
  }

  function aggregates(nodes: readonly DataNode<TRow>[]): Readonly<Record<string, unknown>> {
    return calculateLocalAggregates(
      nodes.map((node) => ({ row: node.row, rowId: node.id })),
      request.aggregates,
      columnsById,
      getValue
    );
  }

  function aggregateRow(nodes: readonly DataNode<TRow>[]): QueryRow<TRow> {
    return { kind: "aggregate", id: "aggregate:local", depth: 0, aggregates: aggregates(nodes) };
  }
}

function validateRequestColumns<TRow>(
  request: QueryRequest,
  columns: ReadonlyMap<string, AnyColumn<TRow>>
): void {
  const ids = new Set<string>();
  request.sorting.forEach((item) => ids.add(item.columnId));
  request.grouping.forEach((item) => ids.add(item.columnId));
  request.aggregates.forEach((item) => ids.add(item.columnId));
  collectFilterColumns(request.filter, ids);
  for (const id of ids) {
    if (!columns.has(id)) throw new Error(`Unknown column id: ${id}`);
  }
}

function collectFilterColumns(filter: QueryRequest["filter"], ids: Set<string>): void {
  if (!filter) return;
  if (filter.kind === "logical") {
    filter.operands.forEach((operand) => collectFilterColumns(operand, ids));
  } else if (filter.kind === "not") {
    collectFilterColumns(filter.operand, ids);
  } else {
    ids.add(filter.columnId);
  }
}

function flattenAll<TRow>(nodes: readonly DataNode<TRow>[]): DataNode<TRow>[] {
  return nodes.flatMap((node) => [node, ...flattenAll(node.children)]);
}

function flattenVisible<TRow>(
  nodes: readonly DataNode<TRow>[],
  expanded: ReadonlySet<string>
): QueryRow<TRow>[] {
  return nodes.flatMap((node) => {
    const isExpanded = expanded.has(node.id);
    const row = dataQueryRow(node, node.depth, isExpanded);
    return isExpanded ? [row, ...flattenVisible(node.children, expanded)] : [row];
  });
}

function dataQueryRow<TRow>(node: DataNode<TRow>, depth: number, expanded?: boolean): QueryRow<TRow> {
  return {
    kind: "data",
    id: node.id,
    original: node.row,
    depth,
    ...(node.parentId ? { parentId: node.parentId } : {}),
    ...(node.children.length > 0 ? { hasChildren: true, expanded: expanded ?? false } : {})
  };
}

function compareCategories(left: unknown, right: unknown, nulls?: "first" | "last"): number {
  return valueCategory(left, nulls) - valueCategory(right, nulls);
}

function valueCategory(value: unknown, nulls?: "first" | "last"): number {
  if (value === null || value === undefined) return nulls === "first" ? 0 : 3;
  if (value === "") return 3;
  if (isEvaluationError(value)) return 2;
  return 1;
}

function isOrdinary(value: unknown): value is string | number | boolean {
  return typeof value === "string" || typeof value === "number" || typeof value === "boolean";
}

function compareOrdinary(left: string | number | boolean, right: string | number | boolean): number {
  if (typeof left !== typeof right) return typeof left < typeof right ? -1 : 1;
  if (left === right) return 0;
  return left < right ? -1 : 1;
}

function isEvaluationError(value: unknown): value is { kind: "error"; code: string } {
  return typeof value === "object" && value !== null && "kind" in value && value.kind === "error" && "code" in value;
}

function toQueryScalar<TRow>(value: unknown, column: AnyColumn<TRow>): QueryScalar {
  if (isEvaluationError(value)) return { type: "error", value: value.code };
  if (value === null || value === undefined) return { type: "null" };
  if (typeof value === "number") return { type: "number", value };
  if (typeof value === "boolean") return { type: "boolean", value };
  if (column.dataType === "date") return { type: "date", value: String(value) };
  if (column.dataType === "datetime") return { type: "datetime", value: String(value) };
  return { type: "string", value: String(value) };
}

function compareScalars(left: QueryScalar, right: QueryScalar): number {
  if (left.type !== right.type) return left.type < right.type ? -1 : 1;
  if (left.type === "null" || right.type === "null") return 0;
  if (left.value === right.value) return 0;
  return left.value < right.value ? -1 : 1;
}
