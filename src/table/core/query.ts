export type QueryScalar =
  | { type: "string"; value: string }
  | { type: "number"; value: number }
  | { type: "boolean"; value: boolean }
  | { type: "date"; value: string }
  | { type: "datetime"; value: string }
  | { type: "error"; value: string }
  | { type: "null" };

export type FilterExpression =
  | {
      kind: "comparison";
      columnId: string;
      operator: "eq" | "neq" | "contains" | "startsWith" | "endsWith" | "gt" | "gte" | "lt" | "lte";
      value: QueryScalar;
    }
  | { kind: "set"; columnId: string; operator: "in" | "notIn"; values: readonly QueryScalar[] }
  | {
      kind: "range";
      columnId: string;
      operator: "between" | "notBetween";
      lower: QueryScalar;
      upper: QueryScalar;
    }
  | {
      kind: "blank";
      columnId: string;
      operator: "isNull" | "isNotNull" | "isEmpty" | "isNotEmpty" | "isBlank" | "isNotBlank";
    }
  | { kind: "logical"; operator: "and" | "or"; operands: readonly FilterExpression[] }
  | { kind: "not"; operand: FilterExpression };

export type TableSort = { columnId: string; direction: "asc" | "desc"; nulls?: "first" | "last" };
export type TableGrouping = { columnId: string; direction?: "asc" | "desc" };
export type TableAggregateRequest = {
  id: string;
  columnId: string;
  function: "sum" | "average" | "count" | "min" | "max";
};

export type PaginationRequest =
  | { kind: "none" }
  | { kind: "offset"; offset: number; limit: number }
  | { kind: "cursor"; cursor?: string; limit: number }
  | { kind: "infinite"; after?: string; limit: number };
export type TotalCount = { kind: "known"; value: number } | { kind: "unknown" };

export type QueryRequest = {
  sorting: readonly TableSort[];
  filter: FilterExpression | null;
  grouping: readonly TableGrouping[];
  aggregates: readonly TableAggregateRequest[];
  pagination: PaginationRequest;
  tree?: { expandedRowIds: readonly string[] };
};

export type QueryRow<TRow> =
  | {
      kind: "data";
      id: string;
      original: TRow;
      depth: number;
      parentId?: string;
      hasChildren?: boolean;
      expanded?: boolean;
    }
  | {
      kind: "group";
      id: string;
      depth: number;
      columnId: string;
      key: QueryScalar;
      count: number;
      aggregates: Readonly<Record<string, unknown>>;
    }
  | { kind: "aggregate"; id: string; depth: number; aggregates: Readonly<Record<string, unknown>> };

export type QueryResult<TRow> = {
  items: readonly QueryRow<TRow>[];
  revision: string;
  completeness: "loadedRows" | "completeDataset";
  pageInfo:
    | { kind: "none"; total: Extract<TotalCount, { kind: "known" }> }
    | { kind: "offset"; offset: number; limit: number; total: TotalCount; hasMore: boolean }
    | { kind: "cursor"; nextCursor?: string; previousCursor?: string; total: TotalCount }
    | { kind: "infinite"; nextCursor?: string; loadedCount: number; total: TotalCount };
};

const COMPARISON_OPERATORS = new Set(["eq", "neq", "contains", "startsWith", "endsWith", "gt", "gte", "lt", "lte"]);
const BLANK_OPERATORS = new Set(["isNull", "isNotNull", "isEmpty", "isNotEmpty", "isBlank", "isNotBlank"]);

export function serializeQueryRequest(request: QueryRequest): string {
  return JSON.stringify(normalizeQueryRequest(request));
}

export function deserializeQueryRequest(serialized: string): QueryRequest {
  let value: unknown;
  try {
    value = JSON.parse(serialized);
  } catch {
    invalid("expected valid JSON");
  }
  return normalizeQueryRequest(value);
}

function normalizeQueryRequest(value: unknown): QueryRequest {
  const source = record(value, "expected an object");
  const sorting = array(source.sorting, "sorting must be an array").map((item, index) => {
    const sort = record(item, `sorting[${index}] must be an object`);
    const columnId = identifier(sort.columnId, `sorting[${index}].columnId`);
    const direction = oneOf(sort.direction, ["asc", "desc"] as const, `sorting[${index}].direction`);
    const nulls = sort.nulls === undefined
      ? undefined
      : oneOf(sort.nulls, ["first", "last"] as const, `sorting[${index}].nulls`);
    return nulls === undefined ? { columnId, direction } : { columnId, direction, nulls };
  });
  const filter = source.filter === null ? null : normalizeFilter(source.filter, "filter", 0);
  const grouping = array(source.grouping, "grouping must be an array").map((item, index) => {
    const group = record(item, `grouping[${index}] must be an object`);
    const columnId = identifier(group.columnId, `grouping[${index}].columnId`);
    const direction = group.direction === undefined
      ? undefined
      : oneOf(group.direction, ["asc", "desc"] as const, `grouping[${index}].direction`);
    return direction === undefined ? { columnId } : { columnId, direction };
  });
  const aggregates = array(source.aggregates, "aggregates must be an array").map((item, index) => {
    const aggregate = record(item, `aggregates[${index}] must be an object`);
    return {
      id: identifier(aggregate.id, `aggregates[${index}].id`),
      columnId: identifier(aggregate.columnId, `aggregates[${index}].columnId`),
      function: oneOf(
        aggregate.function,
        ["sum", "average", "count", "min", "max"] as const,
        `aggregates[${index}].function`
      )
    };
  });
  const pagination = normalizePagination(source.pagination);
  if (grouping.length > 0 && pagination.kind !== "none") {
    invalid("Grouping requires pagination kind none");
  }

  if (source.tree === undefined) {
    return { sorting, filter, grouping, aggregates, pagination };
  }
  const tree = record(source.tree, "tree must be an object");
  const expandedRowIds = array(tree.expandedRowIds, "tree.expandedRowIds must be an array")
    .map((item, index) => identifier(item, `tree.expandedRowIds[${index}]`));
  if (new Set(expandedRowIds).size !== expandedRowIds.length) {
    invalid("tree.expandedRowIds must be unique");
  }
  if (grouping.length > 0) {
    invalid("Tree expansion cannot be combined with grouping");
  }
  if (pagination.kind !== "none") {
    invalid("Tree expansion requires pagination kind none");
  }
  return { sorting, filter, grouping, aggregates, pagination, tree: { expandedRowIds } };
}

function normalizeFilter(value: unknown, path: string, depth: number): FilterExpression {
  if (depth > 100) invalid(`${path} is nested too deeply`);
  const source = record(value, `${path} must be an object`);
  switch (source.kind) {
    case "comparison":
      if (typeof source.operator !== "string" || !COMPARISON_OPERATORS.has(source.operator)) {
        invalid(`${path}.operator is invalid`);
      }
      return {
        kind: "comparison",
        columnId: identifier(source.columnId, `${path}.columnId`),
        operator: source.operator as Extract<FilterExpression, { kind: "comparison" }>["operator"],
        value: normalizeScalar(source.value, `${path}.value`)
      };
    case "set": {
      const operator = oneOf(source.operator, ["in", "notIn"] as const, `${path}.operator`);
      const values = array(source.values, `${path}.values must be an array`)
        .map((item, index) => normalizeScalar(item, `${path}.values[${index}]`));
      return { kind: "set", columnId: identifier(source.columnId, `${path}.columnId`), operator, values };
    }
    case "range":
      return {
        kind: "range",
        columnId: identifier(source.columnId, `${path}.columnId`),
        operator: oneOf(source.operator, ["between", "notBetween"] as const, `${path}.operator`),
        lower: normalizeScalar(source.lower, `${path}.lower`),
        upper: normalizeScalar(source.upper, `${path}.upper`)
      };
    case "blank":
      if (typeof source.operator !== "string" || !BLANK_OPERATORS.has(source.operator)) {
        invalid(`${path}.operator is invalid`);
      }
      return {
        kind: "blank",
        columnId: identifier(source.columnId, `${path}.columnId`),
        operator: source.operator as Extract<FilterExpression, { kind: "blank" }>["operator"]
      };
    case "logical": {
      const operands = array(source.operands, `${path}.operands must be an array`);
      if (operands.length === 0) invalid(`${path}.operands must not be empty`);
      return {
        kind: "logical",
        operator: oneOf(source.operator, ["and", "or"] as const, `${path}.operator`),
        operands: operands.map((item, index) => normalizeFilter(item, `${path}.operands[${index}]`, depth + 1))
      };
    }
    case "not":
      return { kind: "not", operand: normalizeFilter(source.operand, `${path}.operand`, depth + 1) };
    default:
      return invalid(`${path}.kind is invalid`);
  }
}

function normalizeScalar(value: unknown, path: string): QueryScalar {
  const source = record(value, `${path} must be an object`);
  switch (source.type) {
    case "string":
    case "error":
      if (typeof source.value !== "string") invalid(`${path}.value must be a string`);
      return { type: source.type, value: source.value };
    case "number":
      if (typeof source.value !== "number" || !Number.isFinite(source.value)) {
        invalid(`${path}.value must be a finite number`);
      }
      return { type: "number", value: source.value };
    case "boolean":
      if (typeof source.value !== "boolean") invalid(`${path}.value must be a boolean`);
      return { type: "boolean", value: source.value };
    case "date":
      if (typeof source.value !== "string" || !isIsoDate(source.value)) {
        invalid(`${path}.value must be an ISO date`);
      }
      return { type: "date", value: source.value };
    case "datetime":
      if (typeof source.value !== "string" || !isIsoDateTime(source.value)) {
        invalid(`${path}.value must be an ISO datetime`);
      }
      return { type: "datetime", value: source.value };
    case "null":
      return { type: "null" };
    default:
      return invalid(`${path}.type is invalid`);
  }
}

function normalizePagination(value: unknown): PaginationRequest {
  const source = record(value, "pagination must be an object");
  switch (source.kind) {
    case "none":
      return { kind: "none" };
    case "offset":
      return {
        kind: "offset",
        offset: nonNegativeInteger(source.offset, "pagination.offset"),
        limit: positiveInteger(source.limit, "pagination.limit")
      };
    case "cursor": {
      const cursor = optionalString(source.cursor, "pagination.cursor");
      const limit = positiveInteger(source.limit, "pagination.limit");
      return cursor === undefined ? { kind: "cursor", limit } : { kind: "cursor", cursor, limit };
    }
    case "infinite": {
      const after = optionalString(source.after, "pagination.after");
      const limit = positiveInteger(source.limit, "pagination.limit");
      return after === undefined ? { kind: "infinite", limit } : { kind: "infinite", after, limit };
    }
    default:
      return invalid("pagination.kind is invalid");
  }
}

function record(value: unknown, reason: string): Record<string, unknown> {
  if (typeof value !== "object" || value === null || Array.isArray(value)) invalid(reason);
  return value as Record<string, unknown>;
}

function array(value: unknown, reason: string): unknown[] {
  if (!Array.isArray(value)) invalid(reason);
  return value;
}

function identifier(value: unknown, path: string): string {
  if (typeof value !== "string" || value.trim().length === 0) invalid(`${path} must be a non-blank string`);
  return value;
}

function oneOf<const T extends readonly string[]>(value: unknown, choices: T, path: string): T[number] {
  if (typeof value !== "string" || !choices.includes(value)) invalid(`${path} is invalid`);
  return value as T[number];
}

function nonNegativeInteger(value: unknown, path: string): number {
  if (typeof value !== "number" || !Number.isSafeInteger(value) || value < 0) {
    invalid(`${path} must be a non-negative integer`);
  }
  return value;
}

function positiveInteger(value: unknown, path: string): number {
  if (typeof value !== "number" || !Number.isSafeInteger(value) || value <= 0) {
    invalid(`${path} must be a positive integer`);
  }
  return value;
}

function optionalString(value: unknown, path: string): string | undefined {
  if (value === undefined) return undefined;
  if (typeof value !== "string" || value.trim().length === 0) invalid(`${path} must be a non-blank string`);
  return value;
}

function isIsoDate(value: string): boolean {
  const match = /^(\d{4})-(\d{2})-(\d{2})$/.exec(value);
  if (!match) return false;
  const date = new Date(`${value}T00:00:00.000Z`);
  return !Number.isNaN(date.valueOf()) && date.toISOString().slice(0, 10) === value;
}

function isIsoDateTime(value: string): boolean {
  return /^\d{4}-\d{2}-\d{2}T/.test(value) && !Number.isNaN(Date.parse(value));
}

function invalid(reason: string): never {
  throw new Error(`Invalid table query: ${reason}`);
}
