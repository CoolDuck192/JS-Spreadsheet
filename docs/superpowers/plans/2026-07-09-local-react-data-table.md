# Local React DataTable Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add a typed, embeddable React `DataTable` for complete local record sets, backed by a source-specific `RecordTableSession` and the same two-axis viewport and interaction kernel as the spreadsheet.

**Architecture:** DOM-free table contracts describe queries, capabilities, stable row/column identities, lazy cell snapshots, command lifecycle, and source-neutral operation state. `RecordTableSession` owns local record projection, atomic edits, metadata, and bounded inverse history; React owns virtualization, focus, selection, editor placement, clipboard integration, accessible DOM, and the adaptive table shell. The existing spreadsheet remains authoritative over workbook cells and becomes the first adapter over the shared React viewport kernel.

**Tech Stack:** Node 22.13+, pnpm 11.7, TypeScript 6, React 19, Vitest 4, Testing Library, fast-check, axe-core, Playwright, the foundation plan's `CommandResult`, `parseCellInput`, and two-axis measurement primitives.

## Global Constraints

- Complete `2026-07-09-data-table-foundation.md` first; this plan consumes `src/core/commands/types.ts`, `src/core/values/parseCellInput.ts`, and `src/core/viewport/axis.ts`.
- Keep `src/table/core` and `src/table/local` free of React, DOM, storage, authentication, CSS, and browser globals.
- Application records remain authoritative. Stable row IDs always come from the host's `getRowId`; visible indexes are never durable identities.
- `TableViewSnapshot.getCell(rowId, columnId)` is the canonical lazy cell accessor. Do not materialize a record-per-cell matrix or add a competing `getCellValue` API.
- Keep local, remote, and workbook mutation/history implementations separate. Only the query, capability, snapshot, intent, and command-lifecycle contracts are shared.
- Every multi-cell or row-structure command validates completely and commits atomically, or leaves rows, metadata, history, and subscribers unchanged.
- Local sorting, filtering, grouping, aggregation, formulas, and pagination may claim `completeDataset` only because all local records are present.
- Use the approved white, slate, and muted-green palette. Scope every new selector below `.js-spreadsheet-root`; never add global `body`, `button`, `input`, `select`, or unscoped universal selectors.
- Preserve the no-prop spreadsheet, existing `Grid` props, workbook semantics, and all existing spreadsheet tests while extracting the shared viewport.
- Use test-first red-green-refactor for every production change. A test must fail for the intended missing behavior before implementation is written.
- Do not add remote fetching/mutation, structured workbook tables, native table XLSX metadata, package export maps, Laravel, Web Components, or non-React framework bindings in this plan.
- Do not publish or open a pull request after this plan alone; the program-level packaging and pressure plan owns the release gate and PR.

---

### Task 1: Define shared query and capability contracts

**Files:**
- Create: `src/table/core/query.ts`
- Create: `src/table/core/query.test.ts`
- Create: `src/table/core/capabilities.ts`
- Create: `src/table/core/capabilities.test.ts`
- Create: `src/table/core/index.ts`

**Interfaces:**
- Produces: `QueryScalar`, `FilterExpression`, `TableSort`, `TableGrouping`, `TableAggregateRequest`, `PaginationRequest`, `TotalCount`, `QueryRequest`, renderer-neutral `QueryRow`, `QueryResult`, `OperationCapability`, `FormulaCapability`, `TableCapabilities`, `TableFeature`, `TableOperationState`, `serializeQueryRequest`, `deserializeQueryRequest`, `createLocalTableCapabilities`, and `resolveTableOperationStates`.
- Used later by: local row projection, `RecordTableSession`, `DataTable`, `WorkbookTableSession`, and `RemoteTableSession`.

- [ ] **Step 1: Write failing serializable-query tests**

Create `src/table/core/query.test.ts` with concrete values and no test-only production hooks:

```ts
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
      pagination: { kind: "offset", offset: 50, limit: 25 }
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
```

- [ ] **Step 2: Run the query test and verify the intended failure**

Run:

```bash
corepack pnpm exec vitest run src/table/core/query.test.ts
```

Expected: FAIL because `src/table/core/query.ts` does not exist. A syntax or environment error is not an acceptable red state.

- [ ] **Step 3: Implement the exact query AST and runtime decoder**

Create `src/table/core/query.ts` with these public types:

```ts
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
  | { kind: "data"; id: string; original: TRow; depth: number; parentId?: string; hasChildren?: boolean; expanded?: boolean }
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
    | {
        kind: "offset";
        offset: number;
        limit: number;
        total: TotalCount;
        hasMore: boolean;
      }
    | {
        kind: "cursor";
        nextCursor?: string;
        previousCursor?: string;
        total: TotalCount;
      }
    | {
        kind: "infinite";
        nextCursor?: string;
        loadedCount: number;
        total: TotalCount;
      };
};
```

`QueryResult.items` may contain data, server/local group, and aggregate projections. Every `pageInfo.total` counts underlying data records, not the number of rendered group/aggregate items. `completeness: "loadedRows"` means the result is only a projection of loaded records; `"completeDataset"` is an explicit source guarantee.

Implement a private `function normalizeQueryRequest(value: unknown): QueryRequest` that validates and rebuilds every object with the field order shown above while preserving user-significant sort/group/filter/tree operand order. `serializeQueryRequest` returns `JSON.stringify(normalizeQueryRequest(request))`; `deserializeQueryRequest` parses JSON and passes the result through the same function. Require finite numbers, non-blank/unique expanded IDs, positive pagination limits, non-negative offsets, and valid ISO date/datetime strings, and throw `new Error("Invalid table query: <reason>")` on the first invalid node. Keep dates as tagged ISO strings so the AST remains JSON-safe. Reject non-`none` pagination when grouping is non-empty. Also reject tree expansion combined with grouping or non-`none` pagination: the first-release AST does not pretend data-row offsets, flattened group headers, and hierarchical order are equivalent. Grouped/tree results remain virtualized and may still be filtered/sorted within their declared semantics.

- [ ] **Step 4: Run query tests and verify green**

Run: `corepack pnpm exec vitest run src/table/core/query.test.ts`

Expected: PASS with all four tests.

- [ ] **Step 5: Write failing capability-resolution tests**

Create `src/table/core/capabilities.test.ts`:

```ts
import { describe, expect, it } from "vitest";
import { createLocalTableCapabilities, resolveTableOperationStates } from "./capabilities";

describe("table capability resolution", () => {
  it("advertises local operations as client-owned complete-dataset work", () => {
    const states = resolveTableOperationStates(createLocalTableCapabilities(), {});
    expect(states.sort).toEqual({ enabled: true, scopeLabel: "Complete dataset" });
    expect(states.filter).toEqual({ enabled: true, scopeLabel: "Complete dataset" });
    expect(states.formula).toEqual({ enabled: false, reason: "Formula service is not configured" });
  });

  it("intersects source support with host feature configuration", () => {
    const states = resolveTableOperationStates(createLocalTableCapabilities({ formula: "fullLocalDataset" }), {
      sort: false,
      formula: true
    });
    expect(states.sort).toEqual({ enabled: false, reason: "Disabled by host configuration" });
    expect(states.formula).toEqual({ enabled: true, scopeLabel: "Complete dataset" });
  });

  it("rejects a complete-dataset requirement for loaded-row support", () => {
    const capabilities = createLocalTableCapabilities();
    capabilities.sort = { executor: "client", scope: "loadedRows" };
    const states = resolveTableOperationStates(capabilities, {
      sort: { requiredScope: "completeDataset" }
    });
    expect(states.sort).toEqual({
      enabled: false,
      scopeLabel: "Loaded rows",
      reason: "This table only supports sorting loaded rows."
    });
  });
});
```

- [ ] **Step 6: Run capability tests and verify red**

Run: `corepack pnpm exec vitest run src/table/core/capabilities.test.ts`

Expected: FAIL because `capabilities.ts` is missing.

- [ ] **Step 7: Implement source-neutral capabilities and operation states**

Create `src/table/core/capabilities.ts`:

```ts
import type { PaginationRequest } from "./query";

export type OperationCapability = {
  executor: "client" | "server";
  scope: "loadedRows" | "completeDataset";
};

export type FormulaCapability = "none" | "loadedRows" | "fullLocalDataset" | "server";
export type TableFeature =
  | "sort" | "filter" | "group" | "aggregate" | "pagination"
  | "edit" | "bulkEdit" | "metadata" | "validation" | "formula"
  | "subscription" | "undo" | "export";
export type TableFeatureConfiguration = Partial<Record<
  TableFeature,
  | boolean
  | { enabled?: boolean; requiredScope?: "loadedRows" | "completeDataset" }
>>;
export type TableOperationState = { enabled: boolean; scopeLabel?: "Loaded rows" | "Complete dataset"; reason?: string };

export type TableCapabilities = {
  sort: OperationCapability | false;
  filter: OperationCapability | false;
  group: OperationCapability | false;
  aggregate: OperationCapability | false;
  pagination: (OperationCapability & { modes: readonly PaginationRequest["kind"][] }) | false;
  edit: OperationCapability | false;
  bulkEdit: OperationCapability | false;
  metadata: OperationCapability | false;
  validation: OperationCapability | false;
  formula: FormulaCapability;
  subscription: boolean;
  undo: OperationCapability | false;
  export: OperationCapability | false;
};

export function createLocalTableCapabilities(options: { formula?: FormulaCapability; undo?: boolean } = {}): TableCapabilities;
export function resolveTableOperationStates(
  capabilities: TableCapabilities,
  configuration: TableFeatureConfiguration
): Readonly<Record<TableFeature, TableOperationState>>;
```

Use `{ executor: "client", scope: "completeDataset" }` for every supported local operation, including metadata. Support `none` and `offset` pagination locally. Default formulas to `none`, subscription to `false`, and undo to enabled unless `options.undo === false`. Resolution order is host configuration, source support, required-scope compatibility, then scope label. A feature with `{ requiredScope: "completeDataset" }` resolves disabled with a precise reason when the source advertises only loaded rows. Formula `none` must resolve to the exact reason in the test.

- [ ] **Step 8: Export core query/capability APIs and verify both suites**

Create `src/table/core/index.ts` with explicit named and type exports from `query.ts` and `capabilities.ts`; do not use a broad wildcard that can later hide name collisions.

Run:

```bash
corepack pnpm exec vitest run src/table/core/query.test.ts src/table/core/capabilities.test.ts
corepack pnpm run build
```

Expected: both test files and TypeScript build PASS.

- [ ] **Step 9: Commit query and capability contracts**

```bash
git add src/table/core/query.ts src/table/core/query.test.ts src/table/core/capabilities.ts src/table/core/capabilities.test.ts src/table/core/index.ts
git commit -m "feat: define shared table query contracts"
```

### Task 2: Define generic columns, snapshots, intents, and sessions

**Files:**
- Create: `src/table/core/types.ts`
- Create: `src/table/core/columnHelper.ts`
- Create: `src/table/core/columnHelper.test.ts`
- Create: `src/table/core/commandId.ts`
- Create: `src/table/core/commandId.test.ts`
- Create: `src/table/core/safeInvoke.ts`
- Create: `src/table/core/safeInvoke.test.ts`
- Create: `src/table/core/types.types.test.tsx`
- Modify: `src/table/core/index.ts`

**Interfaces:**
- Consumes: `CommandResult`, `TableIssue`, `ParsedCellInput`, query contracts, and capabilities from Task 1.
- Produces: `ColumnDef`, `createColumnHelper`, `CommandIdFactory`, `createCommandIdFactory`, `safeInvokeTableExtension`, `TableIntent`, `TableViewSnapshot`, `TableRowView`, `TableSession`, the platform-neutral `ExportArtifact` and `TableAbortSignal` contracts, metadata/state/update contracts, `operationStates`, `pendingOperations`, `conflicts`, and stable selection types.

- [ ] **Step 1: Write failing column-helper behavior and type-inference tests**

Create `src/table/core/columnHelper.test.ts`:

```ts
import { describe, expect, it } from "vitest";
import { createColumnHelper, normalizeColumns } from "./columnHelper";

type Employee = { id: string; name: string; salary: number };

describe("column definitions", () => {
  it("creates immutable accessor updates with inferred values", () => {
    const column = createColumnHelper<Employee>().accessor("salary", {
      id: "salary",
      header: "Salary",
      dataType: "number"
    });
    const original: Employee = { id: "e1", name: "Ada", salary: 100 };
    const updated = column.update!(original, 125);
    expect(updated).toEqual({ id: "e1", name: "Ada", salary: 125 });
    expect(original.salary).toBe(100);
  });

  it("rejects blank and duplicate stable column ids", () => {
    const helper = createColumnHelper<Employee>();
    expect(() => normalizeColumns([
      helper.accessor("name", { id: "name", header: "Name" }),
      helper.accessor("salary", { id: "name", header: "Salary" })
    ])).toThrow("Duplicate column id: name");
    expect(() => normalizeColumns([
      helper.display({ id: " ", header: "Actions" })
    ])).toThrow("Column id must not be blank");
  });
});
```

Create `src/table/core/types.types.test.tsx` with compile-time assertions:

```tsx
import { expectTypeOf, test } from "vitest";
import { createColumnHelper, type ColumnDef, type ExportArtifact, type TableIntent, type TableSession } from "./index";

type Employee = { id: string; name: string; salary: number };

test("column helper preserves accessor value types", () => {
  const salary = createColumnHelper<Employee>().accessor("salary", {
    id: "salary",
    header: "Salary"
  });
  expectTypeOf(salary.accessor({ id: "1", name: "Ada", salary: 42 })).toEqualTypeOf<number>();
  expectTypeOf<TableSession<Employee>["getSnapshot"]>().toBeFunction();
  expectTypeOf<Awaited<ReturnType<TableSession<Employee>["export"]>>>().toEqualTypeOf<ExportArtifact>();
  expectTypeOf<ColumnDef<Employee>>().toBeObject();

  const reload: TableIntent<Employee> = {
    type: "reload-authoritative",
    operationId: "operation-1",
    rowId: "employee-1"
  };
  const retry: TableIntent<Employee> = {
    type: "retry-with-revision",
    operationId: "operation-1",
    rowId: "employee-1",
    expectedRevision: "revision-2"
  };
  expectTypeOf(reload.operationId).toEqualTypeOf<string>();
  expectTypeOf(retry.expectedRevision).toEqualTypeOf<string>();

  // @ts-expect-error retry requires the authoritative revision being accepted
  const missingRevision: TableIntent<Employee> = {
    type: "retry-with-revision",
    operationId: "operation-1",
    rowId: "employee-1"
  };
  expectTypeOf(missingRevision).toMatchTypeOf<TableIntent<Employee>>();
});
```

- [ ] **Step 2: Run focused tests and verify red**

Run:

```bash
corepack pnpm exec vitest run src/table/core/columnHelper.test.ts src/table/core/types.types.test.tsx
corepack pnpm run build
```

Expected: FAIL because the generic column/session modules are missing.

- [ ] **Step 3: Define the core column and value pipeline contracts**

Create `src/table/core/types.ts`. Import `CommandResult` and `TableIssue` from `../../core/commands/types`, `ParsedCellInput` from `../../core/values/parseCellInput`, and Task 1 types by relative path. Define:

```ts
export type TableDataType = "text" | "number" | "boolean" | "date" | "datetime" | "custom";
export type TableCellRef = { rowId: string; columnId: string };
export type TableSelection = { anchor: TableCellRef; focus: TableCellRef };
export type TableCellIssue = TableIssue & { rowId?: string; columnId?: string };
export type ColumnParseResult<TValue> =
  | { ok: true; value: TValue; formula?: string; evaluatedValue?: unknown }
  | { ok: false; issues: readonly TableCellIssue[] };

export type ColumnValueContext<TRow> = {
  row: TRow;
  rowId: string;
  columnId: string;
  getValue(columnId: string): unknown;
};

export type ColumnValidationContext<TRow, TValue> = ColumnValueContext<TRow> & {
  raw: string;
  parsed: TValue;
  evaluated: unknown;
};

export type TableHeaderAction<TRow> = {
  id: string;
  label: string;
  disabled?: boolean | ((rows: readonly TRow[]) => boolean);
  run(context: { columnId: string; rows: readonly TRow[] }): void | Promise<void>;
};

export type ColumnBaseDef<
  TRow,
  TValue = unknown,
  TCellRenderer = unknown,
  THeaderRenderer = unknown,
  TEditor = unknown
> = {
  id: string;
  header: string | THeaderRenderer;
  dataType?: TableDataType;
  parse?(input: ParsedCellInput, context: ColumnValueContext<TRow>): ColumnParseResult<TValue>;
  format?(value: TValue, context: ColumnValueContext<TRow>): string;
  validate?(context: ColumnValidationContext<TRow, TValue>): readonly TableCellIssue[];
  editable?: boolean | ((context: ColumnValueContext<TRow>) => boolean);
  permitted?: (context: ColumnValueContext<TRow>) => boolean;
  compare?(left: TValue, right: TValue): number;
  sortable?: boolean;
  filterable?: boolean;
  groupable?: boolean;
  aggregatable?: readonly ("sum" | "average" | "count" | "min" | "max")[];
  aggregate?: "sum" | "average" | "count" | "min" | "max";
  headerActions?: readonly TableHeaderAction<TRow>[];
  width?: number;
  minWidth?: number;
  maxWidth?: number;
  visible?: boolean;
  pin?: "left" | "right" | false;
  cell?: TCellRenderer;
  editor?: TEditor;
  meta?: Readonly<Record<string, unknown>>;
};

export type ColumnDef<
  TRow,
  TValue = unknown,
  TCellRenderer = unknown,
  THeaderRenderer = unknown,
  TEditor = unknown
> = ColumnBaseDef<TRow, TValue, TCellRenderer, THeaderRenderer, TEditor> & (
  | { kind?: "accessor"; accessor(row: TRow): TValue; update?(row: TRow, value: TValue): TRow }
  | { kind: "computed"; calculate(context: ColumnValueContext<TRow>): TValue; update?: never }
  | { kind: "display"; accessor?: never; calculate?: never; update?: never }
);
```

Define the metadata, durable/view state, updater, and export contracts in the same file rather than leaving adapter-specific aliases:

```ts
export type TableCellFormat = {
  numberFormat?: "general" | "number" | "currency" | "percent" | "date" | "datetime";
  textColor?: string;
  backgroundColor?: string;
  bold?: boolean;
  italic?: boolean;
  fontFamily?: string;
  fontSize?: number;
  horizontalAlign?: "left" | "center" | "right";
  verticalAlign?: "top" | "middle" | "bottom";
  wrapText?: boolean;
};

export type TableValidation =
  | { kind: "list"; values: readonly string[]; allowBlank?: boolean }
  | { kind: "number"; min?: number; max?: number; allowBlank?: boolean }
  | { kind: "textLength"; min?: number; max?: number; allowBlank?: boolean };

export type TableCellMetadata = {
  format?: TableCellFormat;
  validation?: TableValidation;
  comment?: string;
  formula?: string;
  readOnly?: boolean;
};

export type CalculatedColumnDefinition = { columnId: string; expression: string };
export type NamedTableStyle = { id: string; name: string; format: TableCellFormat };
export type TableMetadataDocument = {
  version: 1;
  cells: Readonly<Record<string, TableCellMetadata>>;
  calculatedColumns: readonly CalculatedColumnDefinition[];
  namedStyles: readonly NamedTableStyle[];
};
export type TableCellMetadataUpdate = { rowId: string; columnId: string; patch: Partial<TableCellMetadata> };

export type TableViewState = {
  sorting: readonly TableSort[];
  filter: FilterExpression | null;
  grouping: readonly TableGrouping[];
  aggregates: readonly TableAggregateRequest[];
  pagination: PaginationRequest;
  selection: TableSelection | null;
  selectedRowIds: readonly string[];
  expandedRowIds: readonly string[];
  columnOrder: readonly string[];
  columnVisibility: Readonly<Record<string, boolean>>;
  columnWidths: Readonly<Record<string, number>>;
  columnPinning: { left: readonly string[]; right: readonly string[] };
};

export type ChangeContext = {
  commandId: string;
  transactionId?: string;
  reason: TableIntent["type"];
  revision: string;
};
export type RowUpdater<TRow> = (previous: readonly TRow[]) => readonly TRow[];
export type TableMetadataUpdater = (previous: TableMetadataDocument) => TableMetadataDocument;
export type TableStateUpdater = (previous: TableViewState) => TableViewState;
export type ExportOptions = {
  format: "csv" | "xlsx";
  scope: "currentView" | "completeDataset";
  includeHeaders?: boolean;
  fileName?: string;
};
export type ExportArtifact = {
  bytes: Uint8Array;
  mediaType: string;
  fileName: string;
};
export type TableAbortSignal = {
  readonly aborted: boolean;
  readonly reason?: unknown;
  addEventListener(type: "abort", listener: () => void, options?: { once?: boolean }): void;
  removeEventListener(type: "abort", listener: () => void): void;
};
export type TableDiagnosticEvent = {
  category: "command" | "validation" | "extension" | "performance" | "remote";
  commandId?: string;
  durationMs?: number;
  metadata: Readonly<Record<string, string | number | boolean | null>>;
};
```

Metadata keys use `JSON.stringify([rowId, columnId])`, never delimiter concatenation. Supplying a controlled state slice makes the host responsible only for keys present in the supplied `state` object.

- [ ] **Step 4: Define source-neutral command, operation, and snapshot contracts**

Add these public shapes to `types.ts`:

```ts
export type TableCellEdit = { rowId: string; columnId: string; rawText: string };
export type RowAnchor =
  | { beforeRowId: string; afterRowId?: never }
  | { beforeRowId?: never; afterRowId: string }
  | { beforeRowId?: never; afterRowId?: never };

export type TableIntent<TRow = unknown> =
  | { type: "edit-cells"; edits: readonly TableCellEdit[] }
  | { type: "clear-cells"; cells: readonly TableCellRef[] }
  | { type: "update-cell-metadata"; updates: readonly TableCellMetadataUpdate[] }
  | { type: "set-selection"; selection: TableSelection | null }
  | { type: "set-row-selection"; rowIds: readonly string[] }
  | { type: "set-row-expanded"; rowId: string; expanded: boolean }
  | { type: "set-sorting"; sorting: readonly TableSort[] }
  | { type: "set-filter"; filter: FilterExpression | null }
  | { type: "set-grouping"; grouping: readonly TableGrouping[] }
  | { type: "set-aggregates"; aggregates: readonly TableAggregateRequest[] }
  | { type: "set-pagination"; pagination: PaginationRequest }
  | { type: "set-column-order"; columnIds: readonly string[] }
  | { type: "resize-column"; columnId: string; width: number }
  | { type: "set-column-visibility"; columnId: string; visible: boolean }
  | { type: "set-column-pinning"; columnId: string; pin: "left" | "right" | false }
  | ({ type: "insert-rows"; rows: readonly TRow[] } & RowAnchor)
  | ({ type: "insert-rows"; count: number } & RowAnchor)
  | { type: "delete-rows"; rowIds: readonly string[] }
  | { type: "reload-authoritative"; operationId: string; rowId: string }
  | { type: "retry-with-revision"; operationId: string; rowId: string; expectedRevision: string }
  | { type: "undo" }
  | { type: "redo" }
  | { type: "refresh" };

export type TablePendingOperation = {
  id: string;
  feature: TableFeature;
  startedAt: number;
  rowIds: readonly string[];
  cells: readonly TableCellRef[];
};

export type TableConflict<TRow> = {
  operationId: string;
  rowId: string;
  columnId?: string;
  attemptedValue?: unknown;
  authoritativeValue?: unknown;
  current: TRow;
  revision: string;
};

export type TableCellSnapshot = {
  rowId: string;
  columnId: string;
  storedValue: unknown;
  evaluatedValue: unknown;
  displayValue: string;
  formula?: string;
  metadata: TableCellMetadata;
  editable: boolean;
  issues: readonly TableCellIssue[];
};

export type TableRowSnapshot<TRow> = QueryRow<TRow>;

export interface TableViewSnapshot<TRow, TColumn = ColumnDef<TRow>> {
  revision: string;
  rows: readonly TableRowSnapshot<TRow>[];
  columns: readonly TColumn[];
  rowCount: number;
  totalRowCount: TotalCount;
  completeness: QueryResult<TRow>["completeness"];
  state: TableViewState;
  selection: TableSelection | null;
  status: { phase: "idle" | "loading" | "ready" | "error"; message?: string };
  issues: readonly TableCellIssue[];
  capabilities: TableCapabilities;
  operationStates: Readonly<Record<TableFeature, TableOperationState>>;
  pendingOperations: readonly TablePendingOperation[];
  conflicts: readonly TableConflict<TRow>[];
  canUndo: boolean;
  canRedo: boolean;
  pageInfo: QueryResult<TRow>["pageInfo"];
  getCell(rowId: string, columnId: string): TableCellSnapshot;
  getRowIndex(rowId: string): number;
  getColumnIndex(columnId: string): number;
}

export interface TableRowView<TRow, TColumn = ColumnDef<TRow>> {
  getSnapshot(): TableViewSnapshot<TRow, TColumn>;
  subscribe(listener: () => void): () => void;
  dispatch(intent: TableIntent<TRow>): Promise<CommandResult>;
}

export interface TableSession<TRow, TColumn = ColumnDef<TRow>> extends TableRowView<TRow, TColumn> {
  refresh(): Promise<void>;
  undo(): Promise<CommandResult>;
  redo(): Promise<CommandResult>;
  export(options: ExportOptions): Promise<ExportArtifact>;
  destroy(): void;
}
```

The two conflict-resolution intents are source-neutral so React never calls remote-only methods. `operationId` identifies the existing conflict and `rowId` prevents positional resolution; `retry-with-revision.expectedRevision` is the authoritative revision explicitly accepted by the host. Local `RecordTableSession` and workbook `WorkbookTableSession` have no authoritative server conflict lifecycle and must return `{ status: "rejected", reason: "unsupported" }` for both intents without changing revision, history, or subscribers. Only a remote adapter with a currently matching conflict may implement them.

`ExportArtifact` is the only export payload permitted in `src/table/core`, `src/table/local`, `src/table/remote`, and workbook adapters. It deliberately contains no `Blob`, `File`, object URL, download, or DOM type. React may convert its owned byte copy into a `Blob` at the component boundary. `TableAbortSignal` is the public structural subset implemented by native abort signals, so the core declarations do not require `lib.dom`. Local snapshots publish empty `pendingOperations` and `conflicts`; remote and workbook adapters reuse the same fields instead of inventing incompatible UI state.

- [ ] **Step 5: Implement the generic column helper and normalization**

Create `src/table/core/columnHelper.ts`. `createColumnHelper<TRow>()` must expose:

```ts
accessor<TKey extends keyof TRow>(
  key: TKey,
  options: ColumnBaseDef<TRow, TRow[TKey]>
): ColumnDef<TRow, TRow[TKey]>;
computed<TValue>(
  options: ColumnBaseDef<TRow, TValue> & { calculate(context: ColumnValueContext<TRow>): TValue }
): ColumnDef<TRow, TValue>;
display(
  options: ColumnBaseDef<TRow>
): ColumnDef<TRow>;
```

The accessor implementation reads `row[key]` and updates object rows using `{ ...row, [key]: value }`. `normalizeColumns` trims IDs, rejects blanks/duplicates, clamps default/min/max widths using the existing spreadsheet dimension bounds, and returns a frozen array without mutating the definitions.

- [ ] **Step 6: Write and implement collision-resistant per-session command IDs**

Create `src/table/core/commandId.test.ts` first:

```ts
import { describe, expect, it } from "vitest";
import { createCommandIdFactory } from "./commandId";

describe("createCommandIdFactory", () => {
  it("combines a collision-resistant session id with a monotonic local sequence", () => {
    const next = createCommandIdFactory(() => "session-uuid");
    expect(next()).toBe("session-uuid:1");
    expect(next()).toBe("session-uuid:2");
  });

  it("keeps independent sessions distinct", () => {
    const first = createCommandIdFactory(() => "first-uuid");
    const second = createCommandIdFactory(() => "second-uuid");
    expect(first()).not.toBe(second());
  });
});
```

Run: `corepack pnpm exec vitest run src/table/core/commandId.test.ts`

Expected: FAIL because `commandId.ts` is missing.

Create `src/table/core/commandId.ts`:

```ts
export type CommandIdFactory = () => string;

function defaultSessionId(): string {
  if (typeof globalThis.crypto?.randomUUID !== "function") {
    throw new Error("crypto.randomUUID is required to create table command ids");
  }
  return globalThis.crypto.randomUUID();
}

export function createCommandIdFactory(createSessionId: () => string = defaultSessionId): CommandIdFactory {
  const sessionId = createSessionId();
  let sequence = 0;
  return () => `${sessionId}:${++sequence}`;
}
```

Each session constructs one factory; do not share a counter at module scope. Tests and deterministic adapters may inject a session-ID function, while normal Node 22/browsers use Web Crypto UUIDs.

- [ ] **Step 7: Add safe host-extension invocation**

Create `src/table/core/safeInvoke.test.ts` first and cover accessor, calculate, format, parse, validate, update, editable/permitted predicate, formula-service, controlled callback, and subscriber exceptions. Create:

```ts
export type TableExtensionKind =
  | "accessor" | "calculate" | "format" | "parse" | "validate"
  | "update" | "permission" | "formula" | "host-callback" | "subscriber";

export function safeInvokeTableExtension<T>(
  kind: TableExtensionKind,
  invoke: () => T
): { ok: true; value: T } | { ok: false; issue: TableCellIssue };
```

Return `TABLE_EXTENSION_ERROR` with a sanitized `Host <kind> extension failed` message; never include thrown error text or values. During dispatch, invoke every row/cell extension against temporary state and reject the whole command before commit on failure. During snapshot/projection, return a stable error value/cell issue and `#ERROR!` display without crashing sibling cells. A controlled callback or subscriber that throws after a commit is caught, reported through sanitized diagnostics, and does not prevent remaining callbacks/subscribers. `getRowId`/column normalization failures are typed construction errors and create no live session.

Run: `corepack pnpm exec vitest run src/table/core/safeInvoke.test.ts`

Expected: PASS after the helper is implemented.

- [ ] **Step 8: Export the contracts and verify runtime plus compile-time tests**

Add explicit exports to `src/table/core/index.ts`.

Run:

```bash
corepack pnpm exec vitest run src/table/core/columnHelper.test.ts src/table/core/commandId.test.ts src/table/core/safeInvoke.test.ts src/table/core/types.types.test.tsx
corepack pnpm run build
```

Expected: both tests and strict TypeScript build PASS. The `src/table/core` dependency graph must contain no React or DOM imports.

- [ ] **Step 9: Commit generic table contracts**

```bash
git add src/table/core/types.ts src/table/core/columnHelper.ts src/table/core/columnHelper.test.ts src/table/core/commandId.ts src/table/core/commandId.test.ts src/table/core/safeInvoke.ts src/table/core/safeInvoke.test.ts src/table/core/types.types.test.tsx src/table/core/index.ts
git commit -m "feat: define generic table session contracts"
```

### Task 3: Build typed complete-dataset local row projection

**Files:**
- Create: `src/table/local/localFilter.ts`
- Create: `src/table/local/localAggregates.ts`
- Create: `src/table/local/localRowModel.ts`
- Create: `src/table/local/localRowModel.test.ts`
- Create: `src/table/local/index.ts`

**Interfaces:**
- Consumes: `ColumnDef`, `QueryRequest`, and stable `getRowId`.
- Produces: `buildLocalRowModel`, `LocalRowModel`, typed filtering/comparison, grouping, aggregation, and offset pagination over the complete local dataset.

- [ ] **Step 1: Write failing typed projection tests**

Create `src/table/local/localRowModel.test.ts` with this self-contained fixture:

```ts
import { describe, expect, it } from "vitest";
import { createColumnHelper } from "../core/columnHelper";
import type { QueryRequest } from "../core/query";
import { buildLocalRowModel } from "./localRowModel";

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
});
```

- [ ] **Step 2: Run projection tests and verify red**

Run: `corepack pnpm exec vitest run src/table/local/localRowModel.test.ts`

Expected: FAIL because the local row-model modules do not exist.

- [ ] **Step 3: Implement typed scalar comparison and filtering**

In `localFilter.ts`, implement and export:

```ts
export function matchesFilter<TRow>(
  row: TRow,
  rowId: string,
  expression: FilterExpression,
  columns: ReadonlyMap<string, ColumnDef<TRow>>,
  getValue: (row: TRow, rowId: string, columnId: string) => unknown
): boolean;
```

Resolve comparison operands only through stable column IDs. `isNull` matches only `null`; `isEmpty` matches only `""`; `isBlank` matches either and never `0`/`false`. String operators use locale-insensitive lower-case comparison. Numeric/date/boolean ordering requires matching typed scalars; incompatible types do not coerce and therefore do not match.

- [ ] **Step 4: Implement aggregation and the row-model pipeline**

In `localAggregates.ts`, implement `sum`, `average`, `count`, `min`, and `max` over non-null typed values; reject unknown aggregate column IDs.

In `localRowModel.ts`, export:

```ts
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

export function buildLocalRowModel<TRow>(
  rows: readonly TRow[],
  columns: readonly ColumnDef<TRow>[],
  request: QueryRequest,
  getRowId: (row: TRow) => string,
  resolveValue?: LocalValueResolver<TRow>,
  getSubRows?: (row: TRow) => readonly TRow[] | undefined
): LocalRowModel<TRow>;
```

The default resolver evaluates accessors/calculated columns lazily with a per-row stable-column cycle guard. `RecordTableSession` supplies a resolver that overlays metadata formulas through `RecordFormulaService`. Sorting, filtering, grouping, and all aggregates must call the same resolver; none may fall back to raw accessors or display strings. Ordinary typed values sort by type, errors remain stable after ordinary values, and blanks sort last. Filters do not coerce errors to blank/zero. Groups represent errors as tagged `{ type: "error", value: code }`, and numeric aggregates ignore errors while appending one deduplicated `TableCellIssue` per affected row/column to `LocalRowModel.issues`. The session copies those issues into `TableViewSnapshot.issues` and emits only sanitized issue counts/codes in diagnostics.

When `getSubRows` is present, recursively validate globally unique stable IDs, reject cycles/shared child objects by ID, and flatten only roots plus descendants whose ancestor IDs appear in `request.tree?.expandedRowIds`. Populate each data `QueryRow` with depth, parentId, hasChildren, and expanded; collapsed descendants remain in `totalDataRowCount` but not rendered items. Normalize IDs once; filter before sorting; sort siblings with original source index as the final tie-breaker; reject grouping combined with non-`none` pagination; group by tagged scalar values and build collision-proof group IDs from `JSON.stringify(["group", ...ancestorColumnAndScalarPairs, columnId, scalar])`; otherwise apply offset pagination to ungrouped flat data rows. Tree mode rejects grouping and pagination in the first release because hierarchical parent/child order must not be silently flattened; filter retains ancestors of matching descendants. Reject cursor/infinite requests as unsupported locally instead of treating them like offset pages. Never mutate `rows` or row objects.

- [ ] **Step 5: Add nested-filter, typed-date, stable-sort, and immutability cases**

Extend the same test file with named cases for nested `and/or/not`, date/datetime ordering, stable equal-value sort, group-ID collision including a repeated child key beneath two different parents, grouping-plus-pagination rejection at offset boundaries, tree expand/collapse, retained matching ancestors, sibling sort, duplicate/cyclic tree IDs, tree/group/pagination rejection, all five aggregates, unknown column rejection, unsupported cursor pagination, and frozen input rows remaining unchanged. Inject a resolver whose formula results are wrapped numeric values `2` and `10` and prove sort, filter, group, and aggregate use those values; add calculated-column dependency, cycle error, formula-error ordering, issue de-duplication, and row/column issue-address cases.

- [ ] **Step 6: Run focused tests and build**

Run:

```bash
corepack pnpm exec vitest run src/table/local/localRowModel.test.ts
corepack pnpm run build
```

Expected: all projection tests and strict TypeScript build PASS.

- [ ] **Step 7: Export and commit local projection**

Export `buildLocalRowModel` and its public types from `src/table/local/index.ts`.

```bash
git add src/table/local/localFilter.ts src/table/local/localAggregates.ts src/table/local/localRowModel.ts src/table/local/localRowModel.test.ts src/table/local/index.ts
git commit -m "feat: project complete local table datasets"
```

### Task 4: Implement atomic local record edits, metadata, and controlled state

**Files:**
- Create: `src/table/local/RecordTableSession.ts`
- Create: `src/table/local/RecordTableSession.test.ts`
- Create: `src/table/local/tableMetadata.ts`
- Create: `src/table/local/tableMetadata.test.ts`
- Create: `src/table/core/session.contract.ts`
- Create: `src/table/core/session.contract.test.ts`
- Modify: `src/table/local/index.ts`

**Interfaces:**
- Consumes: `parseCellInput`, `CommandResult`, Task 2 session/column types, and Task 3 row projection.
- Produces: `LocalRecordSource`, `LocalRecordTableSessionOptions`, `RecordFormulaService`, `RecordTableSession`, and the canonical factory `createLocalRecordTableSession`.

- [ ] **Step 1: Write failing metadata-key tests**

Create `src/table/local/tableMetadata.test.ts`:

```ts
import { describe, expect, it } from "vitest";
import { createTableMetadataKey, parseTableMetadataKey } from "./tableMetadata";

describe("table metadata keys", () => {
  it("round-trips ids containing separators without collisions", () => {
    const first = createTableMetadataKey("row::1", "column/2");
    const second = createTableMetadataKey("row", "1::column/2");
    expect(first).not.toBe(second);
    expect(parseTableMetadataKey(first)).toEqual({ rowId: "row::1", columnId: "column/2" });
    expect(parseTableMetadataKey(second)).toEqual({ rowId: "row", columnId: "1::column/2" });
  });
});
```

- [ ] **Step 2: Run the metadata test and verify red**

Run: `corepack pnpm exec vitest run src/table/local/tableMetadata.test.ts`

Expected: FAIL because `tableMetadata.ts` is missing.

- [ ] **Step 3: Implement collision-safe metadata addressing**

Create `src/table/local/tableMetadata.ts`:

```ts
export function createTableMetadataKey(rowId: string, columnId: string): string {
  return JSON.stringify([rowId, columnId]);
}

export function parseTableMetadataKey(key: string): { rowId: string; columnId: string } {
  const parsed: unknown = JSON.parse(key);
  if (!Array.isArray(parsed) || parsed.length !== 2 || parsed.some((value) => typeof value !== "string")) {
    throw new Error("Invalid table metadata key");
  }
  return { rowId: parsed[0], columnId: parsed[1] };
}
```

Run: `corepack pnpm exec vitest run src/table/local/tableMetadata.test.ts`

Expected: PASS.

- [ ] **Step 4: Write failing local-session transaction tests**

Create `src/table/local/RecordTableSession.test.ts`:

```ts
import { describe, expect, it, vi } from "vitest";
import { createColumnHelper } from "../core/columnHelper";
import { createLocalRecordTableSession } from "./RecordTableSession";

type Employee = { id: string; name: string; salary: number; active: boolean };

function createColumns() {
  const helper = createColumnHelper<Employee>();
  return [
    helper.accessor("name", { id: "name", header: "Name", dataType: "text" }),
    helper.accessor("salary", {
      id: "salary",
      header: "Salary",
      dataType: "number",
      validate: ({ parsed }) => parsed < 0
        ? [{ code: "negative-salary", message: "Salary must be non-negative", columnId: "salary" }]
        : []
    }),
    helper.accessor("active", { id: "active", header: "Active", dataType: "boolean" })
  ] as const;
}

describe("createLocalRecordTableSession", () => {
  it("parses a typed controlled edit and emits one replayable updater", async () => {
    const original: readonly Employee[] = [{ id: "e1", name: "Ada", salary: 100, active: true }];
    const onRowsChange = vi.fn();
    const session = createLocalRecordTableSession({
      source: { kind: "local", rows: original, getRowId: (row) => row.id, onRowsChange },
      columns: createColumns()
    });

    const result = await session.dispatch({
      type: "edit-cells",
      edits: [{ rowId: "e1", columnId: "salary", rawText: "125" }]
    });
    const updater = onRowsChange.mock.calls[0][0] as (rows: readonly Employee[]) => readonly Employee[];

    expect(result).toMatchObject({ status: "committed", changed: true });
    expect(updater(original)).toEqual([{ id: "e1", name: "Ada", salary: 125, active: true }]);
    expect(onRowsChange).toHaveBeenCalledTimes(1);
    expect(session.getSnapshot().getCell("e1", "salary").storedValue).toBe(125);
  });

  it("rejects an invalid batch without partial rows, metadata, history, or publication", async () => {
    const rows: readonly Employee[] = [
      { id: "e1", name: "Ada", salary: 100, active: true },
      { id: "e2", name: "Grace", salary: 200, active: false }
    ];
    const onRowsChange = vi.fn();
    const listener = vi.fn();
    const session = createLocalRecordTableSession({
      source: { kind: "local", rows, getRowId: (row) => row.id, onRowsChange },
      columns: createColumns()
    });
    session.subscribe(listener);

    const result = await session.dispatch({
      type: "edit-cells",
      edits: [
        { rowId: "e1", columnId: "name", rawText: "Changed" },
        { rowId: "e2", columnId: "salary", rawText: "-1" }
      ]
    });

    expect(result).toEqual({
      status: "rejected",
      reason: "validation",
      issues: [{ code: "negative-salary", message: "Salary must be non-negative", columnId: "salary" }]
    });
    expect(onRowsChange).not.toHaveBeenCalled();
    expect(listener).not.toHaveBeenCalled();
    expect(session.getSnapshot().getCell("e1", "name").storedValue).toBe("Ada");
    expect(session.getSnapshot().canUndo).toBe(false);
  });

  it("publishes source-neutral operation state and empty local async state", () => {
    const session = createLocalRecordTableSession({
      source: {
        kind: "local",
        rows: [{ id: "e1", name: "Ada", salary: 100, active: true }],
        getRowId: (row) => row.id
      },
      columns: createColumns()
    });
    const snapshot = session.getSnapshot();
    expect(snapshot.operationStates.sort).toEqual({ enabled: true, scopeLabel: "Complete dataset" });
    expect(snapshot.pendingOperations).toEqual([]);
    expect(snapshot.conflicts).toEqual([]);
  });
});
```

- [ ] **Step 5: Run local-session tests and verify red**

Run: `corepack pnpm exec vitest run src/table/local/RecordTableSession.test.ts`

Expected: FAIL because `RecordTableSession.ts` is missing.

- [ ] **Step 6: Define local source, formula service, and session options**

Create `src/table/local/RecordTableSession.ts` with these exports:

```ts
export type LocalRecordSource<TRow> = {
  kind: "local";
  rows: readonly TRow[];
  getRowId(row: TRow): string;
  getSubRows?(row: TRow): readonly TRow[] | undefined;
  onRowsChange?(updater: RowUpdater<TRow>, context: ChangeContext): void;
  resetKey?: string | number;
};

export type RecordFormulaService<TRow> = {
  evaluate(context: {
    expression: string;
    row: TRow;
    rowId: string;
    columnId: string;
    getValue(columnId: string): unknown;
  }): { value: unknown; displayValue: string } | { issues: readonly TableCellIssue[] };
};

export type LocalRecordTableSessionOptions<TRow, TColumn = ColumnDef<TRow>> = {
  source: LocalRecordSource<TRow>;
  columns: readonly TColumn[];
  document?: TableMetadataDocument;
  defaultDocument?: TableMetadataDocument;
  onDocumentChange?(updater: TableMetadataUpdater, context: ChangeContext): void;
  state?: Partial<TableViewState>;
  defaultState?: Partial<TableViewState>;
  onStateChange?(updater: TableStateUpdater, context: ChangeContext): void;
  formulaService?: RecordFormulaService<TRow>;
  historyLimit?: number;
  features?: TableFeatureConfiguration;
  commandIdFactory?: CommandIdFactory;
  onDiagnostic?(event: TableDiagnosticEvent): void;
};

export class RecordTableSession<TRow, TColumn extends ColumnDef<TRow> = ColumnDef<TRow>>
  implements TableSession<TRow, TColumn> {
  constructor(options: LocalRecordTableSessionOptions<TRow, TColumn>);
  updateOptions(options: LocalRecordTableSessionOptions<TRow, TColumn>): void;
  getSnapshot(): TableViewSnapshot<TRow, TColumn>;
  subscribe(listener: () => void): () => void;
  dispatch(intent: TableIntent<TRow>): Promise<CommandResult>;
  refresh(): Promise<void>;
  undo(): Promise<CommandResult>;
  redo(): Promise<CommandResult>;
  export(options: ExportOptions): Promise<ExportArtifact>;
  destroy(): void;
}

export function createLocalRecordTableSession<TRow, TColumn extends ColumnDef<TRow> = ColumnDef<TRow>>(
  options: LocalRecordTableSessionOptions<TRow, TColumn>
): RecordTableSession<TRow, TColumn> {
  return new RecordTableSession(options);
}
```

Reject `document` together with `defaultDocument`, and reject controlled state/document values without their matching callback. Row authority is determined by callback presence: `onRowsChange` present is controlled; absent is an internally owned copy initialized from `source.rows`. In controlled mode, `updateOptions` synchronously accepts a changed host `rows` reference without publishing during React render. In uncontrolled mode, later `rows` prop identities are ignored unless `resetKey` changes; a reset replaces rows, clears edit history/redo, drops invalid selection, and preserves compatible view state. Switching controlled/uncontrolled mode is an explicit reset with the same rules. Preserve view state and history for unchanged controlled row identity/order.

Construct one `createCommandIdFactory()` per session when `commandIdFactory` is absent. Every public dispatch/undo/redo/refresh operation obtains exactly one ID before validation, and uses it consistently in `ChangeContext`, diagnostics, pending state, and results. Tests inject `createCommandIdFactory(() => "local-test")` for deterministic assertions.

Initialize uncontrolled values exactly as follows:

```ts
const EMPTY_TABLE_DOCUMENT: TableMetadataDocument = {
  version: 1,
  cells: {},
  calculatedColumns: [],
  namedStyles: []
};
const DEFAULT_TABLE_VIEW_STATE: TableViewState = {
  sorting: [],
  filter: null,
  grouping: [],
  aggregates: [],
  pagination: { kind: "none" },
  selection: null,
  selectedRowIds: [],
  expandedRowIds: [],
  columnOrder: [],
  columnVisibility: {},
  columnWidths: {},
  columnPinning: { left: [], right: [] }
};
```

- [ ] **Step 7: Implement atomic edit and lazy snapshot behavior**

For `edit-cells`, perform these operations against temporary row/document copies before changing private fields:

1. Resolve every row and column by stable ID.
2. Reject display/computed/read-only/unpermitted cells.
3. Call `parseCellInput(rawText)`.
4. Call the column parser if supplied; otherwise require a stored value compatible with `dataType`.
5. Evaluate formula text only when `formulaService` exists; otherwise reject with reason `unsupported`.
6. Call validation with raw, parsed, and evaluated candidates.
7. Apply column `update` functions only after every candidate passes.
8. Create one ID-based `RowUpdater`, one optional metadata updater, one inverse journal entry, and one new immutable snapshot.
9. Publish exactly once and call each controlled callback exactly once.

`getSnapshot().getCell(rowId, columnId)` resolves data/group/aggregate rows lazily, applies calculated columns and metadata, and returns display text without building a cell matrix. Track the current calculated-column evaluation stack by stable column ID; a cycle returns a `calculated-column-cycle` issue and spreadsheet-style error display instead of recursing. Memoize the snapshot object until revision, view state, rows, columns, or document changes.

Build every projected row model with the session's same lazy evaluated-value resolver used by `getCell`, including metadata formulas and calculated columns. Cache evaluations only for the current immutable session revision and clear them on row/document/column/formula changes.

View commands update only their matching state slice. A key present in `options.state` is controlled and must emit `onStateChange`; absent keys use internal state initialized from `defaultState`. Local snapshots always use `status.phase = "ready"`, `completeness = "completeDataset"`, `totalRowCount = { kind: "known", value: model.totalDataRowCount }`, empty pending/conflict arrays, and Task 1 operation states.

Maintain the grouping/pagination invariant before building a row model. Starting non-empty grouping atomically sets pagination to `{ kind: "none" }` in the same uncontrolled update or controlled `TableStateUpdater`. A non-`none` pagination intent while grouped is rejected with `TABLE_GROUPING_PAGINATION_CONFLICT`. If a controlled host supplies the invalid combination directly, preserve the last valid projection, set snapshot status/error plus a sanitized issue, disable pagination with the same reason, and request correction through `onStateChange`; `getSnapshot()` never throws.

Measure each dispatch with `performance.now()` when available and `Date.now()` otherwise. Send `onDiagnostic` only command ID/type, duration, changed/rejected category, and affected row/cell counts; never include row objects, stored/display values, formulas, raw edit text, comments, or validation payload values.

- [ ] **Step 8: Add controlled document/state, calculated column, formula, permission, and destroy tests**

Extend `RecordTableSession.test.ts` with explicit named tests that assert:

```ts
it("uses callback presence as the controlled row authority contract");
it("keeps an uncontrolled edit across rerenders with new rows array identities");
it("resets uncontrolled rows and history only when resetKey changes");
it("accepts a changed controlled rows reference synchronously");
it("resets safely when switching controlled and uncontrolled modes");
it("updates only the controlled sorting slice through onStateChange");
it("updates cell metadata through one collision-safe document updater");
it("recalculates a computed column after an accessor edit");
it("reports a calculated-column cycle as a cell issue without recursive overflow");
it("rejects formula text when no formula service is configured");
it("uses a configured formula service and validates its evaluated value");
it("rejects edits after a permission predicate changes");
it("suppresses no-op publications and reports changed false");
it("exports the current view and complete local dataset as platform-neutral artifacts with explicit scope");
it("rejects remote conflict-resolution intents without changing local state");
it("does not publish after destroy");
it("keeps two sessions completely isolated");
it("uses collision-resistant per-session command ids in change contexts and diagnostics");
```

Use only inline rows, `createColumns()`, `vi.fn()`, and directly constructed formula services in these tests; do not introduce test-only methods on the session.

Implement `RecordTableSession.export` against the same lazy projection and return `ExportArtifact`, never `Blob`. CSV uses RFC 4180 escaping; XLSX uses the already-installed ExcelJS dependency and native number/boolean/date/date-time values. Copy the produced bytes into a fresh `Uint8Array`, use `text/csv;charset=utf-8` or `application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`, and return a nonblank, extension-correct `fileName`. `scope: "currentView"` walks the snapshot's projected data rows, while `scope: "completeDataset"` rebuilds an unpaginated local query without mutating view state. Assert exact bytes/media type/file name in the local tests. Task 10 adds only the React-owned `Blob` conversion; it must not duplicate or replace the working session export implementation after the capability has advertised export support.

The dispatch switch handles `reload-authoritative` and `retry-with-revision` explicitly before any row lookup and returns `rejected/unsupported`. The test dispatches both typed intents with nonblank stable IDs and proves rows, document, view state, revision, history, callbacks, and subscriber counts are unchanged. The later workbook harness inherits the same source-neutral contract expectation and must also reject both intents because workbook conflicts are revision conflicts returned directly by workbook commands, not retained remote row conflicts.

- [ ] **Step 9: Add the shared capability-driven session contract**

Before the focused run, create the shared adapter contract in `src/table/core/session.contract.ts`. This file is test infrastructure and must not be exported from `src/table/core/index.ts` or any package entry:

```ts
import { describe, expect, it } from "vitest";
import type { TableFeature } from "./capabilities";
import type { CommandResult } from "../../core/commands/types";
import type { TableCellRef, TableIntent, TableSession } from "./types";

export type TableSessionContractHarness<TRow> = {
  session: TableSession<TRow>;
  editableCell: TableCellRef;
  validRawText: string;
  featureOperations: Partial<Record<TableFeature, () => Promise<CommandResult>>>;
  conflictResolution?: {
    reload: Extract<TableIntent<TRow>, { type: "reload-authoritative" }>;
    retry: Extract<TableIntent<TRow>, { type: "retry-with-revision" }>;
  };
  cleanup(): void;
};

export function defineTableSessionContract<TRow>(
  name: string,
  createHarness: () => TableSessionContractHarness<TRow>
): void {
  describe(`${name} TableSession contract`, () => {
    it("keeps snapshot identity stable until one publication", async () => {
      const harness = createHarness();
      const before = harness.session.getSnapshot();
      expect(harness.session.getSnapshot()).toBe(before);
      let publications = 0;
      harness.session.subscribe(() => { publications += 1; });
      const result = await harness.session.dispatch({
        type: "edit-cells",
        edits: [{ ...harness.editableCell, rawText: harness.validRawText }]
      });
      if (before.operationStates.edit.enabled) {
        expect(result.status).toBe("committed");
        expect(publications).toBe(1);
        expect(harness.session.getSnapshot()).not.toBe(before);
      } else {
        expect(result).toMatchObject({ status: "rejected", reason: "unsupported" });
        expect(publications).toBe(0);
      }
      harness.cleanup();
    });

    it("matches every advertised feature state to command behavior", async () => {
      const harness = createHarness();
      const snapshot = harness.session.getSnapshot();
      for (const [feature, operation] of Object.entries(harness.featureOperations) as Array<[TableFeature, () => Promise<CommandResult>]>) {
        const result = await operation();
        if (snapshot.operationStates[feature].enabled) {
          expect(result).not.toMatchObject({ status: "rejected", reason: "unsupported" });
        } else {
          expect(result).toMatchObject({ status: "rejected", reason: "unsupported" });
        }
      }
      harness.cleanup();
    });

    for (const kind of ["reload", "retry"] as const) {
      it(`${kind} conflict resolution is supported only by a configured source harness`, async () => {
        const harness = createHarness();
        const configured = harness.conflictResolution;
        const intent: TableIntent<TRow> = configured?.[kind] ?? (kind === "reload"
          ? { type: "reload-authoritative", operationId: "contract-conflict", rowId: "contract-row" }
          : {
              type: "retry-with-revision",
              operationId: "contract-conflict",
              rowId: "contract-row",
              expectedRevision: "contract-revision"
            });
        const result = await harness.session.dispatch(intent);
        if (configured) {
          expect(result).not.toMatchObject({ status: "rejected", reason: "unsupported" });
        } else {
          expect(result).toMatchObject({ status: "rejected", reason: "unsupported" });
        }
        harness.cleanup();
      });
    }
  });
}
```

Create `src/table/core/session.contract.test.ts` with the complete local harness:

```ts
import { expect } from "vitest";
import { createColumnHelper } from "./columnHelper";
import { defineTableSessionContract } from "./session.contract";
import { createLocalRecordTableSession } from "../local/RecordTableSession";

type Employee = { id: string; name: string; department: string; salary: number };

defineTableSessionContract<Employee>("local records", () => {
  const helper = createColumnHelper<Employee>();
  const session = createLocalRecordTableSession({
    source: {
      kind: "local",
      rows: [{ id: "e1", name: "Ada", department: "Finance", salary: 100 }],
      getRowId: (row) => row.id
    },
    columns: [
      helper.accessor("name", { id: "name", header: "Name" }),
      helper.accessor("department", { id: "department", header: "Department" }),
      helper.accessor("salary", { id: "salary", header: "Salary", dataType: "number" })
    ]
  });
  return {
    session,
    editableCell: { rowId: "e1", columnId: "name" },
    validRawText: "Grace",
    featureOperations: {
      sort: () => session.dispatch({ type: "set-sorting", sorting: [{ columnId: "name", direction: "asc" }] }),
      filter: () => session.dispatch({
        type: "set-filter",
        filter: { kind: "comparison", columnId: "department", operator: "eq", value: { type: "string", value: "Finance" } }
      }),
      group: () => session.dispatch({ type: "set-grouping", grouping: [{ columnId: "department" }] }),
      aggregate: () => session.dispatch({
        type: "set-aggregates",
        aggregates: [{ id: "salary-sum", columnId: "salary", function: "sum" }]
      }),
      pagination: () => session.dispatch({ type: "set-pagination", pagination: { kind: "offset", offset: 0, limit: 25 } }),
      edit: () => session.dispatch({
        type: "edit-cells",
        edits: [{ rowId: "e1", columnId: "name", rawText: "Katherine" }]
      }),
      formula: () => session.dispatch({
        type: "edit-cells",
        edits: [{ rowId: "e1", columnId: "salary", rawText: "=[salary]*2" }]
      }),
      undo: () => session.undo(),
      export: async () => {
        const artifact = await session.export({ format: "csv", scope: "completeDataset" });
        expect(artifact).toMatchObject({
          bytes: expect.any(Uint8Array),
          mediaType: "text/csv;charset=utf-8"
        });
        expect(artifact.fileName).toMatch(/\.csv$/);
        return { status: "committed" as const, revision: session.getSnapshot().revision };
      }
    },
    cleanup: () => session.destroy()
  };
});
```

Later remote/workbook plans invoke the same contract factory with their own advertised features and source-specific fixtures. The local harness intentionally omits `conflictResolution`, which proves both intents are unsupported. The workbook harness also omits it for the same assertion. The remote harness adds real, stable conflict intents only after Task 5 implements conflict retention and resolution; before that point its read-only harness must omit the field and reject them.

- [ ] **Step 10: Run local tests, shared contract, and strict build**

Run:

```bash
corepack pnpm exec vitest run \
  src/table/local/tableMetadata.test.ts \
  src/table/local/RecordTableSession.test.ts \
  src/table/core/session.contract.test.ts
corepack pnpm run build
```

Expected: all tests and TypeScript build PASS with no React act warnings or console errors.

- [ ] **Step 11: Export and commit the atomic local session**

Add explicit exports to `src/table/local/index.ts`.

```bash
git add src/table/local/RecordTableSession.ts src/table/local/RecordTableSession.test.ts src/table/local/tableMetadata.ts src/table/local/tableMetadata.test.ts src/table/local/index.ts src/table/core/session.contract.ts src/table/core/session.contract.test.ts
git commit -m "feat: add atomic local record sessions"
```

### Task 5: Add structural row commands and bounded inverse history

**Files:**
- Create: `src/table/local/localHistory.ts`
- Create: `src/table/local/localHistory.test.ts`
- Create: `src/table/local/RecordTableSession.property.test.ts`
- Modify: `src/table/local/RecordTableSession.ts`
- Modify: `src/table/local/RecordTableSession.test.ts`

**Interfaces:**
- Consumes: Task 4 atomic transactions.
- Produces: bounded local inverse operations for edit, metadata, insert, and delete commands; deterministic history statistics used by pressure tests.

- [ ] **Step 1: Verify the foundation's pinned property-test dependency**

Run:

```bash
corepack pnpm list fast-check --depth 0
```

Expected: exact `fast-check 4.9.0` is already present from Foundation Task 7. Do not change its version in this plan.

- [ ] **Step 2: Write failing bounded-history tests**

Create `src/table/local/localHistory.test.ts`:

```ts
import { describe, expect, it } from "vitest";
import { createLocalHistory, pushLocalHistory, redoLocalHistory, undoLocalHistory } from "./localHistory";

describe("local table history", () => {
  it("bounds entries, clears future after a new edit, and returns exact inverses", () => {
    let history = createLocalHistory<string>(2);
    history = pushLocalHistory(history, "undo-1", "redo-1");
    history = pushLocalHistory(history, "undo-2", "redo-2");
    history = pushLocalHistory(history, "undo-3", "redo-3");
    expect(history.past.map((entry) => entry.undo)).toEqual(["undo-2", "undo-3"]);

    const undone = undoLocalHistory(history);
    expect(undone.operation).toBe("undo-3");
    const redone = redoLocalHistory(undone.history);
    expect(redone.operation).toBe("redo-3");

    const replaced = pushLocalHistory(undone.history, "undo-4", "redo-4");
    expect(replaced.future).toEqual([]);
  });
});
```

- [ ] **Step 3: Run history tests and verify red**

Run: `corepack pnpm exec vitest run src/table/local/localHistory.test.ts`

Expected: FAIL because `localHistory.ts` is missing.

- [ ] **Step 4: Implement bounded inverse operation storage**

Create `src/table/local/localHistory.ts`:

```ts
export type LocalHistoryEntry<TOperation> = { undo: TOperation; redo: TOperation };
export type LocalHistory<TOperation> = {
  limit: number;
  past: readonly LocalHistoryEntry<TOperation>[];
  future: readonly LocalHistoryEntry<TOperation>[];
};

export function createLocalHistory<TOperation>(limit = 100): LocalHistory<TOperation>;
export function pushLocalHistory<TOperation>(
  history: LocalHistory<TOperation>,
  undo: TOperation,
  redo: TOperation
): LocalHistory<TOperation>;
export function undoLocalHistory<TOperation>(history: LocalHistory<TOperation>): {
  history: LocalHistory<TOperation>;
  operation: TOperation | null;
};
export function redoLocalHistory<TOperation>(history: LocalHistory<TOperation>): {
  history: LocalHistory<TOperation>;
  operation: TOperation | null;
};
```

Require an integer limit from 1 through 1,000. Keep inverse operations as affected row objects/order positions and affected metadata entries, never whole-dataset snapshots.

- [ ] **Step 5: Add structural session tests before implementation**

Extend `RecordTableSession.test.ts` with concrete rows and assertions for:

```ts
it("inserts supplied local records before a stable row id");
it("inserts supplied local records after a stable row id");
it("appends supplied local records when no anchor is present");
it("rejects count-only insertion because local records require host values");
it("rejects duplicate inserted row ids atomically");
it("deletes stable row ids regardless of current sort order");
it("undoes and redoes values, row order, and metadata together");
it("clears redo after a new durable command");
it("never retains more than the configured history limit");
```

Run: `corepack pnpm exec vitest run src/table/local/RecordTableSession.test.ts`

Expected: the new insert/delete/undo assertions FAIL while Task 4 cases remain green.

- [ ] **Step 6: Implement structural commands and inverse replay**

In `RecordTableSession.ts`, make record-carrying `insert-rows` validate all new host IDs and exactly one optional anchor before mutation. Count-only insertion returns `{ status: "rejected", reason: "unsupported" }` for local sources. `delete-rows` captures deleted records, original order positions, and metadata keys. Undo/redo replay the captured ID-based inverse inside the same atomic transaction path, emit controlled updaters, increment revision, and publish once. View-state commands never enter durable history.

- [ ] **Step 7: Write the seeded inverse property test**

Create `src/table/local/RecordTableSession.property.test.ts`:

```ts
import fc from "fast-check";
import { expect, it } from "vitest";
import { createColumnHelper } from "../core/columnHelper";
import { createLocalRecordTableSession } from "./RecordTableSession";

type Row = { id: string; value: number };

it("restores the exact original dataset after random edits are undone", async () => {
  await fc.assert(fc.asyncProperty(
    fc.array(fc.record({ rowIndex: fc.integer({ min: 0, max: 4 }), value: fc.integer() }), { maxLength: 100 }),
    async (operations) => {
      const original: readonly Row[] = Array.from({ length: 5 }, (_, index) => ({ id: `row-${index}`, value: index }));
      const helper = createColumnHelper<Row>();
      const session = createLocalRecordTableSession({
        source: { kind: "local", rows: original, getRowId: (row) => row.id },
        columns: [helper.accessor("value", { id: "value", header: "Value", dataType: "number" })],
        historyLimit: 100
      });

      for (const operation of operations) {
        await session.dispatch({
          type: "edit-cells",
          edits: [{ rowId: `row-${operation.rowIndex}`, columnId: "value", rawText: String(operation.value) }]
        });
      }
      for (let index = 0; index < Math.min(operations.length, 100); index += 1) {
        await session.undo();
      }

      const restored = original.map((row) => ({
        id: row.id,
        value: session.getSnapshot().getCell(row.id, "value").storedValue
      }));
      expect(restored).toEqual(original);
      session.destroy();
    }
  ), { seed: 20260709, numRuns: 200 });
});
```

- [ ] **Step 8: Run history, session, and property suites**

Run:

```bash
corepack pnpm exec vitest run \
  src/table/local/localHistory.test.ts \
  src/table/local/RecordTableSession.test.ts \
  src/table/local/RecordTableSession.property.test.ts
corepack pnpm run build
```

Expected: all tests and build PASS; the property test prints its fixed seed on failure.

- [ ] **Step 9: Commit structural local history**

```bash
git add package.json pnpm-lock.yaml src/table/local/localHistory.ts src/table/local/localHistory.test.ts src/table/local/RecordTableSession.ts src/table/local/RecordTableSession.test.ts src/table/local/RecordTableSession.property.test.ts
git commit -m "feat: add bounded local table history"
```

### Task 6: Extract a reusable React two-axis virtualizer from the spreadsheet

**Files:**
- Create: `src/react/viewport/useTwoAxisVirtualizer.ts`
- Create: `src/react/viewport/useTwoAxisVirtualizer.test.tsx`
- Modify: `src/components/Grid.tsx`
- Modify: `src/components/Grid.test.tsx`

**Interfaces:**
- Consumes: foundation `measureAxis` and `findVisibleRange` from `src/core/viewport/axis.ts`.
- Produces: `useTwoAxisVirtualizer`, visible row/column measurements, exact canvas dimensions, lazy cell rectangles, and `ensureCellVisible` shared by Spreadsheet and DataTable.

- [ ] **Step 1: Write a failing hook test with variable rows and columns**

Create `src/react/viewport/useTwoAxisVirtualizer.test.tsx`:

```tsx
import { fireEvent, render, screen } from "@testing-library/react";
import { useRef } from "react";
import { describe, expect, it } from "vitest";
import { useTwoAxisVirtualizer } from "./useTwoAxisVirtualizer";

function Harness() {
  const scrollRef = useRef<HTMLDivElement>(null);
  const virtualizer = useTwoAxisVirtualizer({
    scrollRef,
    rowCount: 100_000,
    columnCount: 10_000,
    getRowKey: (index) => `row-${index}`,
    getColumnKey: (index) => `column-${index}`,
    getRowSize: (index) => index === 0 ? 40 : 28,
    getColumnSize: (index) => index === 0 ? 160 : 96,
    rowOverscan: 4,
    columnOverscan: 2
  });
  return (
    <div>
      <div ref={scrollRef} data-testid="scroll" onScroll={virtualizer.onScroll} />
      <output data-testid="window">
        {JSON.stringify({
          rows: virtualizer.visibleRows.map((item) => item.index),
          columns: virtualizer.visibleColumns.map((item) => item.index),
          totalHeight: virtualizer.totalHeight,
          totalWidth: virtualizer.totalWidth
        })}
      </output>
    </div>
  );
}

describe("useTwoAxisVirtualizer", () => {
  it("bounds both axes and scrolls to a distant cell", () => {
    render(<Harness />);
    const scroll = screen.getByTestId("scroll");
    Object.defineProperties(scroll, {
      clientHeight: { configurable: true, value: 280 },
      clientWidth: { configurable: true, value: 480 },
      scrollTop: { configurable: true, writable: true, value: 28_000 },
      scrollLeft: { configurable: true, writable: true, value: 9_600 }
    });
    fireEvent.scroll(scroll);
    const windowState = JSON.parse(screen.getByTestId("window").textContent ?? "{}");
    expect(windowState.rows.length).toBeLessThan(24);
    expect(windowState.columns.length).toBeLessThan(12);
    expect(windowState.rows[0]).toBeGreaterThan(900);
    expect(windowState.columns[0]).toBeGreaterThan(90);
    expect(windowState.totalHeight).toBeGreaterThan(2_000_000);
    expect(windowState.totalWidth).toBeGreaterThan(900_000);
  });
});
```

- [ ] **Step 2: Run the hook test and verify red**

Run: `corepack pnpm exec vitest run src/react/viewport/useTwoAxisVirtualizer.test.tsx`

Expected: FAIL because the React virtualizer hook is missing.

- [ ] **Step 3: Implement the hook over foundation axis primitives**

Create `src/react/viewport/useTwoAxisVirtualizer.ts` with:

```ts
import type { RefObject, UIEvent as ReactUIEvent } from "react";

export type TwoAxisVirtualizerOptions = {
  scrollRef: RefObject<HTMLElement | null>;
  rowCount: number;
  columnCount: number;
  getRowKey(index: number): string;
  getColumnKey(index: number): string;
  getRowSize(index: number): number;
  getColumnSize(index: number): number;
  rowOverscan?: number;
  columnOverscan?: number;
};

export type TwoAxisVirtualizer = {
  visibleRows: readonly AxisMeasurement[];
  visibleColumns: readonly AxisMeasurement[];
  totalHeight: number;
  totalWidth: number;
  onScroll(event: ReactUIEvent<HTMLElement>): void;
  getCellRect(rowIndex: number, columnIndex: number): { top: number; left: number; width: number; height: number };
  ensureCellVisible(rowIndex: number, columnIndex: number): void;
};

export function useTwoAxisVirtualizer(options: TwoAxisVirtualizerOptions): TwoAxisVirtualizer;
```

Memoize full row/column measurements from `measureAxis`, derive both visible ranges with `findVisibleRange`, observe container width and height with `ResizeObserver`, and use leading-plus-trailing `requestAnimationFrame` scroll updates. `getCellRect` reads measurements directly. `ensureCellVisible` changes only the axis outside the viewport and honors variable sizes.

- [ ] **Step 4: Replace Grid's private viewport calculations with the hook**

Keep `GridProps`, `GridScrollApi`, selection geometry, merges, hidden/frozen indexes, and rendered class names unchanged. Remove duplicated row/column visibility binary searches only after the hook test is green. Feed foundation-visible indexes into the existing spreadsheet row/cell renderer. Preserve the full measurement arrays for hit testing and overlay geometry.

- [ ] **Step 5: Add spreadsheet parity and wide-column assertions**

In `src/components/Grid.test.tsx`, retain all existing tests and add:

```ts
it("renders bounded row and column windows for a 100000 by 10000 sheet");
it("keeps selection and editor geometry correct after two-axis scrolling");
it("retains a frozen first column outside the ordinary horizontal window");
it("ensureCellVisible reaches a distant row and column");
```

Use the existing `Grid` fixture pattern, set `clientHeight`, `clientWidth`, `scrollTop`, and `scrollLeft` directly, and assert rendered gridcells stay below the calculated visible-row × visible-column bound rather than relying on timing.

- [ ] **Step 6: Run hook, Grid, and App parity suites**

Run:

```bash
corepack pnpm exec vitest run \
  src/react/viewport/useTwoAxisVirtualizer.test.tsx \
  src/components/Grid.test.tsx \
  src/App.test.tsx
corepack pnpm run build
```

Expected: all new and existing tests PASS; no spreadsheet behavior or public `Grid` callback changes.

- [ ] **Step 7: Commit the extracted virtualizer**

```bash
git add src/react/viewport/useTwoAxisVirtualizer.ts src/react/viewport/useTwoAxisVirtualizer.test.tsx src/components/Grid.tsx src/components/Grid.test.tsx
git commit -m "refactor: share two-axis viewport measurement"
```

### Task 7: Extract the shared interaction reducer and accessible grid viewport

**Files:**
- Create: `src/react/viewport/types.ts`
- Create: `src/react/viewport/gridInteraction.ts`
- Create: `src/react/viewport/gridInteraction.test.ts`
- Create: `src/react/viewport/useGridInteraction.ts`
- Create: `src/react/viewport/GridViewport.tsx`
- Create: `src/react/viewport/GridViewport.test.tsx`

**Interfaces:**
- Consumes: Task 6 `useTwoAxisVirtualizer`.
- Produces: renderer-neutral React viewport descriptors, stable-ID selection/navigation, pointer drag, editor lifecycle, clipboard events, focus restoration, accessible grid DOM, and live announcements.

- [ ] **Step 1: Write failing pure interaction-reducer tests**

Create `src/react/viewport/gridInteraction.test.ts`:

```ts
import { describe, expect, it } from "vitest";
import { createGridInteractionState, reduceGridInteraction } from "./gridInteraction";

const rowIds = ["row-1", "row-2", "row-3"] as const;
const columnIds = ["name", "salary", "active"] as const;
const model = { rowIds, columnIds };

describe("grid interaction reducer", () => {
  it("moves and extends selection using stable ids", () => {
    const initial = createGridInteractionState({ rowId: "row-1", columnId: "name" });
    const moved = reduceGridInteraction(initial, { type: "move", rowDelta: 1, columnDelta: 1, extend: false }, model);
    const extended = reduceGridInteraction(moved, { type: "move", rowDelta: 1, columnDelta: 0, extend: true }, model);
    expect(moved.selection).toEqual({
      anchor: { rowId: "row-2", columnId: "salary" },
      focus: { rowId: "row-2", columnId: "salary" }
    });
    expect(extended.selection).toEqual({
      anchor: { rowId: "row-2", columnId: "salary" },
      focus: { rowId: "row-3", columnId: "salary" }
    });
  });

  it("starts, updates, commits, and cancels raw editor state", () => {
    const initial = createGridInteractionState({ rowId: "row-1", columnId: "name" });
    const editing = reduceGridInteraction(initial, { type: "start-edit", rawText: "A" }, model);
    const changed = reduceGridInteraction(editing, { type: "change-edit", rawText: "Ada" }, model);
    const cancelled = reduceGridInteraction(changed, { type: "cancel-edit" }, model);
    expect(changed.editing).toEqual({ rowId: "row-1", columnId: "name", rawText: "Ada" });
    expect(cancelled.editing).toBeNull();
    expect(cancelled.selection).toEqual(initial.selection);
  });
});
```

- [ ] **Step 2: Run reducer tests and verify red**

Run: `corepack pnpm exec vitest run src/react/viewport/gridInteraction.test.ts`

Expected: FAIL because `gridInteraction.ts` is missing.

- [ ] **Step 3: Define viewport descriptors and interaction events**

Create `src/react/viewport/types.ts`:

```ts
import type { CSSProperties, KeyboardEvent as ReactKeyboardEvent, ReactNode } from "react";
import type { TableCellRef, TableSelection } from "../../table/core/types";

export type GridViewportRow = {
  id: string;
  label: string;
  height: number;
  kind: "data" | "group" | "aggregate";
  ariaRowIndex: number;
};
export type GridViewportColumn = { id: string; label: string; width: number; minWidth: number; maxWidth: number; pinned?: "left" | "right" };
export type GridViewportCell = {
  ref: TableCellRef;
  ariaLabel: string;
  displayValue: string;
  editable: boolean;
  invalid: boolean;
  className?: string;
  style?: CSSProperties;
  columnSpan?: number;
  rowSpan?: number;
};
export type GridEditorState = TableCellRef & { rawText: string };
export type GridViewportInteraction =
  | { type: "selection-change"; selection: TableSelection }
  | { type: "edit-start"; cell: TableCellRef; initialRawText: string }
  | { type: "edit-change"; rawText: string }
  | { type: "edit-commit"; cell: TableCellRef; rawText: string; move?: "up" | "down" | "left" | "right" }
  | { type: "edit-cancel" }
  | { type: "copy"; selection: TableSelection }
  | { type: "paste"; text: string }
  | { type: "column-resize"; columnId: string; width: number };

export type GridViewportRenderContext = { row: GridViewportRow; column: GridViewportColumn; cell: GridViewportCell; selected: boolean; active: boolean };
export type GridViewportProps = {
  idPrefix: string;
  ariaLabel: string;
  rows: readonly GridViewportRow[];
  ariaRowCount: number;
  columns: readonly GridViewportColumn[];
  ariaColumnCount: number;
  getCell(rowId: string, columnId: string): GridViewportCell;
  selection: TableSelection | null;
  editing: GridEditorState | null;
  onInteraction(interaction: GridViewportInteraction): void;
  renderCell?(context: GridViewportRenderContext): ReactNode;
  renderEditor?(context: GridViewportRenderContext & { editing: GridEditorState }): ReactNode;
  renderColumnHeader?(column: GridViewportColumn): ReactNode;
  renderRowHeader?(row: GridViewportRow): ReactNode;
  announce?: string;
  onUnhandledKeyDown?(event: ReactKeyboardEvent<HTMLElement>): void;
};
```

Set `ariaRowCount` to the known total plus the header row, or `-1` when the final rendered total is unknown. Each adapter computes the one-based `GridViewportRow.ariaRowIndex`; the viewport renders that explicit value and never substitutes a virtual-window index. Local/workbook rows use their logical position plus 2, offset pagination adds its server offset, and cursor/infinite projections use their stable loaded sequence position while keeping `ariaRowCount=-1` until a known final count arrives. `ariaColumnCount` includes the row-header column when `renderRowHeader` is supplied; data-cell `aria-colindex` values shift by one in that case.

- [ ] **Step 4: Implement the pure reducer and React hook**

Create `gridInteraction.ts` with `GridInteractionState`, `GridInteractionAction`, `createGridInteractionState`, and `reduceGridInteraction`. Clamp movement at edges, retain the selection anchor while extending, navigate by stable ID arrays, and retain raw edit text independently of stored values.

Create `useGridInteraction.ts` with:

```ts
export function useGridInteraction(options: {
  rows: readonly GridViewportRow[];
  columns: readonly GridViewportColumn[];
  selection: TableSelection | null;
  editing: GridEditorState | null;
  onInteraction(interaction: GridViewportInteraction): void;
  ensureCellVisible(rowIndex: number, columnIndex: number): void;
  rootRef: RefObject<HTMLDivElement | null>;
}): {
  onKeyDown(event: React.KeyboardEvent<HTMLDivElement>): void;
  onCellPointerDown(cell: TableCellRef, extend: boolean): void;
  onCellPointerEnter(cell: TableCellRef): void;
  onPointerUp(): void;
  activeDescendantId: string | undefined;
};
```

Handle arrows, Shift+arrows, Home, End, PageUp, PageDown, Tab/Shift+Tab, Enter, Shift+Enter, F2, Escape, printable characters, command-copy, and paste. Leave application-specific command shortcuts to `onUnhandledKeyDown`. Maintain one roving cell target with `tabIndex=0`; when virtualization would unmount it, focus the stable grid root, scroll it into view, then restore the cell on the next layout effect.

- [ ] **Step 5: Write failing accessible viewport integration tests**

Create `src/react/viewport/GridViewport.test.tsx` using three inline rows and columns. Render a stateful harness with `useState` for selection/editing and assert:

```ts
it("renders correct grid, row, header, and cell roles with one-based indexes");
it("uses aria-rowcount minus one for unknown totals and preserves offset row indexes");
it("selects by click and extends by shift-click and drag");
it("navigates by keyboard and keeps exactly one roving tab stop");
it("starts edit with F2 or printable text and commits with Enter or Tab");
it("cancels edit with Escape and restores grid focus");
it("emits plain-text clipboard paste without mutating data itself");
it("renders only the two-axis virtual window");
it("uses instance-prefixed ids for cells and live regions");
it("uses total and page offset for aria indexes instead of virtual indexes");
it("sets aria-rowcount to minus one when the remote total is unknown");
```

Run: `corepack pnpm exec vitest run src/react/viewport/GridViewport.test.tsx`

Expected: FAIL because `GridViewport.tsx` is missing.

- [ ] **Step 6: Implement the shared semantic viewport**

Create `GridViewport.tsx`. The root owns `role="grid"`, `data-viewport-kernel="shared"`, `aria-rowcount={ariaRowCount}`, `aria-colcount={ariaColumnCount}`, `aria-label`, and the focus fallback. ARIA permits `aria-rowcount="-1"` when the total is unknown. Render a virtualized header row at ARIA row index 1, virtualized rows using each descriptor's explicit `ariaRowIndex`, optional row headers, and only the visible cell cross-product on an absolutely positioned canvas sized by Task 6. Never derive accessibility indexes from the current virtual-array index. Set `role`, `aria-colindex`, `aria-selected`, `aria-readonly`, `aria-invalid`, and deterministic encoded IDs on each cell. Render the editor in an overlay positioned from `getCellRect`, not as canonical state inside the viewport. Add one visually hidden `aria-live="polite"` status node and no module-level IDs or listeners.

- [ ] **Step 7: Run interaction and viewport suites**

Run:

```bash
corepack pnpm exec vitest run \
  src/react/viewport/gridInteraction.test.ts \
  src/react/viewport/GridViewport.test.tsx \
  src/react/viewport/useTwoAxisVirtualizer.test.tsx
corepack pnpm run build
```

Expected: all tests and strict build PASS without accessibility-role warnings.

- [ ] **Step 8: Commit the interaction kernel**

```bash
git add src/react/viewport/types.ts src/react/viewport/gridInteraction.ts src/react/viewport/gridInteraction.test.ts src/react/viewport/useGridInteraction.ts src/react/viewport/GridViewport.tsx src/react/viewport/GridViewport.test.tsx
git commit -m "feat: add shared accessible grid interaction"
```

### Task 8: Adapt the spreadsheet Grid to the shared viewport

**Files:**
- Create: `src/components/grid/SpreadsheetCell.tsx`
- Create: `src/components/grid/SpreadsheetGridHeaders.tsx`
- Modify: `src/components/Grid.tsx`
- Modify: `src/components/Grid.test.tsx`
- Modify: `src/App.test.tsx`
- Modify: `src/App.css`

**Interfaces:**
- Consumes: `GridViewport`, `GridViewportInteraction`, and the existing foundation-controlled workbook commands.
- Produces: a spreadsheet adapter retaining the existing `GridProps`, `GridScrollApi`, A1 semantics, workbook editor behavior, and visual parity.

- [ ] **Step 1: Strengthen spreadsheet parity tests before moving DOM ownership**

Add these focused tests to the existing Grid/App suites before production changes:

```ts
it("renders the spreadsheet through the shared viewport kernel");
it("keeps formula suggestions and validation choices inside the active editor overlay");
it("keeps AutoFilter menus and cell context interactions source-specific");
it("preserves merged-cell spans and selection geometry in a virtual window");
it("preserves fill-handle drag and four-direction fill targets");
it("preserves frozen row and column selection styling");
it("returns focus to the spreadsheet grid after commit and cancel");
```

Run:

```bash
corepack pnpm exec vitest run src/components/Grid.test.tsx src/App.test.tsx
```

For the first test, assert the semantic grid has `data-viewport-kernel="shared"`; Task 7's `GridViewport` owns that diagnostic attribute. Expected: this new adapter assertion FAILS because Spreadsheet still owns its private semantic grid, while all pre-existing behavior assertions PASS. If any behavior characterization fails independently, fix it in the foundation plan before continuing this task.

- [ ] **Step 2: Move source-specific cell content into a renderer slot**

Create `src/components/grid/SpreadsheetCell.tsx` by moving, without semantic changes, the existing display formatting, hyperlink, comment, conditional formatting, data bar, validation dropdown, AutoFilter menu, merge coverage, read-only, formula suggestion, and frozen-pane rendering from `Grid.tsx`. Export one component:

```ts
export function SpreadsheetCell(props: SpreadsheetCellProps): React.ReactNode;
```

`SpreadsheetCellProps` must receive the already-resolved address, value, format, validation, rule values, merge info, callbacks, and `GridViewportRenderContext`; it must not read App state or mutate a workbook.

- [ ] **Step 3: Move spreadsheet header chrome into viewport slots**

Create `SpreadsheetGridHeaders.tsx` exporting `SpreadsheetColumnHeader` and `SpreadsheetRowHeader`. Preserve current labels, selection states, resize/auto-fit callbacks, hidden indexes, and sticky/frozen CSS classes. The shared viewport owns semantic header containers; these slots render only source-specific contents and controls.

- [ ] **Step 4: Map stable viewport ids to sheet coordinates**

In `Grid.tsx`, use `row:${rowIndex}` and `column:${columnIndex}` as view-only IDs, with explicit conversion functions local to the adapter:

```ts
function sheetRowId(row: number): string { return `row:${row}`; }
function sheetColumnId(column: number): string { return `column:${column}`; }
function parseSheetIndex(id: string, prefix: "row:" | "column:"): number {
  if (!id.startsWith(prefix) || !Number.isInteger(Number(id.slice(prefix.length)))) {
    throw new Error(`Invalid spreadsheet viewport id: ${id}`);
  }
  return Number(id.slice(prefix.length));
}
```

Map `GridViewportInteraction` back to the existing `onSelectionChange`, edit, paste, resize, and key-command callbacks. Continue deriving values from `SheetModel` and `FormulaEngine`; never create record rows or a `RecordTableSession` for the spreadsheet.

Spreadsheet viewport rows use `ariaRowIndex: row + 2`, `ariaRowCount: sheet.rowCount + 1`, and `ariaColumnCount: sheet.columnCount + 1` when headers are shown; the column header is ARIA row 1 and the row header occupies ARIA column 1. Hidden rows/columns retain their logical index so accessibility positions do not collapse to the virtual/rendered order.

- [ ] **Step 5: Remove duplicate viewport DOM only after adapter tests pass**

Replace Grid's private scroll canvas, semantic cell wrappers, pointer selection, editor positioning, and focus lifecycle with `GridViewport`. Retain spreadsheet-only fill/copy/merge overlays through viewport render slots. Keep `.grid-scroll`, `.spreadsheet-grid`, `.cell`, `.column-header`, and `.row-header` compatibility classes during this plan so App visual tests do not change for unrelated reasons.

- [ ] **Step 6: Run full spreadsheet parity**

Run:

```bash
corepack pnpm exec vitest run src/react/viewport src/components/Grid.test.tsx src/App.test.tsx
corepack pnpm run build
corepack pnpm run test:e2e
```

Expected: every command exits 0; existing spreadsheet keyboard, edit, filter, formatting, formula, fill, resize, freeze, and browser tests remain green.

- [ ] **Step 7: Commit the spreadsheet adapter**

```bash
git add src/components/grid/SpreadsheetCell.tsx src/components/grid/SpreadsheetGridHeaders.tsx src/components/Grid.tsx src/components/Grid.test.tsx src/App.test.tsx src/App.css
git commit -m "refactor: adapt spreadsheet to shared grid viewport"
```

### Task 9: Add the typed React session hook and mutually exclusive public props

**Files:**
- Create: `src/react/tableTypes.ts`
- Create: `src/react/useTableSession.ts`
- Create: `src/react/useTableSession.test.tsx`
- Create: `src/react/useTableSnapshot.ts`
- Create: `src/react/TableSessionContext.tsx`
- Create: `src/react/TableSessionContext.test.tsx`
- Create: `src/react/DataTable.types.test.tsx`
- Create: `src/react/index.ts`

**Interfaces:**
- Consumes: `createLocalRecordTableSession` and core generic column/session types.
- Produces: React-specialized render/editor/context types, `ColumnDef`, `DataTableProps`, `DataTableHandle`, `useTableSession`, `useTableSnapshot`, `TableSessionProvider`, and `useTableSessionContext` without destroying host-owned sessions.

- [ ] **Step 1: Write failing public generic and exclusivity tests**

Create `src/react/DataTable.types.test.tsx`:

```tsx
import { expectTypeOf, test } from "vitest";
import type { ComponentProps } from "react";
import type { ExportArtifact, TableSession } from "../table/core/types";
import type { ColumnDef, DataTableHandle, DataTableProps } from "./tableTypes";

type Employee = { id: string; name: string; salary: number };

test("public local props infer row and renderer values", () => {
  const columns: readonly ColumnDef<Employee>[] = [{
    id: "name",
    header: "Name",
    accessor: (row) => row.name,
    update: (row, value) => ({ ...row, name: String(value) }),
    cell: ({ row }) => row.original.name
  }];
  const props: DataTableProps<Employee> = {
    rows: [{ id: "1", name: "Ada", salary: 100 }],
    columns,
    getRowId: (row) => row.id,
    onRowsChange: (updater) => updater([{ id: "1", name: "Ada", salary: 100 }])
  };
  expectTypeOf(props.getRowId).toBeFunction();
  expectTypeOf<DataTableHandle["scrollToRow"]>().toBeFunction();
  expectTypeOf<Awaited<ReturnType<TableSession<Employee>["export"]>>>().toEqualTypeOf<ExportArtifact>();
  expectTypeOf<Awaited<ReturnType<DataTableHandle["export"]>>>().toEqualTypeOf<Blob>();
  expectTypeOf<ComponentProps<"div">["className"]>().toEqualTypeOf<string | undefined>();
});

test("simple and session modes cannot be mixed", () => {
  const columns: readonly ColumnDef<Employee>[] = [];
  // @ts-expect-error session mode may not also provide rows
  const invalid: DataTableProps<Employee> = { session: {} as never, rows: [], columns, getRowId: (row) => row.id };
  expectTypeOf(invalid).toMatchTypeOf<DataTableProps<Employee>>();
});
```

- [ ] **Step 2: Write failing ownership and StrictMode hook tests**

Create `src/react/useTableSession.test.tsx` using `renderHook`, `StrictMode`, one inline employee source, and `vi.spyOn(session, "destroy")`. Cover:

```ts
it("retains one owned session and synchronizes controlled options on rerender");
it("survives StrictMode simulated effect cleanup");
it("destroys an owned session after real unmount");
it("never destroys a session supplied to DataTable");
it("subscribes through useSyncExternalStore without tearing");
```

Run:

```bash
corepack pnpm exec vitest run src/react/DataTable.types.test.tsx src/react/useTableSession.test.tsx
corepack pnpm run build
```

Expected: FAIL because React table types and hooks do not exist.

- [ ] **Step 3: Define React-specialized columns, props, render contexts, and handle**

Create `src/react/tableTypes.ts`:

```ts
import type { ComponentType, CSSProperties, ReactNode } from "react";
import type { CommandResult } from "../core/commands/types";
import type { QueryRow } from "../table/core/query";
import type {
  ChangeContext,
  ColumnDef as CoreColumnDef,
  ExportOptions,
  RowUpdater,
  TableCellSnapshot,
  TableDiagnosticEvent,
  TableIntent,
  TableMetadataDocument,
  TableMetadataUpdater,
  TableSelection,
  TableSession,
  TableStateUpdater,
  TableViewState
} from "../table/core/types";
import type { TableFeatureConfiguration } from "../table/core/capabilities";

export type CellRenderContext<TRow, TValue> = {
  row: Extract<QueryRow<TRow>, { kind: "data" }>;
  column: CoreColumnDef<TRow, TValue>;
  cell: TableCellSnapshot;
};
export type HeaderRenderContext<TRow, TValue> = { column: CoreColumnDef<TRow, TValue> };
export type CellEditorProps<TRow, TValue> = CellRenderContext<TRow, TValue> & {
  rawText: string;
  onChange(rawText: string): void;
  onCommit(): void;
  onCancel(): void;
};

export type ColumnDef<TRow, TValue = unknown> = CoreColumnDef<
  TRow,
  TValue,
  (context: CellRenderContext<TRow, TValue>) => ReactNode,
  (context: HeaderRenderContext<TRow, TValue>) => ReactNode,
  ComponentType<CellEditorProps<TRow, TValue>>
>;

export type DataTablePresentationProps = {
  "aria-label"?: string;
  className?: string;
  style?: CSSProperties;
  rowSelection?: "none" | "single" | "multiple";
  inlineFilters?: boolean;
  layout?: "fixed" | "fluid" | "ratio";
  rowHeight?: number | "auto";
  noDataMessage?: ReactNode;
  onRowSelectionChange?(rowIds: readonly string[]): void;
  onDiagnostic?(event: TableDiagnosticEvent): void;
};

export type DataTableProps<TRow> =
  | (DataTablePresentationProps & {
      rows: readonly TRow[];
      columns: readonly ColumnDef<TRow>[];
      getRowId(row: TRow): string;
      getSubRows?(row: TRow): readonly TRow[] | undefined;
      resetKey?: string | number;
      onRowsChange?(updater: RowUpdater<TRow>, context: ChangeContext): void;
      features?: TableFeatureConfiguration;
      state?: Partial<TableViewState>;
      defaultState?: Partial<TableViewState>;
      onStateChange?(updater: TableStateUpdater, context: ChangeContext): void;
      document?: TableMetadataDocument;
      defaultDocument?: TableMetadataDocument;
      onDocumentChange?(updater: TableMetadataUpdater, context: ChangeContext): void;
      session?: never;
    })
  | (DataTablePresentationProps & {
      session: TableSession<TRow, ColumnDef<TRow>>;
      rows?: never;
      columns?: never;
      getRowId?: never;
      getSubRows?: never;
      resetKey?: never;
      onRowsChange?: never;
      features?: never;
      state?: never;
      defaultState?: never;
      onStateChange?: never;
      document?: never;
      defaultDocument?: never;
      onDocumentChange?: never;
    });

export interface DataTableHandle {
  focus(): void;
  refresh(): Promise<void>;
  dispatch(intent: TableIntent): Promise<CommandResult>;
  undo(): Promise<CommandResult>;
  redo(): Promise<CommandResult>;
  scrollToRow(rowId: string): void;
  getSelection(): TableSelection | null;
  export(options: ExportOptions): Promise<Blob>;
}
```

`DataTableHandle` is intentionally part of the React/browser entrypoint, so its convenience `Blob` return is permitted. It must delegate to `TableSession.export`, copy the returned `ExportArtifact.bytes`, and create the `Blob` only in React code; no headless session or source contract may mention `Blob`.

- [ ] **Step 4: Implement the local session hook and external-store subscription**

Create `useTableSession.ts`:

```ts
export function useTableSession<TRow>(
  options: LocalRecordTableSessionOptions<TRow, ColumnDef<TRow>>
): RecordTableSession<TRow, ColumnDef<TRow>>;
```

Create the session once in a ref, call idempotent `updateOptions(options)` during render so controlled rows are never one render stale, and use an owned-session cleanup token: schedule `destroy()` in `queueMicrotask` during cleanup and cancel that token when StrictMode immediately re-runs the effect. Do not apply this ownership lifecycle to a `session` prop.

Create `useTableSnapshot.ts`:

```ts
export function useTableSnapshot<TRow, TColumn>(
  session: TableSession<TRow, TColumn>
): TableViewSnapshot<TRow, TColumn> {
  return useSyncExternalStore(session.subscribe.bind(session), session.getSnapshot.bind(session), session.getSnapshot.bind(session));
}
```

Create `TableSessionContext.tsx` with `TableSessionProvider<TRow>({ session, children })` and `useTableSessionContext<TRow>()`. The provider only supplies a host-owned session and never destroys it. The hook throws the exact error `useTableSessionContext must be used inside TableSessionProvider` when absent. `TableSessionContext.test.tsx` covers the error, nested-provider isolation, snapshot updates through `useTableSnapshot`, StrictMode, and no destroy on provider unmount.

- [ ] **Step 5: Verify type inference, ownership, and StrictMode**

Run:

```bash
corepack pnpm exec vitest run src/react/DataTable.types.test.tsx src/react/useTableSession.test.tsx src/react/TableSessionContext.test.tsx
corepack pnpm run build
```

Expected: tests and strict build PASS, including the intentional `@ts-expect-error`.

- [ ] **Step 6: Export and commit the React session API**

Create `src/react/index.ts` with the foundation's `Spreadsheet`/`useWorkbookSession` exports plus explicit table types/hooks; do not export unfinished `DataTable` yet.

```bash
git add src/react/tableTypes.ts src/react/useTableSession.ts src/react/useTableSession.test.tsx src/react/useTableSnapshot.ts src/react/TableSessionContext.tsx src/react/TableSessionContext.test.tsx src/react/DataTable.types.test.tsx src/react/index.ts
git commit -m "feat: add typed React table session API"
```

### Task 10: Implement simple and session-backed local DataTable

**Files:**
- Create: `src/react/DataTable.tsx`
- Create: `src/react/DataTable.test.tsx`
- Create: `src/react/DataTableCell.tsx`
- Create: `src/react/DataTableToolbar.tsx`
- Create: `src/react/DataTableColumnMenu.tsx`
- Create: `src/react/DataTableExtensionBoundary.tsx`
- Create: `src/react/exportArtifact.ts`
- Create: `src/react/exportArtifact.test.ts`
- Modify: `src/react/index.ts`
- Modify: `src/index.ts`

**Interfaces:**
- Consumes: Tasks 7-9 shared viewport, local session hook, and public React types.
- Produces: generic `DataTable`, compact local toolbar/column menu/editor, imperative handle, CSV/XLSX record export, and current source-level exports.

- [ ] **Step 1: Write failing simple/session component tests**

Create `src/react/DataTable.test.tsx` with one `Employee` fixture and accessor columns. Write named tests with Testing Library/user-event for:

```ts
it("renders simple local rows and typed headers");
it("edits a simple controlled row through one onRowsChange updater");
it("retains uncontrolled simple edits across rerenders until resetKey changes");
it("resets uncontrolled simple rows and history when resetKey changes");
it("renders a prebuilt session without creating or destroying it");
it("uses custom cell, header, and editor renderers");
it("isolates a throwing custom renderer and reports safe diagnostics without row values");
it("keeps the editor focused when validation rejects a commit");
it("sorts, filters, groups, aggregates, and paginates the complete local dataset");
it("copies and atomically pastes a tab/newline matrix");
it("undoes and redoes edits from toolbar and handle");
it("edits format, validation, comment, read-only, and formula metadata from quick tools");
it("resizes, hides, and pins columns by stable id");
it("reorders columns visibly by pointer drag and keyboard controls");
it("renders checkbox row selection with single and multiple modes");
it("expands and collapses tree rows by stable id");
it("renders an inline typed filter row without replacing column menus");
it("runs custom header actions and reports safe host errors");
it("measures wrapped auto-height rows without breaking virtualization");
it("supports fixed fluid and ratio layouts plus a custom empty message");
it("visibly disables a formula action when no formula service exists");
it("isolates two simultaneous DataTable instances");
```

The controlled-edit assertion must extract the callback updater and apply it to the original rows, as in Task 4. The paste test must include one invalid cell and prove no cell changed.

- [ ] **Step 2: Run DataTable tests and verify red**

Run: `corepack pnpm exec vitest run src/react/DataTable.test.tsx`

Expected: FAIL because `DataTable.tsx` and its child components are missing.

- [ ] **Step 3: Implement the generic forwarded-ref component**

Use the standard generic forward-ref cast so JSX infers `TRow`:

```ts
function DataTableInner<TRow>(
  props: DataTableProps<TRow>,
  forwardedRef: ForwardedRef<DataTableHandle>
): ReactElement {
  return props.session !== undefined
    ? <SessionDataTable props={props} forwardedRef={forwardedRef} />
    : <OwnedLocalDataTable props={props} forwardedRef={forwardedRef} />;
}

export const DataTable = forwardRef(DataTableInner) as <TRow>(
  props: DataTableProps<TRow> & RefAttributes<DataTableHandle>
) => ReactElement;
```

`OwnedLocalDataTable` calls `useTableSession` with `{ source: { kind: "local", rows, getRowId, onRowsChange, resetKey }, columns }`; `SessionDataTable` uses the supplied session directly and never destroys it. Both render a shared internal `DataTableSurface`. This component split prevents conditional hook calls if a host switches modes. Convert each `QueryRow` into a viewport row and call only `snapshot.getCell(rowId, columnId)` for cells. For an ungrouped known-total offset page, set each `GridViewportRow.ariaRowIndex` to `offset + itemIndex + 2` (row 1 is the column-header row) and `ariaRowCount` to `total.value + 1`. Use `ariaRowCount = -1` for unknown totals or server group projections whose final rendered group count is unknown; locally completed group projections use `snapshot.rows.length + 1`. DataTable renders numbered row headers, so pass `ariaColumnCount = snapshot.columns.length + 1` and offset data-cell `aria-colindex` by one. Group/aggregate rows render their renderer-neutral aggregate fields and never pretend to be `TRow` data rows.

When `getSubRows` is supplied, include it in the local source shown above, render an accessible disclosure button in the first visible data column, and dispatch `set-row-expanded` by stable ID. When row selection is enabled, render a leading checkbox column (header checkbox for multi-select), keep checkbox selection separate from cell selection, and dispatch `set-row-selection`; invoke `onRowSelectionChange` only after committed view state.

- [ ] **Step 4: Implement interaction-to-command mapping and editors**

`DataTableCell.tsx` renders default display text or the typed column renderer. Its editor uses the custom editor component when present, otherwise a text input or boolean/list control based on column data type/metadata. Map viewport interactions as follows:

```ts
"selection-change" -> { type: "set-selection", selection }
"edit-commit" -> { type: "edit-cells", edits: [{ rowId, columnId, rawText }] }
"column-resize" -> { type: "resize-column", columnId, width }
"paste" -> parse a rectangular tab/newline matrix, then dispatch one "edit-cells" batch
```

On validation rejection, keep editor state and focus and announce the first issue. On commit, move according to Enter/Tab. Never update row objects directly from React.

Wrap every custom header, cell, and editor slot in `DataTableExtensionBoundary.tsx`, a React error boundary keyed by row/column/slot. On failure, render an accessible `"Cell renderer unavailable"` fallback, leave the rest of the grid operable, and call `onDiagnostic({ category: "extension", metadata: { columnId, slot } })`. Never include row objects, cell values, formulas, or edit text in diagnostic metadata.

- [ ] **Step 5: Implement compact toolbar and column menus through capabilities**

`DataTableToolbar.tsx` renders undo, redo, refresh, export, current selection summary, and a toggleable quick-tools drawer. Quick tools edit selected cells' number format, text/fill emphasis, validation, comment, read-only flag, and formula metadata by dispatching one `update-cell-metadata` command; source capability/permission determines whether each control is enabled. `DataTableColumnMenu.tsx` renders sort ascending/descending, clear sort, typed filter controls, group/ungroup, aggregate selection, hide/show, and pin controls. Intersect `snapshot.operationStates` with the column's `sortable`, `filterable`, `groupable`, and `aggregatable` settings and the current cell permission. A disabled command must show its reason via accessible description and must not fall back to loaded/visible rows.

When `inlineFilters` is true, render one typed filter control per visible filterable column directly below headers and compose it with the current filter AST. Render `ColumnDef.headerActions` after built-in menu items, invoke them through `safeInvokeTableExtension`, and emit sanitized diagnostics. Add a visible column-order drag handle plus keyboard move-left/move-right actions that dispatch `set-column-order` and announce the new position. `layout` maps to fixed-pixel, fluid-content, or ratio-based column sizing without viewport assumptions. `rowHeight="auto"` measures wrapped content with ResizeObserver and updates the measured axis while preserving scroll anchoring; a numeric value remains fixed-height. Render `noDataMessage` only when the current projection has no items.

- [ ] **Step 6: Write failing browser-artifact conversion tests before implementation**

Create `src/react/exportArtifact.test.ts` with deterministic CSV and XLSX `ExportArtifact` fixtures plus a stub `TableSession`. Assert:

```ts
it("copies artifact bytes and media type into a browser Blob");
it("does not let later artifact byte mutation alter the returned Blob");
it("delegates exact export options once through the DataTable handle");
it("uses the artifact file name for toolbar download and revokes its object URL");
```

The stub session returns `ExportArtifact`; the handle assertion must read `await blob.arrayBuffer()` and compare exact bytes. The toolbar test spies on `URL.createObjectURL`/`URL.revokeObjectURL` but never puts object-URL behavior into the session.

Run: `corepack pnpm exec vitest run src/react/exportArtifact.test.ts`

Expected: FAIL because `exportArtifact.ts` and the handle conversion are missing.

- [ ] **Step 7: Implement the React-only Blob conversion**

Create:

```ts
export function exportArtifactToBlob(artifact: ExportArtifact): Blob;
```

Validate the artifact defensively, clone with `Uint8Array.from(artifact.bytes)`, and construct the `Blob` from that clone's owned `ArrayBuffer` with `artifact.mediaType`. Do not reserialize rows or inspect session internals here: local, remote, and workbook sessions already own export scope and byte generation. The toolbar calls `session.export(options)`, converts the returned artifact, uses `artifact.fileName` as the download name, and always revokes its temporary object URL. Reject malformed artifacts rather than producing an empty or incorrectly typed download.

- [ ] **Step 8: Expose the imperative handle**

Use `useImperativeHandle` to implement the exact Task 9 handle. `scrollToRow` resolves `snapshot.getRowIndex(rowId)`, throws `Unknown row id: <id>` for `-1`, and delegates to viewport `ensureCellVisible`. `getSelection` reads the current snapshot. Other methods delegate to the session. The handle's `export` awaits `session.export(options)` and returns `exportArtifactToBlob(artifact)`; this conversion is covered by both the helper and forwarded-ref contract tests.

- [ ] **Step 9: Run component, exporter, viewport, and type suites**

Run:

```bash
corepack pnpm exec vitest run \
  src/react/DataTable.test.tsx \
  src/react/exportArtifact.test.ts \
  src/react/DataTable.types.test.tsx \
  src/react/viewport
corepack pnpm run build
```

Expected: all tests and strict build PASS without console errors, act warnings, or unhandled promises.

- [ ] **Step 10: Export and commit the local React component**

Export `DataTable`, `DataTableHandle`, React `ColumnDef`, hooks, and session helpers explicitly from `src/react/index.ts`. Add source-level named exports to `src/index.ts`; the packaging plan later replaces the temporary source entry with final subpath exports.

```bash
git add src/react/DataTable.tsx src/react/DataTable.test.tsx src/react/DataTableCell.tsx src/react/DataTableToolbar.tsx src/react/DataTableColumnMenu.tsx src/react/DataTableExtensionBoundary.tsx src/react/exportArtifact.ts src/react/exportArtifact.test.ts src/react/index.ts src/index.ts
git commit -m "feat: add embeddable local React DataTable"
```

### Task 11: Add the approved adaptive neutral theme without host leakage

**Files:**
- Create: `src/styles/data-table.css`
- Create: `src/react/DataTable.styles.test.tsx`
- Create: `src/demo/DataTableDemo.tsx`
- Modify: `src/main.tsx`

**Interfaces:**
- Consumes: Task 10 `DataTable` class structure.
- Produces: `.js-spreadsheet-root.js-spreadsheet-data-table`, `data-js-spreadsheet-root="data-table"`, exported prefixed tokens, container-adaptive compact layout, reduced-motion rules, and a network-reachable demo route.

- [ ] **Step 1: Write a failing root/scope test before adding CSS**

Create `src/react/DataTable.styles.test.tsx`:

```tsx
import { readFileSync } from "node:fs";
import { render, screen } from "@testing-library/react";
import { describe, expect, it } from "vitest";
import { DataTable } from "./DataTable";

type Row = { id: string; name: string };

describe("DataTable style boundary", () => {
  it("renders the exact public style root", () => {
    render(<DataTable
      aria-label="Employees"
      rows={[{ id: "1", name: "Ada" }]}
      columns={[{
        id: "name",
        header: "Name",
        accessor: (row: Row) => row.name,
        update: (row, value) => ({ ...row, name: String(value) })
      }]}
      getRowId={(row) => row.id}
    />);
    const root = screen.getByRole("grid", { name: "Employees" }).closest(".js-spreadsheet-root");
    expect(root).toHaveClass("js-spreadsheet-data-table");
    expect(root).toHaveAttribute("data-js-spreadsheet-root", "data-table");
  });

  it("contains no unscoped element or universal selectors", () => {
    const css = readFileSync(new URL("../styles/data-table.css", import.meta.url), "utf8");
    const withoutComments = css.replace(/\/\*[\s\S]*?\*\//g, "");
    expect(withoutComments).not.toMatch(
      /(?:^|[{}])\s*(?:body|html|:root|button|input|select|textarea|\*)(?=\s|,|\{|\.|:|#|\[)/m
    );
  });
});
```

- [ ] **Step 2: Run style tests and verify red**

Run: `corepack pnpm exec vitest run src/react/DataTable.styles.test.tsx`

Expected: FAIL because the exact root attributes and `src/styles/data-table.css` do not exist.

- [ ] **Step 3: Add the exact root contract to DataTable**

In `DataTable.tsx`, wrap the component with:

```tsx
<div
  className={["js-spreadsheet-root", "js-spreadsheet-data-table", props.className].filter(Boolean).join(" ")}
  data-js-spreadsheet-root="data-table"
  style={props.style}
>
  {/* toolbar, edit bar, viewport, quick tools, status */}
</div>
```

Do not forward `className` to the semantic grid itself. Instance IDs, session state, and accessibility attributes remain internal.

- [ ] **Step 4: Create the scoped token and layout stylesheet**

Create `src/styles/data-table.css`. Every ordinary selector begins with `.js-spreadsheet-root.js-spreadsheet-data-table` or is nested inside an at-rule and still begins with that root. Start with these exact public variables:

```css
.js-spreadsheet-root.js-spreadsheet-data-table {
  --js-spreadsheet-accent: #22795d;
  --js-spreadsheet-accent-strong: #17634a;
  --js-spreadsheet-accent-soft: #eaf7f2;
  --js-spreadsheet-border: #cbd5e1;
  --js-spreadsheet-border-strong: #94a3b8;
  --js-spreadsheet-chrome: #f8fafc;
  --js-spreadsheet-chrome-strong: #eef3f7;
  --js-spreadsheet-text: #172033;
  --js-spreadsheet-muted: #5f6e82;
  --js-spreadsheet-surface: #ffffff;
  container: js-spreadsheet-table / inline-size;
  position: relative;
  display: grid;
  grid-template-rows: auto auto minmax(0, 1fr) auto;
  min-width: 0;
  min-height: 220px;
  height: 100%;
  color: var(--js-spreadsheet-text);
  background: var(--js-spreadsheet-surface);
  border: 1px solid var(--js-spreadsheet-border-strong);
  border-radius: 8px;
  overflow: hidden;
  font-family: Inter, ui-sans-serif, system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
}

.js-spreadsheet-root.js-spreadsheet-data-table,
.js-spreadsheet-root.js-spreadsheet-data-table *,
.js-spreadsheet-root.js-spreadsheet-data-table *::before,
.js-spreadsheet-root.js-spreadsheet-data-table *::after {
  box-sizing: border-box;
}
```

Add scoped toolbar, edit bar, column-menu, viewport, cell, editor, selection, validation, conflict, status, and visually-hidden classes. Use muted green only for active selection/focus/primary actions, white work surfaces, slate chrome, and hairline borders. Use `:focus-visible` with at least a 2px accent outline. Use `@container js-spreadsheet-table (max-width: 640px)` to collapse text labels into accessible icon buttons and move quick tools into a drawer. Do not use `100vh`, global viewport width, or unprefixed custom properties.

- [ ] **Step 5: Add reduced-motion and forced-colors handling**

Include:

```css
@media (prefers-reduced-motion: reduce) {
  .js-spreadsheet-root.js-spreadsheet-data-table *,
  .js-spreadsheet-root.js-spreadsheet-data-table *::before,
  .js-spreadsheet-root.js-spreadsheet-data-table *::after {
    scroll-behavior: auto !important;
    transition-duration: 0.01ms !important;
    animation-duration: 0.01ms !important;
    animation-iteration-count: 1 !important;
  }
}

@media (forced-colors: active) {
  .js-spreadsheet-root.js-spreadsheet-data-table .js-spreadsheet-data-table__active-cell {
    outline: 2px solid Highlight;
  }
}
```

- [ ] **Step 6: Add an isolated demo route without loading spreadsheet global CSS**

Create `src/demo/DataTableDemo.tsx` with a controlled employee tree, typed columns, checkbox row selection, inline filters, a custom header action, visible column reorder, auto row height, sort/filter/edit examples, and a second small table to expose instance-isolation defects. Modify `src/main.tsx` so `/datatable` dynamically imports `DataTableDemo` and `./styles/data-table.css`; every other path dynamically imports the existing App and `./App.css`. This separation is required for the later host-leak test—do not load `App.css` on `/datatable`.

For manual review, document the command in a source comment only:

```bash
corepack pnpm exec vite --host 0.0.0.0 --port 5173
```

Use the machine IP and port when sharing a link; do not give the user a localhost URL.

- [ ] **Step 7: Run style, component, and build tests**

Run:

```bash
corepack pnpm exec vitest run src/react/DataTable.styles.test.tsx src/react/DataTable.test.tsx
corepack pnpm run build
```

Expected: PASS; the production build contains separate DataTable and App CSS chunks, and the style test finds no unscoped selector.

- [ ] **Step 8: Commit the adaptive theme**

```bash
git add src/styles/data-table.css src/react/DataTable.styles.test.tsx src/demo/DataTableDemo.tsx src/main.tsx src/react/DataTable.tsx
git commit -m "style: add scoped adaptive data table theme"
```

### Task 12: Enforce accessibility and keyboard behavior in unit and browser tests

**Files:**
- Create: `src/react/DataTable.a11y.test.tsx`
- Create: `tests/datatable.spec.ts`
- Modify: `package.json`
- Modify: `pnpm-lock.yaml`

**Interfaces:**
- Consumes: the themed DataTable demo.
- Produces: automated axe coverage, keyboard-only workflows, focus retention, announcement assertions, and host-style isolation in Chromium. The final packaging plan expands the same spec to Firefox and WebKit.

- [ ] **Step 1: Install accessibility dependencies before importing them**

Run:

```bash
corepack pnpm add -D axe-core @axe-core/playwright
```

Expected: both packages appear in `devDependencies`; lockfile changes successfully under Node 22.13+.

- [ ] **Step 2: Write the failing unit accessibility suite**

Create `src/react/DataTable.a11y.test.tsx`:

```tsx
import axe from "axe-core";
import { fireEvent, render, screen } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import { describe, expect, it } from "vitest";
import { DataTable } from "./DataTable";

type Row = { id: string; name: string; salary: number };
const rows: readonly Row[] = [
  { id: "1", name: "Ada", salary: 100 },
  { id: "2", name: "Grace", salary: 200 }
];
const columns = [
  { id: "name", header: "Name", accessor: (row: Row) => row.name, update: (row: Row, value: unknown) => ({ ...row, name: String(value) }) },
  { id: "salary", header: "Salary", dataType: "number" as const, accessor: (row: Row) => row.salary, update: (row: Row, value: unknown) => ({ ...row, salary: Number(value) }) }
];

describe("DataTable accessibility", () => {
  it("has no automated axe violations", async () => {
    const { container } = render(<DataTable aria-label="Employees" rows={rows} columns={columns} getRowId={(row) => row.id} />);
    const result = await axe.run(container);
    expect(result.violations).toEqual([]);
  });

  it("supports a keyboard-only edit and announces the result", async () => {
    const user = userEvent.setup();
    render(<DataTable aria-label="Employees" rows={rows} columns={columns} getRowId={(row) => row.id} />);
    const first = screen.getByRole("gridcell", { name: /Ada/ });
    first.focus();
    await user.keyboard("{F2}Augusta{Enter}");
    expect(screen.getByRole("status")).toHaveTextContent(/Updated Name for row 1/);
  });

  it("retains focus through virtualization and sort", async () => {
    render(<DataTable aria-label="Employees" rows={rows} columns={columns} getRowId={(row) => row.id} />);
    const first = screen.getByRole("gridcell", { name: /Ada/ });
    first.focus();
    fireEvent.keyDown(first, { key: "ArrowDown" });
    expect(screen.getByRole("gridcell", { name: /Grace/ })).toHaveFocus();
  });
});
```

- [ ] **Step 3: Run the accessibility suite and verify the intended failures**

Run: `corepack pnpm exec vitest run src/react/DataTable.a11y.test.tsx`

Expected: at least the missing/incorrect announcement or focus assertion FAILS. If every test passes immediately, strengthen the assertions so the suite proves new behavior rather than merely exercising existing markup.

- [ ] **Step 4: Complete roles, names, focus, and announcements**

Ensure the implementation provides:

- Correct `grid`, `row`, `columnheader`, optional `rowheader`, and `gridcell` roles.
- `aria-rowcount`/`aria-colcount` for the logical dataset and one-based indexes on virtual children.
- Header association through deterministic IDs and `aria-labelledby`.
- Exactly one roving cell tab stop, with root fallback while the active cell is virtualized away.
- Visible focus indication and editor labels containing row and column names.
- Live announcements for selection, edit success/rejection, validation, sort, filter, grouping, pagination, loading, pending mutation, correction, and conflict. Local mode exercises the applicable subset; remote mode later populates the shared pending/conflict fields.
- Focus retention by stable row/column ID after sort/filter/refresh, falling back to the nearest visible cell only when the selected ID no longer exists.
- Checkbox row selection with labelled row/header checkboxes, mixed/select-all state, and announcements.
- Tree disclosure buttons with `aria-expanded`, stable `aria-level`, keyboard left/right expand/collapse, and retained parent context.
- Inline filter controls with column-specific accessible names and a keyboard path back to the grid.
- Column reorder controls operable without drag, with position announcements and stable focus.
- Auto-height rows that retain logical ARIA indexes after ResizeObserver measurement.
- No accessibility semantics inside presentation-only spacer/canvas elements.

- [ ] **Step 5: Write browser axe, keyboard, focus, and host-leak tests**

Create `tests/datatable.spec.ts`:

```ts
import AxeBuilder from "@axe-core/playwright";
import { expect, test } from "@playwright/test";

test.beforeEach(async ({ context, page }) => {
  await context.grantPermissions(["clipboard-read", "clipboard-write"]);
  await page.goto("/datatable");
});

test("passes axe and exposes virtual grid semantics", async ({ page }) => {
  const results = await new AxeBuilder({ page }).include(".js-spreadsheet-data-table").analyze();
  expect(results.violations).toEqual([]);
  const grid = page.getByRole("grid", { name: "Employees" });
  await expect(grid).toHaveAttribute("aria-rowcount", "3");
  await expect(grid).toHaveAttribute("aria-colcount", "3");
  await expect(grid.getByRole("columnheader")).toHaveCount(2);
});

test("completes edit, sort, filter, copy, paste, undo, and redo by keyboard", async ({ page }) => {
  const employeesTable = page.locator(".js-spreadsheet-data-table").filter({
    has: page.getByRole("grid", { name: "Employees" })
  });
  const ada = employeesTable.getByRole("gridcell", { name: /Ada/ });
  await ada.focus();
  await page.keyboard.press("F2");
  await page.keyboard.type("Augusta");
  await page.keyboard.press("Enter");
  await expect(employeesTable.getByRole("status")).toContainText("Updated Name for row 1");
  await page.keyboard.press("ControlOrMeta+C");
  await page.keyboard.press("ArrowDown");
  await page.evaluate(() => navigator.clipboard.writeText("Copied employee"));
  await page.keyboard.press("ControlOrMeta+V");
  await page.keyboard.press("ControlOrMeta+Z");
  await page.keyboard.press("ControlOrMeta+Shift+Z");
});

test("does not alter host controls outside the component root", async ({ page }) => {
  const host = page.getByTestId("host-control");
  await expect(host).toHaveCSS("background-color", "rgb(255, 0, 255)");
  await expect(host).toHaveCSS("font-family", "monospace");
  await expect(host).toHaveCSS("border-top-width", "7px");
});
```

Add a deliberately magenta, monospace, 7px-bordered `data-testid="host-control"` button outside the DataTable root in `DataTableDemo.tsx` so the leakage assertion is real.

- [ ] **Step 6: Run unit and Chromium accessibility flows**

Run:

```bash
corepack pnpm exec vitest run src/react/DataTable.a11y.test.tsx src/react/DataTable.styles.test.tsx
corepack pnpm exec playwright test tests/datatable.spec.ts --project=chromium
```

Expected: zero axe violations, every keyboard/focus assertion passes, and external host styles remain exact.

- [ ] **Step 7: Commit accessibility coverage**

```bash
git add package.json pnpm-lock.yaml src/react/DataTable.a11y.test.tsx src/react/DataTable.tsx src/demo/DataTableDemo.tsx tests/datatable.spec.ts
git commit -m "test: enforce data table accessibility"
```

### Task 13: Pressure-test local tables and document the source API

**Files:**
- Create: `src/react/DataTable.perf.test.tsx`
- Modify: `src/demo/DataTableDemo.tsx`
- Modify: `tests/datatable.spec.ts`
- Modify: `src/index.ts`
- Modify: `docs/embedding.md`
- Modify: `docs/performance.md`

**Interfaces:**
- Produces: bounded 100,000-row/wide-column evidence, multiple-instance/StrictMode cleanup evidence, local API documentation, and the Plan 2 checkpoint required by the implementation program.

- [ ] **Step 1: Write a failing 100,000-row × 1,000-column DOM-bound test**

Create `src/react/DataTable.perf.test.tsx`:

```tsx
import { fireEvent, render, screen } from "@testing-library/react";
import { expect, it } from "vitest";
import { DataTable, type ColumnDef } from "./index";

type Row = { id: string; ordinal: number };

it("bounds DOM for 100000 rows and 1000 columns", () => {
  const rows: readonly Row[] = Array.from({ length: 100_000 }, (_, ordinal) => ({ id: `row-${ordinal}`, ordinal }));
  const columns: readonly ColumnDef<Row>[] = Array.from({ length: 1_000 }, (_, column) => ({
    id: `column-${column}`,
    header: `Column ${column}`,
    width: 96,
    accessor: (row) => `${row.ordinal}:${column}`
  }));
  render(<DataTable aria-label="Stress table" rows={rows} columns={columns} getRowId={(row) => row.id} />);
  const grid = screen.getByRole("grid", { name: "Stress table" });
  Object.defineProperties(grid, {
    clientHeight: { configurable: true, value: 560 },
    clientWidth: { configurable: true, value: 960 },
    scrollTop: { configurable: true, writable: true, value: 2_700_000 },
    scrollLeft: { configurable: true, writable: true, value: 90_000 }
  });
  fireEvent.scroll(grid);
  expect(screen.getAllByRole("gridcell").length).toBeLessThan(600);
  expect(screen.getAllByRole("columnheader").length).toBeLessThan(24);
  expect(screen.queryByText("0:0")).not.toBeInTheDocument();
});
```

- [ ] **Step 2: Run the DOM-bound test and verify red if the bounds are not yet met**

Run: `corepack pnpm exec vitest run src/react/DataTable.perf.test.tsx`

Expected: FAIL if any DataTable layer enumerates all rows, columns, or cells. It must not fail from Node memory exhaustion; if it does, find and remove eager cell projection before proceeding.

- [ ] **Step 3: Remove eager projection and stabilize hot-path identities**

Memoize normalized columns, row-ID maps, query projection, snapshot, viewport descriptors, and stable callback proxies by their real immutable inputs. Do not memoize stale controlled props. Ensure cell rendering calls `snapshot.getCell` only for the virtual cross-product. Keep history operations proportional to affected rows/cells. Add a development assertion that duplicate IDs fail once during normalization rather than once per cell.

- [ ] **Step 4: Add stress-demo and browser bounds**

When `/datatable?stress=1` is loaded, `DataTableDemo` creates the same 100,000 lightweight rows and 1,000 computed accessor columns without materializing a 100-million-cell array. Add visible buttons `Scroll to final row` and `Reset stress viewport`; the first invokes `DataTableHandle.scrollToRow("row-99999")` and sets the grid's `scrollLeft` to `scrollWidth`.

Extend `tests/datatable.spec.ts`:

```ts
test("keeps a 100000 by 1000 table DOM bounded at the far corner", async ({ page }) => {
  await page.goto("/datatable?stress=1");
  await page.getByRole("button", { name: "Scroll to final row" }).click();
  const grid = page.getByRole("grid", { name: "Stress table" });
  await expect.poll(() => grid.getByRole("gridcell").count()).toBeLessThan(600);
  await expect.poll(() => grid.getByRole("columnheader").count()).toBeLessThan(24);
  await expect(grid.getByText(/99999:/).first()).toBeVisible();
});
```

- [ ] **Step 5: Add cleanup, bounded-history, and warning assertions**

Extend unit/browser suites with explicit cases for:

```ts
it("keeps history at or below 100 after 1000 edits");
it("mounts and unmounts 50 StrictMode tables without retained subscribers");
it("keeps two 10000-row instances independent");
it("publishes no state after unmount");
```

In Playwright, register `page.on("console")` and `page.on("pageerror")`; fail on React warnings, unhandled rejections, duplicate-key warnings, or post-unmount updates. Do not use JavaScript heap thresholds in this plan; the final pinned packaging pressure plan owns reproducible memory budgets.

- [ ] **Step 6: Document both local APIs and explicit feature scope**

Update `docs/embedding.md` with complete compilable examples:

```tsx
import { DataTable, type ColumnDef } from "js-spreadsheet";
import "js-spreadsheet/styles.css";

type Employee = { id: string; name: string; salary: number };
const columns: readonly ColumnDef<Employee>[] = [
  {
    id: "name",
    header: "Name",
    accessor: (row) => row.name,
    update: (row, name) => ({ ...row, name: String(name) })
  },
  {
    id: "salary",
    header: "Salary",
    dataType: "number",
    accessor: (row) => row.salary,
    update: (row, salary) => ({ ...row, salary: Number(salary) })
  }
];

<DataTable
  rows={employees}
  columns={columns}
  getRowId={(row) => row.id}
  onRowsChange={(update) => setEmployees((current) => [...update(current)])}
/>;
```

Also show `createLocalRecordTableSession` + `useTableSession`/session-backed `DataTable`, controlled metadata/view-state slices, custom render/editor, and `DataTableHandle`. State explicitly that local operations cover the complete supplied dataset; remote/workbook modes arrive through their dedicated plans; hosts must import the scoped CSS; and GPL/HyperFormula licensing remains applicable.

Update `docs/performance.md` with the deterministic DOM-count results and the production-preview command using `--host 0.0.0.0`.

- [ ] **Step 7: Run the complete local-table checkpoint**

Run from Node 22.13+:

```bash
node --version
corepack pnpm --version
corepack pnpm exec vitest run src/table/core src/table/local src/react
corepack pnpm run build
corepack pnpm exec playwright test tests/datatable.spec.ts --project=chromium
corepack pnpm test
corepack pnpm run test:e2e
git status --short
```

Expected:

- Node is at least 22.13 and pnpm is 11.7.0.
- All table core/local/React tests pass, including fixed-seed property tests and axe.
- The full unit/build/spreadsheet E2E suites pass.
- The 100,000 × 1,000 test stays below 600 gridcells and 24 headers.
- No React warnings, console errors, page errors, unhandled rejection, style leak, or post-unmount publication occurs.
- `git status --short` lists only intentional implementation, test, documentation, package, and plan changes; `.superpowers/` state is absent.

- [ ] **Step 8: Commit the local pressure checkpoint**

```bash
git add src/react/DataTable.perf.test.tsx src/demo/DataTableDemo.tsx tests/datatable.spec.ts src/index.ts docs/embedding.md docs/performance.md
git commit -m "test: pressure local React data tables"
```

## Plan 2 completion boundary

Stop after Task 13 and report the exact verification output to the program owner. The feature branch may proceed to the structured-workbook and remote plans in parallel only when this checkpoint is green. Do not publish, tag, open a PR, or claim the complete product is ready: native workbook tables, remote authority/races/conflicts, final ESM subpaths, Vite/webpack consumer builds, Firefox/WebKit, combined memory baselines, independent review, and the release PR remain in downstream plans.
