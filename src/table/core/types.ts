import type { CommandResult, TableIssue } from "../../core/commands/types";
import type { ParsedCellInput } from "../../core/values/parseCellInput";
import type {
  FilterExpression,
  PaginationRequest,
  QueryResult,
  QueryRow,
  TableAggregateRequest,
  TableGrouping,
  TableSort,
  TotalCount
} from "./query";
import type {
  TableCapabilities,
  TableFeature,
  TableOperationState
} from "./capabilities";

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

export type TableCellFormat = {
  numberFormat?: "general" | "number" | "currency" | "percent" | "date" | "datetime" | "financial" | "financial2" | "accounting";
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

export type TablePageGap = {
  kind: "evicted-pages";
  at: number;
  omittedPages: number;
  omittedItems: number;
};

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
  pageGaps?: readonly TablePageGap[];
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
