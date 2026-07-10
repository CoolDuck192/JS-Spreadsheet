import type { ComponentType, CSSProperties, ReactNode } from "react";
import type { CommandResult } from "../core/commands/types";
import type { TableFeatureConfiguration } from "../table/core/capabilities";
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

export type CellRenderContext<TRow, TValue> = {
  row: Extract<QueryRow<TRow>, { kind: "data" }>;
  column: CoreColumnDef<TRow, TValue>;
  cell: TableCellSnapshot;
};

export type HeaderRenderContext<TRow, TValue> = {
  column: CoreColumnDef<TRow, TValue>;
};

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
