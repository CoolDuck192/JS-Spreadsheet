import type {
  CSSProperties,
  DragEvent as ReactDragEvent,
  KeyboardEvent as ReactKeyboardEvent,
  MouseEvent as ReactMouseEvent,
  ReactNode,
  RefObject
} from "react";
import type { TableCellRef, TableSelection } from "../../table/core/types";

export type GridViewportRow = {
  id: string;
  label: string;
  height: number;
  kind: "data" | "group" | "aggregate";
  ariaRowIndex: number;
  ariaLabel?: string;
  headerAriaLabel?: string;
  pinned?: "top" | "bottom";
};

export type GridViewportColumn = {
  id: string;
  label: string;
  width: number;
  minWidth: number;
  maxWidth: number;
  pinned?: "left" | "right";
  ariaColumnIndex?: number;
};

export type GridViewportCell = {
  ref: TableCellRef;
  ariaLabel: string;
  displayValue: string;
  editable: boolean;
  invalid: boolean;
  className?: string;
  style?: CSSProperties;
  title?: string;
  dataAttributes?: Readonly<Record<`data-${string}`, string | number | boolean | undefined>>;
  columnSpan?: number;
  rowSpan?: number;
};

export type GridViewportApi = {
  ensureCellVisible(rowId: string, columnId: string): void;
};

export type GridViewportHeaderState = {
  className?: string;
  ariaSelected?: boolean;
  tabIndex?: number;
};

export type GridEditorState = TableCellRef & { rawText: string };

export type GridViewportInteraction =
  | { type: "selection-change"; selection: TableSelection }
  | { type: "edit-start"; cell: TableCellRef; initialRawText: string }
  | { type: "edit-change"; rawText: string }
  | {
      type: "edit-commit";
      cell: TableCellRef;
      rawText: string;
      move?: "up" | "down" | "left" | "right";
    }
  | { type: "edit-cancel" }
  | { type: "copy"; selection: TableSelection }
  | { type: "paste"; text: string }
  | { type: "column-resize"; columnId: string; width: number };

export type GridViewportRenderContext = {
  row: GridViewportRow;
  column: GridViewportColumn;
  cell: GridViewportCell;
  selected: boolean;
  active: boolean;
};

export type GridViewportProps = {
  idPrefix: string;
  ariaLabel: string;
  rows: readonly GridViewportRow[];
  ariaRowCount: number;
  columns: readonly GridViewportColumn[];
  ariaColumnCount: number;
  getCell(rowId: string, columnId: string): GridViewportCell;
  selection: TableSelection | null;
  activeCell?: TableCellRef | null;
  editing: GridEditorState | null;
  onInteraction(interaction: GridViewportInteraction): void;
  renderCell?(context: GridViewportRenderContext): ReactNode;
  renderEditor?(context: GridViewportRenderContext & { editing: GridEditorState }): ReactNode;
  renderColumnHeader?(column: GridViewportColumn): ReactNode;
  renderRowHeader?(row: GridViewportRow): ReactNode;
  renderCornerHeader?(): ReactNode;
  renderOverlay?(): ReactNode;
  announce?: string;
  scrollRef?: RefObject<HTMLDivElement | null>;
  showColumnHeaders?: boolean;
  rowHeaderWidth?: number;
  columnHeaderHeight?: number;
  rowOverscan?: number;
  columnOverscan?: number;
  scale?: number;
  resetKey?: string | number;
  retainedRowIds?: readonly string[];
  retainedColumnIds?: readonly string[];
  rootClassName?: string;
  rootStyle?: CSSProperties;
  rootDataAttributes?: Readonly<Record<string, string | number | boolean | undefined>>;
  canvasClassName?: string;
  canvasStyle?: CSSProperties;
  interactionEventMode?: "pointer" | "mouse";
  onBeforeKeyDown?(event: ReactKeyboardEvent<HTMLDivElement>): void;
  onUnhandledKeyDown?(event: ReactKeyboardEvent<HTMLElement>): void;
  onCellMouseEnter?(cell: TableCellRef): boolean | void;
  onCellContextMenu?(cell: TableCellRef, event: ReactMouseEvent<HTMLDivElement>): void;
  getColumnHeaderState?(column: GridViewportColumn): GridViewportHeaderState;
  getRowHeaderState?(row: GridViewportRow): GridViewportHeaderState;
  onColumnHeaderMouseDown?(column: GridViewportColumn, event: ReactMouseEvent<HTMLDivElement>): void;
  onColumnHeaderMouseEnter?(column: GridViewportColumn, event: ReactMouseEvent<HTMLDivElement>): void;
  onColumnHeaderClick?(column: GridViewportColumn, event: ReactMouseEvent<HTMLDivElement>): void;
  onColumnHeaderKeyDown?(column: GridViewportColumn, event: ReactKeyboardEvent<HTMLDivElement>): void;
  onColumnHeaderDragOver?(column: GridViewportColumn, event: ReactDragEvent<HTMLDivElement>): void;
  onColumnHeaderDrop?(column: GridViewportColumn, event: ReactDragEvent<HTMLDivElement>): void;
  onRowHeaderMouseDown?(row: GridViewportRow, event: ReactMouseEvent<HTMLDivElement>): void;
  onRowHeaderMouseEnter?(row: GridViewportRow, event: ReactMouseEvent<HTMLDivElement>): void;
  onRowHeaderClick?(row: GridViewportRow, event: ReactMouseEvent<HTMLDivElement>): void;
  onRowHeaderKeyDown?(row: GridViewportRow, event: ReactKeyboardEvent<HTMLDivElement>): void;
  onRootMouseUp?(event: ReactMouseEvent<HTMLDivElement>): void;
  onRegisterApi?(api: GridViewportApi): void;
};
