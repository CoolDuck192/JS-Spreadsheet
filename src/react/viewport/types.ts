import type {
  CSSProperties,
  KeyboardEvent as ReactKeyboardEvent,
  ReactNode
} from "react";
import type { TableCellRef, TableSelection } from "../../table/core/types";

export type GridViewportRow = {
  id: string;
  label: string;
  height: number;
  kind: "data" | "group" | "aggregate";
  ariaRowIndex: number;
};

export type GridViewportColumn = {
  id: string;
  label: string;
  width: number;
  minWidth: number;
  maxWidth: number;
  pinned?: "left" | "right";
};

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
  editing: GridEditorState | null;
  onInteraction(interaction: GridViewportInteraction): void;
  renderCell?(context: GridViewportRenderContext): ReactNode;
  renderEditor?(context: GridViewportRenderContext & { editing: GridEditorState }): ReactNode;
  renderColumnHeader?(column: GridViewportColumn): ReactNode;
  renderRowHeader?(row: GridViewportRow): ReactNode;
  announce?: string;
  onUnhandledKeyDown?(event: ReactKeyboardEvent<HTMLElement>): void;
};
