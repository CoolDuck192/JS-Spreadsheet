import {
  useLayoutEffect,
  useRef,
  type ChangeEvent,
  type KeyboardEvent,
  type ReactNode
} from "react";
import type { QueryRow, QueryScalar } from "../table/core/query";
import type { TableCellSnapshot, TableDiagnosticEvent } from "../table/core/types";
import { DataTableExtensionBoundary } from "./DataTableExtensionBoundary";
import type { CellEditorProps, CellRenderContext, ColumnDef } from "./tableTypes";

export function DataTableCell<TRow>({
  row,
  column,
  cell,
  isFirstColumn,
  onToggleExpanded,
  onMeasure,
  onDiagnostic
}: {
  row: QueryRow<TRow>;
  column: ColumnDef<TRow>;
  cell: TableCellSnapshot;
  isFirstColumn: boolean;
  onToggleExpanded?(rowId: string, expanded: boolean): void;
  onMeasure?(rowId: string, columnId: string, height: number): void;
  onDiagnostic?(event: TableDiagnosticEvent): void;
}) {
  const content = row.kind === "data" && column.cell
    ? (
        <DataTableExtensionBoundary columnId={column.id} slot="cell" onDiagnostic={onDiagnostic}>
          <InvokeCellRenderer render={column.cell} context={{ row, column, cell }} />
        </DataTableExtensionBoundary>
      )
    : defaultCellContent(row, column.id, cell.displayValue);

  return (
    <MeasuredCell rowId={row.id} columnId={column.id} onMeasure={onMeasure}>
      <span
        className="js-spreadsheet-data-table__cell-content"
        style={row.depth > 0 && isFirstColumn ? { paddingInlineStart: row.depth * 16 } : undefined}
      >
        {row.kind === "data" && row.hasChildren && isFirstColumn ? (
          <button
            type="button"
            aria-label={`${row.expanded ? "Collapse" : "Expand"} ${String(cell.displayValue || row.id)}`}
            aria-expanded={row.expanded === true}
            onPointerDown={(event) => {
              event.preventDefault();
              event.stopPropagation();
            }}
            onMouseDown={(event) => {
              event.preventDefault();
              event.stopPropagation();
            }}
            onClick={(event) => {
              event.stopPropagation();
              onToggleExpanded?.(row.id, row.expanded !== true);
            }}
          >
            {row.expanded ? "−" : "+"}
          </button>
        ) : null}
        {content}
      </span>
    </MeasuredCell>
  );
}

export function DataTableEditor<TRow>({
  row,
  column,
  cell,
  rawText,
  onChange,
  onCommit,
  onCancel,
  onDiagnostic
}: Omit<CellEditorProps<TRow, unknown>, "column" | "onCommit"> & {
  column: ColumnDef<TRow>;
  onCommit(move?: "up" | "down" | "left" | "right"): void;
  onDiagnostic?(event: TableDiagnosticEvent): void;
}) {
  if (column.editor) {
    const Editor = column.editor;
    return (
      <DataTableExtensionBoundary columnId={column.id} slot="editor" onDiagnostic={onDiagnostic}>
        <Editor
          row={row}
          column={column}
          cell={cell}
          rawText={rawText}
          onChange={onChange}
          onCommit={() => onCommit()}
          onCancel={onCancel}
        />
      </DataTableExtensionBoundary>
    );
  }

  const listValues = cell.metadata.validation?.kind === "list"
    ? cell.metadata.validation.values
    : null;
  if (column.dataType === "boolean" || listValues) {
    const values = listValues ?? ["true", "false"];
    return (
      <select
        autoFocus
        aria-label={`Edit ${cell.rowId} ${columnLabel(column)}`}
        value={rawText}
        onChange={(event) => onChange(event.currentTarget.value)}
        onKeyDown={(event) => handleEditorKeyDown(event, onCommit, onCancel)}
      >
        {values.map((value) => <option key={String(value)} value={String(value)}>{String(value)}</option>)}
      </select>
    );
  }

  return (
    <input
      autoFocus
      aria-label={`Edit ${cell.rowId} ${columnLabel(column)}`}
      value={rawText}
      onChange={(event: ChangeEvent<HTMLInputElement>) => onChange(event.currentTarget.value)}
      onKeyDown={(event) => handleEditorKeyDown(event, onCommit, onCancel)}
    />
  );
}

function InvokeCellRenderer<TRow>({
  render,
  context
}: {
  render: (context: CellRenderContext<TRow, unknown>) => ReactNode;
  context: CellRenderContext<TRow, unknown>;
}) {
  return <>{render(context)}</>;
}

function MeasuredCell({
  rowId,
  columnId,
  onMeasure,
  children
}: {
  rowId: string;
  columnId: string;
  onMeasure?(rowId: string, columnId: string, height: number): void;
  children: ReactNode;
}) {
  const ref = useRef<HTMLDivElement>(null);
  useLayoutEffect(() => {
    const element = ref.current;
    if (!element || !onMeasure) return;
    const publish = (height: number) => onMeasure(rowId, columnId, Math.max(28, Math.ceil(height)));
    publish(Math.max(element.scrollHeight, element.getBoundingClientRect().height));
    if (typeof ResizeObserver === "undefined") return;
    const observer = new ResizeObserver((entries) => {
      const entry = entries[0];
      if (entry) publish(entry.contentRect.height);
    });
    observer.observe(element);
    return () => observer.disconnect();
  }, [columnId, onMeasure, rowId]);
  return <div ref={ref} data-row-measure={rowId}>{children}</div>;
}

function handleEditorKeyDown(
  event: KeyboardEvent<HTMLInputElement | HTMLSelectElement>,
  onCommit: (move?: "up" | "down" | "left" | "right") => void,
  onCancel: () => void
) {
  if (event.key === "Escape") {
    event.preventDefault();
    event.stopPropagation();
    onCancel();
  } else if (event.key === "Enter" || event.key === "Tab") {
    event.preventDefault();
    event.stopPropagation();
    onCommit(
      event.key === "Tab"
        ? event.shiftKey ? "left" : "right"
        : event.shiftKey ? "up" : "down"
    );
  }
}

function defaultCellContent<TRow>(row: QueryRow<TRow>, columnId: string, displayValue: string): ReactNode {
  if (row.kind === "data") return displayValue;
  if (row.kind === "group" && row.columnId === columnId) {
    return `${formatScalar(row.key)} (${row.count})`;
  }
  const aggregate = row.aggregates[columnId];
  return aggregate === undefined || aggregate === null ? displayValue : String(aggregate);
}

function formatScalar(value: QueryScalar): string {
  return value.type === "null" ? "Blank" : String(value.value);
}

export function columnLabel<TRow>(column: ColumnDef<TRow>): string {
  return typeof column.header === "string"
    ? column.header
    : column.id.replace(/[-_]+/g, " ").replace(/^./, (character) => character.toUpperCase());
}
