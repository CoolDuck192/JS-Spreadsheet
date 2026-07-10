import {
  useCallback,
  useMemo,
  useRef,
  type CSSProperties,
  type KeyboardEvent as ReactKeyboardEvent
} from "react";
import type { TableCellRef, TableSelection } from "../../table/core/types";
import { useGridInteraction, gridCellDomId, gridLiveRegionDomId } from "./useGridInteraction";
import { useTwoAxisVirtualizer } from "./useTwoAxisVirtualizer";
import type {
  GridViewportCell,
  GridViewportColumn,
  GridViewportProps,
  GridViewportRenderContext,
  GridViewportRow
} from "./types";

const HEADER_HEIGHT = 28;
const ROW_HEADER_WIDTH = 56;

export function GridViewport({
  idPrefix,
  ariaLabel,
  rows,
  ariaRowCount,
  columns,
  ariaColumnCount,
  getCell,
  selection,
  editing,
  onInteraction,
  renderCell,
  renderEditor,
  renderColumnHeader,
  renderRowHeader,
  announce = "",
  onUnhandledKeyDown
}: GridViewportProps) {
  const rootRef = useRef<HTMLDivElement>(null);
  const rowHeaderWidth = renderRowHeader ? ROW_HEADER_WIDTH : 0;
  const getRowKey = useCallback((index: number) => rows[index].id, [rows]);
  const getColumnKey = useCallback((index: number) => columns[index].id, [columns]);
  const getRowSize = useCallback((index: number) => rows[index].height, [rows]);
  const getColumnSize = useCallback((index: number) => columns[index].width, [columns]);
  const virtualizer = useTwoAxisVirtualizer({
    scrollRef: rootRef,
    rowCount: rows.length,
    columnCount: columns.length,
    getRowKey,
    getColumnKey,
    getRowSize,
    getColumnSize,
    rowOverscan: 2,
    columnOverscan: 2,
    rowViewportInset: HEADER_HEIGHT,
    columnViewportInset: rowHeaderWidth
  });
  const interaction = useGridInteraction({
    idPrefix,
    rows,
    columns,
    selection,
    editing,
    onInteraction,
    ensureCellVisible: virtualizer.ensureCellVisible,
    rootRef,
    getInitialRawText: (cell) => getCell(cell.rowId, cell.columnId).displayValue,
    isCellEditable: (cell) => getCell(cell.rowId, cell.columnId).editable
  });
  const selectedBounds = useMemo(() => selectionBounds(selection, rows, columns), [columns, rows, selection]);
  const activeCell = selection?.focus ?? null;
  const liveRegionId = gridLiveRegionDomId(idPrefix);
  const canvasWidth = rowHeaderWidth + virtualizer.totalWidth;
  const canvasHeight = HEADER_HEIGHT + virtualizer.totalHeight;
  const activeDescendantId =
    selection &&
    virtualizer.visibleRows.some((measurement) => rows[measurement.index].id === selection.focus.rowId) &&
    virtualizer.visibleColumns.some((measurement) => columns[measurement.index].id === selection.focus.columnId)
      ? interaction.activeDescendantId
      : undefined;

  function handleKeyDown(event: ReactKeyboardEvent<HTMLDivElement>) {
    if (isEditorEventTarget(event.target)) {
      return;
    }
    interaction.onKeyDown(event);
    if (!event.defaultPrevented) {
      onUnhandledKeyDown?.(event);
    }
  }

  return (
    <div
      ref={rootRef}
      role="grid"
      aria-label={ariaLabel}
      aria-rowcount={ariaRowCount}
      aria-colcount={ariaColumnCount}
      aria-activedescendant={activeDescendantId}
      aria-describedby={liveRegionId}
      data-viewport-kernel="shared"
      tabIndex={-1}
      style={{ position: "relative", overflow: "auto" }}
      onScroll={virtualizer.onScroll}
      onKeyDown={handleKeyDown}
      onPointerUp={interaction.onPointerUp}
      onPointerCancel={interaction.onPointerUp}
      onCopy={(event) => {
        if (isEditorEventTarget(event.target)) {
          return;
        }
        if (!selection) {
          return;
        }
        event.preventDefault();
        onInteraction({ type: "copy", selection });
      }}
      onPaste={(event) => {
        if (isEditorEventTarget(event.target)) {
          return;
        }
        event.preventDefault();
        const text = event.clipboardData.getData("text/plain") || event.clipboardData.getData("Text");
        onInteraction({ type: "paste", text });
      }}
    >
      <div
        data-grid-viewport-canvas="true"
        style={{ position: "relative", width: canvasWidth, height: canvasHeight, minWidth: "100%" }}
      >
        <div
          role="row"
          aria-label="Column headers"
          aria-rowindex={1}
          style={{ position: "sticky", top: 0, zIndex: 3, height: HEADER_HEIGHT }}
        >
          {renderRowHeader ? (
            <div
              role="columnheader"
              aria-label="Row headers"
              aria-colindex={1}
              style={{ ...headerCellStyle(0, ROW_HEADER_WIDTH), position: "sticky", left: 0, zIndex: 4 }}
            />
          ) : null}
          {virtualizer.visibleColumns.map((measurement) => {
            const column = columns[measurement.index];
            return (
              <div
                key={column.id}
                role="columnheader"
                aria-label={column.label}
                aria-colindex={measurement.index + 1 + (renderRowHeader ? 1 : 0)}
                style={headerCellStyle(rowHeaderWidth + measurement.start, measurement.size)}
              >
                {renderColumnHeader ? renderColumnHeader(column) : column.label}
              </div>
            );
          })}
        </div>

        {virtualizer.visibleRows.map((rowMeasurement) => {
          const row = rows[rowMeasurement.index];
          return (
            <div
              key={row.id}
              role="row"
              aria-label={`Row ${row.label}`}
              aria-rowindex={row.ariaRowIndex}
              data-row-kind={row.kind}
              style={{
                position: "absolute",
                top: HEADER_HEIGHT + rowMeasurement.start,
                left: 0,
                width: canvasWidth,
                height: rowMeasurement.size
              }}
            >
              {renderRowHeader ? (
                <div
                  role="rowheader"
                  aria-label={row.label}
                  aria-colindex={1}
                  style={rowHeaderCellStyle(rowMeasurement.size)}
                >
                  {renderRowHeader(row)}
                </div>
              ) : null}
              {virtualizer.visibleColumns.map((columnMeasurement) => {
                const column = columns[columnMeasurement.index];
                const cell = getCell(row.id, column.id);
                const selected = isSelected(rowMeasurement.index, columnMeasurement.index, selectedBounds);
                const active = sameCell(cell.ref, activeCell);
                const context: GridViewportRenderContext = { row, column, cell, selected, active };
                return (
                  <div
                    key={column.id}
                    id={gridCellDomId(idPrefix, cell.ref)}
                    role="gridcell"
                    aria-label={cell.ariaLabel}
                    aria-colindex={columnMeasurement.index + 1 + (renderRowHeader ? 1 : 0)}
                    aria-selected={selected}
                    aria-readonly={!cell.editable}
                    aria-invalid={cell.invalid}
                    aria-colspan={cell.columnSpan}
                    aria-rowspan={cell.rowSpan}
                    className={cell.className}
                    tabIndex={active ? 0 : -1}
                    style={cellStyle(
                      rowHeaderWidth + columnMeasurement.start,
                      columnMeasurement.size,
                      rowMeasurement.size,
                      cell.style
                    )}
                    onPointerDown={(event) => {
                      if (event.button !== 0) {
                        return;
                      }
                      event.currentTarget.focus({ preventScroll: true });
                      interaction.onCellPointerDown(cell.ref, event.shiftKey);
                    }}
                    onPointerEnter={() => interaction.onCellPointerEnter(cell.ref)}
                    onClick={(event) => {
                      if (event.detail === 0) {
                        interaction.onCellPointerDown(cell.ref, event.shiftKey);
                        interaction.onPointerUp();
                      }
                    }}
                    onDoubleClick={() => {
                      if (cell.editable) {
                        onInteraction({ type: "edit-start", cell: cell.ref, initialRawText: cell.displayValue });
                      }
                    }}
                  >
                    {renderCell ? renderCell(context) : cell.displayValue}
                  </div>
                );
              })}
            </div>
          );
        })}

        {editing ? (
          <EditorOverlay
            editing={editing}
            rows={rows}
            columns={columns}
            rowHeaderWidth={rowHeaderWidth}
            getCell={getCell}
            getCellRect={virtualizer.getCellRect}
            renderEditor={renderEditor}
            onInteraction={onInteraction}
          />
        ) : null}
      </div>
      <div id={liveRegionId} role="status" aria-live="polite" style={visuallyHiddenStyle}>
        {announce}
      </div>
    </div>
  );
}

function EditorOverlay({
  editing,
  rows,
  columns,
  rowHeaderWidth,
  getCell,
  getCellRect,
  renderEditor,
  onInteraction
}: {
  editing: NonNullable<GridViewportProps["editing"]>;
  rows: readonly GridViewportRow[];
  columns: readonly GridViewportColumn[];
  rowHeaderWidth: number;
  getCell: GridViewportProps["getCell"];
  getCellRect(rowIndex: number, columnIndex: number): { top: number; left: number; width: number; height: number };
  renderEditor: GridViewportProps["renderEditor"];
  onInteraction: GridViewportProps["onInteraction"];
}) {
  const rowIndex = rows.findIndex((row) => row.id === editing.rowId);
  const columnIndex = columns.findIndex((column) => column.id === editing.columnId);
  if (rowIndex < 0 || columnIndex < 0) {
    return null;
  }
  const row = rows[rowIndex];
  const column = columns[columnIndex];
  const cell = getCell(row.id, column.id);
  const rect = getCellRect(rowIndex, columnIndex);
  const context = { row, column, cell, selected: true, active: true, editing };
  return (
    <div
      data-grid-editor-overlay="true"
      style={{
        position: "absolute",
        zIndex: 5,
        top: HEADER_HEIGHT + rect.top,
        left: rowHeaderWidth + rect.left,
        width: rect.width,
        height: rect.height
      }}
    >
      {renderEditor ? (
        renderEditor(context)
      ) : (
        <input
          autoFocus
          aria-label={`Edit ${cell.ariaLabel}`}
          value={editing.rawText}
          onChange={(event) => onInteraction({ type: "edit-change", rawText: event.currentTarget.value })}
          onKeyDown={(event) => {
            if (event.key === "Escape") {
              event.preventDefault();
              event.stopPropagation();
              onInteraction({ type: "edit-cancel" });
              return;
            }
            if (event.key === "Enter" || event.key === "Tab") {
              event.preventDefault();
              event.stopPropagation();
              const move =
                event.key === "Tab"
                  ? event.shiftKey
                    ? "left"
                    : "right"
                  : event.shiftKey
                    ? "up"
                    : "down";
              onInteraction({ type: "edit-commit", cell: cell.ref, rawText: editing.rawText, move });
            }
          }}
          style={{ boxSizing: "border-box", width: "100%", height: "100%" }}
        />
      )}
    </div>
  );
}

type SelectionBounds = { firstRow: number; lastRow: number; firstColumn: number; lastColumn: number } | null;

function selectionBounds(
  selection: TableSelection | null,
  rows: readonly GridViewportRow[],
  columns: readonly GridViewportColumn[]
): SelectionBounds {
  if (!selection) {
    return null;
  }
  const anchorRow = rows.findIndex((row) => row.id === selection.anchor.rowId);
  const focusRow = rows.findIndex((row) => row.id === selection.focus.rowId);
  const anchorColumn = columns.findIndex((column) => column.id === selection.anchor.columnId);
  const focusColumn = columns.findIndex((column) => column.id === selection.focus.columnId);
  if (anchorRow < 0 || focusRow < 0 || anchorColumn < 0 || focusColumn < 0) {
    return null;
  }
  return {
    firstRow: Math.min(anchorRow, focusRow),
    lastRow: Math.max(anchorRow, focusRow),
    firstColumn: Math.min(anchorColumn, focusColumn),
    lastColumn: Math.max(anchorColumn, focusColumn)
  };
}

function isSelected(row: number, column: number, bounds: SelectionBounds): boolean {
  return Boolean(
    bounds &&
      row >= bounds.firstRow &&
      row <= bounds.lastRow &&
      column >= bounds.firstColumn &&
      column <= bounds.lastColumn
  );
}

function sameCell(left: TableCellRef, right: TableCellRef | null): boolean {
  return Boolean(right && left.rowId === right.rowId && left.columnId === right.columnId);
}

function headerCellStyle(left: number, width: number): CSSProperties {
  return {
    position: "absolute",
    top: 0,
    left,
    width,
    height: HEADER_HEIGHT,
    boxSizing: "border-box"
  };
}

function rowHeaderCellStyle(height: number): CSSProperties {
  return {
    position: "sticky",
    top: 0,
    left: 0,
    zIndex: 2,
    width: ROW_HEADER_WIDTH,
    height,
    boxSizing: "border-box"
  };
}

function cellStyle(left: number, width: number, height: number, style: CSSProperties | undefined): CSSProperties {
  return {
    ...style,
    position: "absolute",
    top: 0,
    left,
    width,
    height,
    boxSizing: "border-box"
  };
}

function isEditorEventTarget(target: EventTarget | null): boolean {
  return target instanceof Element && Boolean(target.closest('[data-grid-editor-overlay="true"]'));
}

const visuallyHiddenStyle: CSSProperties = {
  position: "absolute",
  width: 1,
  height: 1,
  padding: 0,
  margin: -1,
  overflow: "hidden",
  clip: "rect(0 0 0 0)",
  whiteSpace: "nowrap",
  border: 0
};
