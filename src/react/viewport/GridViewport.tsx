import {
  useCallback,
  useLayoutEffect,
  useMemo,
  useRef,
  type CSSProperties,
  type KeyboardEvent as ReactKeyboardEvent,
  type MouseEvent as ReactMouseEvent
} from "react";
import type { AxisMeasurement } from "../../core/viewport/axis";
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

const DEFAULT_HEADER_HEIGHT = 28;
const DEFAULT_ROW_HEADER_WIDTH = 56;

export function GridViewport({
  idPrefix,
  ariaLabel,
  rows,
  ariaRowCount,
  columns,
  ariaColumnCount,
  getCell,
  selection,
  activeCell: controlledActiveCell,
  editing,
  onInteraction,
  renderCell,
  renderEditor,
  renderColumnHeader,
  renderRowHeader,
  renderCornerHeader,
  renderOverlay,
  announce = "",
  scrollRef,
  showColumnHeaders = true,
  rowHeaderWidth: requestedRowHeaderWidth = DEFAULT_ROW_HEADER_WIDTH,
  columnHeaderHeight: requestedColumnHeaderHeight = DEFAULT_HEADER_HEIGHT,
  rowOverscan = 2,
  columnOverscan = 2,
  scale = 1,
  resetKey,
  retainedRowIds = [],
  retainedColumnIds = [],
  rootClassName,
  rootStyle,
  rootDataAttributes,
  canvasClassName,
  canvasStyle,
  interactionEventMode = "pointer",
  onBeforeKeyDown,
  onUnhandledKeyDown,
  onCellMouseEnter,
  onCellContextMenu,
  onColumnHeaderContextMenu,
  onReadOnlyCellEditAttempt,
  getColumnHeaderState,
  getRowHeaderState,
  onColumnHeaderMouseDown,
  onColumnHeaderMouseEnter,
  onColumnHeaderClick,
  onColumnHeaderKeyDown,
  onColumnHeaderDragOver,
  onColumnHeaderDrop,
  onRowHeaderMouseDown,
  onRowHeaderMouseEnter,
  onRowHeaderClick,
  onRowHeaderKeyDown,
  onRootMouseUp,
  onRegisterApi
}: GridViewportProps) {
  const internalRootRef = useRef<HTMLDivElement>(null);
  const rootRef = scrollRef ?? internalRootRef;
  const pendingFocusFrameRef = useRef<number | null>(null);
  const rowHeaderWidth = renderRowHeader ? requestedRowHeaderWidth : 0;
  const headerHeight = showColumnHeaders ? requestedColumnHeaderHeight : 0;
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
    rowOverscan,
    columnOverscan,
    rowViewportInset: headerHeight,
    columnViewportInset: rowHeaderWidth,
    scale,
    resetKey
  });
  const activeCell = controlledActiveCell ?? selection?.focus ?? null;
  const interaction = useGridInteraction({
    idPrefix,
    rows,
    columns,
    selection,
    activeCell,
    editing,
    onInteraction,
    ensureCellVisible: virtualizer.ensureCellVisible,
    rootRef,
    getInitialRawText: (cell) => getCell(cell.rowId, cell.columnId).displayValue,
    isCellEditable: (cell) => getCell(cell.rowId, cell.columnId).editable
  });
  const liveRegionId = gridLiveRegionDomId(idPrefix);
  const canvasWidth = rowHeaderWidth + virtualizer.totalWidth;
  const canvasHeight = headerHeight + virtualizer.totalHeight;
  const pinnedColumnOffsets = useMemo(() => createPinnedColumnOffsets(columns), [columns]);
  const pinnedRowOffsets = useMemo(() => createPinnedRowOffsets(rows), [rows]);
  const rowMeasurementsById = useMemo(
    () => new Map(virtualizer.rowMeasurements.map((measurement) => [measurement.key, measurement])),
    [virtualizer.rowMeasurements]
  );
  const columnMeasurementsById = useMemo(
    () => new Map(virtualizer.columnMeasurements.map((measurement) => [measurement.key, measurement])),
    [virtualizer.columnMeasurements]
  );
  const selectedBounds = useMemo(
    () => selectionBounds(selection, rowMeasurementsById, columnMeasurementsById),
    [columnMeasurementsById, rowMeasurementsById, selection]
  );
  const renderedRows = useMemo(
    () =>
      collectRenderedMeasurements(
        virtualizer.visibleRows,
        rowMeasurementsById,
        new Set([
          ...retainedRowIds,
          ...rows.filter((row) => row.pinned).map((row) => row.id),
          ...(editing ? [editing.rowId] : [])
        ])
      ),
    [editing, retainedRowIds, rowMeasurementsById, rows, virtualizer.visibleRows]
  );
  const renderedColumns = useMemo(
    () =>
      collectRenderedMeasurements(
        virtualizer.visibleColumns,
        columnMeasurementsById,
        new Set([
          ...retainedColumnIds,
          ...columns.filter((column) => column.pinned).map((column) => column.id),
          ...(editing ? [editing.columnId] : [])
        ])
      ),
    [columnMeasurementsById, columns, editing, retainedColumnIds, virtualizer.visibleColumns]
  );
  const activeDescendantId =
    activeCell &&
    renderedRows.some((measurement) => rows[measurement.index].id === activeCell.rowId) &&
    renderedColumns.some((measurement) => columns[measurement.index].id === activeCell.columnId)
      ? interaction.activeDescendantId
      : undefined;

  const ensureCellVisible = useCallback(
    (rowId: string, columnId: string) => {
      const row = rowMeasurementsById.get(rowId);
      const column = columnMeasurementsById.get(columnId);
      if (row && column) {
        virtualizer.ensureCellVisible(row.index, column.index);
      }
    },
    [columnMeasurementsById, rowMeasurementsById, virtualizer.ensureCellVisible]
  );

  const cancelPendingFocus = useCallback(() => {
    if (pendingFocusFrameRef.current !== null) {
      cancelAnimationFrame(pendingFocusFrameRef.current);
      pendingFocusFrameRef.current = null;
    }
  }, []);

  const focusCell = useCallback(
    (rowId: string, columnId: string) => {
      cancelPendingFocus();
      ensureCellVisible(rowId, columnId);
      const cellId = gridCellDomId(idPrefix, { rowId, columnId });
      const focusRenderedCell = () => {
        const cell = rootRef.current?.querySelector<HTMLElement>(`[id="${cellId}"]`);
        if (cell instanceof HTMLElement) {
          cell.focus({ preventScroll: true });
          return true;
        }
        return false;
      };
      if (!focusRenderedCell()) {
        pendingFocusFrameRef.current = requestAnimationFrame(() => {
          pendingFocusFrameRef.current = null;
          focusRenderedCell();
        });
      }
    },
    [cancelPendingFocus, ensureCellVisible, idPrefix, rootRef]
  );

  useLayoutEffect(() => cancelPendingFocus, [cancelPendingFocus, focusCell]);

  useLayoutEffect(() => {
    onRegisterApi?.({ ensureCellVisible, focusCell });
  }, [ensureCellVisible, focusCell, onRegisterApi]);

  function handleKeyDown(event: ReactKeyboardEvent<HTMLDivElement>) {
    if (isEditorEventTarget(event.target)) {
      return;
    }
    onBeforeKeyDown?.(event);
    if (event.defaultPrevented) {
      return;
    }
    interaction.onKeyDown(event);
    if (!event.defaultPrevented) {
      onUnhandledKeyDown?.(event);
    }
  }

  function handleCellDown(cell: GridViewportCell, event: ReactMouseEvent<HTMLDivElement>) {
    if (event.button !== 0) {
      return;
    }
    event.currentTarget.focus({ preventScroll: true });
    interaction.onCellPointerDown(cell.ref, event.shiftKey);
  }

  function handleCellEnter(cell: GridViewportCell) {
    if (onCellMouseEnter?.(cell.ref)) {
      return;
    }
    interaction.onCellPointerEnter(cell.ref);
  }

  return (
    <div
      {...rootDataAttributes}
      ref={rootRef}
      role="grid"
      aria-label={ariaLabel}
      aria-rowcount={ariaRowCount}
      aria-colcount={ariaColumnCount}
      aria-activedescendant={activeDescendantId}
      aria-describedby={liveRegionId}
      data-viewport-kernel="shared"
      className={rootClassName}
      tabIndex={-1}
      style={{ ...rootStyle, position: "relative", overflow: "auto" }}
      onScroll={virtualizer.onScroll}
      onKeyDown={handleKeyDown}
      onPointerUp={interactionEventMode === "pointer" ? interaction.onPointerUp : undefined}
      onPointerCancel={interactionEventMode === "pointer" ? interaction.onPointerUp : undefined}
      onMouseUp={(event) => {
        if (interactionEventMode === "mouse") {
          interaction.onPointerUp();
        }
        onRootMouseUp?.(event);
      }}
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
        className={canvasClassName}
        style={{
          ...canvasStyle,
          position: "relative",
          width: canvasWidth,
          height: canvasHeight,
          minWidth: "100%"
        }}
      >
        {showColumnHeaders ? (
          <div
            role="row"
            aria-label="Column headers"
            aria-rowindex={1}
            style={{ position: "sticky", top: 0, zIndex: 3, height: headerHeight }}
          >
            {renderRowHeader ? (
              <div
                role="columnheader"
                aria-label="Row headers"
                aria-colindex={1}
                style={{
                  ...headerCellStyle(0, rowHeaderWidth, headerHeight),
                  position: "sticky",
                  left: 0,
                  zIndex: 4
                }}
              >
                {renderCornerHeader?.()}
              </div>
            ) : null}
            {renderedColumns.map((measurement) => {
              const column = columns[measurement.index];
              const headerState = getColumnHeaderState?.(column);
              return (
                <div
                  key={column.id}
                  role="columnheader"
                  aria-label={column.label}
                  aria-colindex={column.ariaColumnIndex ?? measurement.index + 1 + (renderRowHeader ? 1 : 0)}
                  aria-selected={headerState?.ariaSelected}
                  tabIndex={headerState?.tabIndex}
                  className={headerState?.className}
                  style={columnPositionStyle(
                    headerCellStyle(rowHeaderWidth + measurement.start, measurement.size, headerHeight),
                    column,
                    rowHeaderWidth,
                    pinnedColumnOffsets
                  )}
                  onMouseDown={(event) => onColumnHeaderMouseDown?.(column, event)}
                  onMouseEnter={(event) => onColumnHeaderMouseEnter?.(column, event)}
                  onClick={(event) => onColumnHeaderClick?.(column, event)}
                  onContextMenu={onColumnHeaderContextMenu
                    ? (event) => onColumnHeaderContextMenu(column, event)
                    : undefined}
                  onKeyDown={(event) => onColumnHeaderKeyDown?.(column, event)}
                  onDragOver={(event) => onColumnHeaderDragOver?.(column, event)}
                  onDrop={(event) => onColumnHeaderDrop?.(column, event)}
                >
                  {renderColumnHeader ? renderColumnHeader(column) : column.label}
                </div>
              );
            })}
          </div>
        ) : null}

        {renderedRows.map((rowMeasurement) => {
          const row = rows[rowMeasurement.index];
          const rowHeaderState = getRowHeaderState?.(row);
          return (
            <div
              key={row.id}
              role="row"
              aria-label={row.ariaLabel ?? `Row ${row.label}`}
              aria-rowindex={row.ariaRowIndex}
              data-row-kind={row.kind}
              style={rowPositionStyle(
                rowMeasurement,
                row,
                canvasWidth,
                headerHeight,
                pinnedRowOffsets
              )}
            >
              {renderRowHeader ? (
                <div
                  role="rowheader"
                  aria-label={row.headerAriaLabel ?? row.label}
                  aria-colindex={1}
                  aria-selected={rowHeaderState?.ariaSelected}
                  tabIndex={rowHeaderState?.tabIndex}
                  className={rowHeaderState?.className}
                  style={rowHeaderCellStyle(rowMeasurement.size, rowHeaderWidth)}
                  onMouseDown={(event) => onRowHeaderMouseDown?.(row, event)}
                  onMouseEnter={(event) => onRowHeaderMouseEnter?.(row, event)}
                  onClick={(event) => onRowHeaderClick?.(row, event)}
                  onKeyDown={(event) => onRowHeaderKeyDown?.(row, event)}
                >
                  {renderRowHeader(row)}
                </div>
              ) : null}
              {renderedColumns.map((columnMeasurement) => {
                const column = columns[columnMeasurement.index];
                const cell = getCell(row.id, column.id);
                const selected = isSelected(rowMeasurement.index, columnMeasurement.index, selectedBounds);
                const active = sameCell(cell.ref, activeCell);
                const context: GridViewportRenderContext = { row, column, cell, selected, active };
                const width = spanSize(
                  virtualizer.columnMeasurements,
                  columnMeasurement.index,
                  cell.columnSpan,
                  columnMeasurement.size
                );
                const height = spanSize(
                  virtualizer.rowMeasurements,
                  rowMeasurement.index,
                  cell.rowSpan,
                  rowMeasurement.size
                );
                return (
                  <div
                    key={column.id}
                    id={gridCellDomId(idPrefix, cell.ref)}
                    role="gridcell"
                    aria-label={cell.ariaLabel}
                    aria-colindex={column.ariaColumnIndex ?? columnMeasurement.index + 1 + (renderRowHeader ? 1 : 0)}
                    aria-selected={selected}
                    aria-readonly={!cell.editable}
                    aria-invalid={cell.invalid}
                    aria-colspan={cell.columnSpan}
                    aria-rowspan={cell.rowSpan}
                    className={cell.className}
                    title={cell.title}
                    {...cell.dataAttributes}
                    tabIndex={active ? 0 : -1}
                    style={cellPositionStyle(
                      rowHeaderWidth + columnMeasurement.start,
                      width,
                      height,
                      cell.style,
                      column,
                      rowHeaderWidth,
                      pinnedColumnOffsets
                    )}
                    onPointerDown={
                      interactionEventMode === "pointer"
                        ? (event) => handleCellDown(cell, event)
                        : undefined
                    }
                    onMouseDown={
                      interactionEventMode === "mouse"
                        ? (event) => handleCellDown(cell, event)
                        : undefined
                    }
                    onPointerEnter={
                      interactionEventMode === "pointer" ? () => handleCellEnter(cell) : undefined
                    }
                    onMouseEnter={
                      interactionEventMode === "mouse" ? () => handleCellEnter(cell) : undefined
                    }
                    onClick={(event) => {
                      if (event.detail === 0) {
                        interaction.onCellPointerDown(cell.ref, event.shiftKey);
                        interaction.onPointerUp();
                      }
                    }}
                    onContextMenu={(event) => onCellContextMenu?.(cell.ref, event)}
                    onDoubleClick={() => {
                      if (cell.editable) {
                        onInteraction({ type: "edit-start", cell: cell.ref, initialRawText: cell.displayValue });
                      } else {
                        onReadOnlyCellEditAttempt?.(cell.ref);
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
            headerHeight={headerHeight}
            getCell={getCell}
            getCellRect={virtualizer.getCellRect}
            pinnedColumnOffsets={pinnedColumnOffsets}
            pinnedRowOffsets={pinnedRowOffsets}
            scrollElement={rootRef.current}
            scrollTop={rootRef.current?.scrollTop ?? 0}
            scrollLeft={rootRef.current?.scrollLeft ?? 0}
            viewportHeight={rootRef.current?.clientHeight ?? 0}
            viewportWidth={rootRef.current?.clientWidth ?? 0}
            scale={scale}
            renderEditor={renderEditor}
            onInteraction={onInteraction}
          />
        ) : null}
        {renderOverlay?.()}
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
  headerHeight,
  getCell,
  getCellRect,
  pinnedColumnOffsets,
  pinnedRowOffsets,
  scrollElement,
  scrollTop,
  scrollLeft,
  viewportHeight,
  viewportWidth,
  scale,
  renderEditor,
  onInteraction
}: {
  editing: NonNullable<GridViewportProps["editing"]>;
  rows: readonly GridViewportRow[];
  columns: readonly GridViewportColumn[];
  rowHeaderWidth: number;
  headerHeight: number;
  getCell: GridViewportProps["getCell"];
  getCellRect(rowIndex: number, columnIndex: number): { top: number; left: number; width: number; height: number };
  pinnedColumnOffsets: ReadonlyMap<string, number>;
  pinnedRowOffsets: ReadonlyMap<string, number>;
  scrollElement: HTMLElement | null;
  scrollTop: number;
  scrollLeft: number;
  viewportHeight: number;
  viewportWidth: number;
  scale: number;
  renderEditor: GridViewportProps["renderEditor"];
  onInteraction: GridViewportProps["onInteraction"];
}) {
  useLayoutEffect(() => {
    if (!scrollElement) return;
    scrollElement.scrollTop = scrollTop;
    scrollElement.scrollLeft = scrollLeft;
  }, [editing.columnId, editing.rowId, scrollElement, scrollLeft, scrollTop]);
  const rowIndex = rows.findIndex((row) => row.id === editing.rowId);
  const columnIndex = columns.findIndex((column) => column.id === editing.columnId);
  if (rowIndex < 0 || columnIndex < 0) {
    return null;
  }
  const row = rows[rowIndex];
  const column = columns[columnIndex];
  const cell = getCell(row.id, column.id);
  const rect = getCellRect(rowIndex, columnIndex);
  const safeScale = Number.isFinite(scale) && scale > 0 ? scale : 1;
  const top = row.pinned === "top"
    ? scrollTop / safeScale + headerHeight + (pinnedRowOffsets.get(row.id) ?? 0)
    : row.pinned === "bottom"
      ? scrollTop / safeScale + viewportHeight / safeScale - rect.height - (pinnedRowOffsets.get(row.id) ?? 0)
      : headerHeight + rect.top;
  const left = column.pinned === "left"
    ? scrollLeft / safeScale + rowHeaderWidth + (pinnedColumnOffsets.get(column.id) ?? 0)
    : column.pinned === "right"
      ? scrollLeft / safeScale + viewportWidth / safeScale - rect.width - (pinnedColumnOffsets.get(column.id) ?? 0)
      : rowHeaderWidth + rect.left;
  const context = { row, column, cell, selected: true, active: true, editing };
  return (
    <div
      data-grid-editor-overlay="true"
      className={mergeClassNames("cell", "editing-cell", cell.className)}
      style={{
        ...cell.style,
        position: "absolute",
        zIndex: 8,
        top,
        left,
        width: cell.style?.width ?? rect.width,
        height: cell.style?.height ?? rect.height,
        minWidth: undefined,
        maxWidth: undefined,
        minHeight: undefined
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
  rows: ReadonlyMap<string, AxisMeasurement<string>>,
  columns: ReadonlyMap<string, AxisMeasurement<string>>
): SelectionBounds {
  if (!selection) {
    return null;
  }
  const anchorRow = rows.get(selection.anchor.rowId)?.index;
  const focusRow = rows.get(selection.focus.rowId)?.index;
  const anchorColumn = columns.get(selection.anchor.columnId)?.index;
  const focusColumn = columns.get(selection.focus.columnId)?.index;
  if (
    anchorRow === undefined ||
    focusRow === undefined ||
    anchorColumn === undefined ||
    focusColumn === undefined
  ) {
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

function headerCellStyle(left: number, width: number, height: number): CSSProperties {
  return {
    position: "absolute",
    top: 0,
    left,
    width,
    height,
    boxSizing: "border-box"
  };
}

function rowHeaderCellStyle(height: number, width: number): CSSProperties {
  return {
    position: "sticky",
    top: 0,
    left: 0,
    zIndex: 2,
    width,
    height,
    boxSizing: "border-box"
  };
}

function rowPositionStyle(
  measurement: AxisMeasurement<string>,
  row: GridViewportRow,
  width: number,
  headerHeight: number,
  offsets: ReadonlyMap<string, number>
): CSSProperties {
  const base: CSSProperties = {
    position: "absolute",
    top: headerHeight + measurement.start,
    left: 0,
    width,
    height: measurement.size
  };
  if (row.pinned === "top") {
    return { ...base, position: "sticky", top: headerHeight + (offsets.get(row.id) ?? 0), zIndex: 4 };
  }
  if (row.pinned === "bottom") {
    return { ...base, position: "sticky", top: undefined, bottom: offsets.get(row.id) ?? 0, zIndex: 4 };
  }
  return base;
}

function columnPositionStyle(
  base: CSSProperties,
  column: GridViewportColumn,
  rowHeaderWidth: number,
  offsets: ReadonlyMap<string, number>
): CSSProperties {
  if (column.pinned === "left") {
    return {
      ...base,
      position: "sticky",
      left: rowHeaderWidth + (offsets.get(column.id) ?? 0),
      zIndex: 4
    };
  }
  if (column.pinned === "right") {
    return { ...base, position: "sticky", left: undefined, right: offsets.get(column.id) ?? 0, zIndex: 4 };
  }
  return base;
}

function cellPositionStyle(
  left: number,
  width: number,
  height: number,
  style: CSSProperties | undefined,
  column: GridViewportColumn,
  rowHeaderWidth: number,
  offsets: ReadonlyMap<string, number>
): CSSProperties {
  return columnPositionStyle(
    {
      ...style,
      position: "absolute",
      top: 0,
      left,
      width: style?.width ?? width,
      height: style?.height ?? height,
      boxSizing: "border-box"
    },
    column,
    rowHeaderWidth,
    offsets
  );
}

function spanSize(
  measurements: readonly AxisMeasurement<string>[],
  index: number,
  span: number | undefined,
  fallback: number
): number {
  if (!span || span <= 1) {
    return fallback;
  }
  const first = measurements[index];
  const last = measurements[Math.min(measurements.length - 1, index + span - 1)];
  return first && last ? last.end - first.start : fallback;
}

function collectRenderedMeasurements(
  visible: readonly AxisMeasurement<string>[],
  measurementsById: ReadonlyMap<string, AxisMeasurement<string>>,
  retainedIds: ReadonlySet<string>
): AxisMeasurement<string>[] {
  const rendered = new Map<number, AxisMeasurement<string>>();
  for (const measurement of visible) {
    rendered.set(measurement.index, measurement);
  }
  for (const id of retainedIds) {
    const measurement = measurementsById.get(id);
    if (measurement) {
      rendered.set(measurement.index, measurement);
    }
  }
  return [...rendered.values()].sort((left, right) => left.start - right.start);
}

function createPinnedColumnOffsets(columns: readonly GridViewportColumn[]): ReadonlyMap<string, number> {
  const offsets = new Map<string, number>();
  let left = 0;
  for (const column of columns) {
    if (column.pinned === "left") {
      offsets.set(column.id, left);
      left += column.width;
    }
  }
  let right = 0;
  for (let index = columns.length - 1; index >= 0; index -= 1) {
    const column = columns[index];
    if (column.pinned === "right") {
      offsets.set(column.id, right);
      right += column.width;
    }
  }
  return offsets;
}

function createPinnedRowOffsets(rows: readonly GridViewportRow[]): ReadonlyMap<string, number> {
  const offsets = new Map<string, number>();
  let top = 0;
  for (const row of rows) {
    if (row.pinned === "top") {
      offsets.set(row.id, top);
      top += row.height;
    }
  }
  let bottom = 0;
  for (let index = rows.length - 1; index >= 0; index -= 1) {
    const row = rows[index];
    if (row.pinned === "bottom") {
      offsets.set(row.id, bottom);
      bottom += row.height;
    }
  }
  return offsets;
}

function mergeClassNames(...values: Array<string | undefined>): string {
  return [...new Set(values.flatMap((value) => value?.split(/\s+/).filter(Boolean) ?? []))].join(" ");
}

function isEditorEventTarget(target: EventTarget | null): boolean {
  return target instanceof Element && Boolean(target.closest(
    '[data-grid-editor-overlay="true"], input, select, textarea, button, [contenteditable="true"]'
  ));
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
