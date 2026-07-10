import {
  forwardRef,
  useCallback,
  useId,
  useImperativeHandle,
  useLayoutEffect,
  useMemo,
  useRef,
  useState,
  type DragEvent,
  type ForwardedRef,
  type ReactElement,
  type ReactNode,
  type RefAttributes
} from "react";
import type { FilterExpression, QueryRow, QueryScalar } from "../table/core/query";
import type {
  TableCellRef,
  TableDiagnosticEvent,
  TableIntent,
  TableSelection,
  TableSession,
  TableViewSnapshot
} from "../table/core/types";
import { GridViewport } from "./viewport/GridViewport";
import type {
  GridEditorState,
  GridViewportApi,
  GridViewportCell,
  GridViewportColumn,
  GridViewportInteraction,
  GridViewportRow
} from "./viewport/types";
import { DataTableCell, DataTableEditor, columnLabel } from "./DataTableCell";
import { DataTableColumnMenu } from "./DataTableColumnMenu";
import { DataTableExtensionBoundary } from "./DataTableExtensionBoundary";
import { DataTableToolbar } from "./DataTableToolbar";
import { exportArtifactToBlob } from "./exportArtifact";
import type {
  ColumnDef,
  DataTableHandle,
  DataTablePresentationProps,
  DataTableProps,
  HeaderRenderContext
} from "./tableTypes";
import { useTableSession } from "./useTableSession";
import { useTableSnapshot } from "./useTableSnapshot";

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

function OwnedLocalDataTable<TRow>({
  props,
  forwardedRef
}: {
  props: Extract<DataTableProps<TRow>, { session?: never }>;
  forwardedRef: ForwardedRef<DataTableHandle>;
}) {
  const session = useTableSession({
    source: {
      kind: "local",
      rows: props.rows,
      getRowId: props.getRowId,
      getSubRows: props.getSubRows,
      onRowsChange: props.onRowsChange,
      resetKey: props.resetKey
    },
    columns: props.columns,
    features: props.features,
    state: props.state,
    defaultState: props.defaultState,
    onStateChange: props.onStateChange,
    document: props.document,
    defaultDocument: props.defaultDocument,
    onDocumentChange: props.onDocumentChange,
    onDiagnostic: props.onDiagnostic
  });
  return <DataTableSurface session={session} presentation={props} forwardedRef={forwardedRef} />;
}

function SessionDataTable<TRow>({
  props,
  forwardedRef
}: {
  props: Extract<DataTableProps<TRow>, { session: TableSession<TRow, ColumnDef<TRow>> }>;
  forwardedRef: ForwardedRef<DataTableHandle>;
}) {
  return <DataTableSurface session={props.session} presentation={props} forwardedRef={forwardedRef} />;
}

function DataTableSurface<TRow>({
  session,
  presentation,
  forwardedRef
}: {
  session: TableSession<TRow, ColumnDef<TRow>>;
  presentation: DataTablePresentationProps;
  forwardedRef: ForwardedRef<DataTableHandle>;
}) {
  const snapshot = useTableSnapshot(session);
  const id = useId();
  const idPrefix = `js-spreadsheet-data-table-${id.replace(/[^A-Za-z0-9_-]/g, "")}`;
  const scrollRef = useRef<HTMLDivElement>(null);
  const viewportApiRef = useRef<GridViewportApi | null>(null);
  const draggedColumnId = useRef<string | null>(null);
  const [editing, setEditing] = useState<GridEditorState | null>(null);
  const [issue, setIssue] = useState("");
  const [announcement, setAnnouncement] = useState("");
  const [openColumnId, setOpenColumnId] = useState<string | null>(null);
  const [inlineFilterValues, setInlineFilterValues] = useState<Record<string, string>>({});
  const [containerInlineSize, setContainerInlineSize] = useState(0);
  const [measuredCellHeights, setMeasuredCellHeights] = useState<Record<string, Readonly<Record<string, number>>>>({});
  const measuredCellHeightsRef = useRef(measuredCellHeights);
  const pendingScrollAdjustmentRef = useRef(0);
  const inlineFilterBaseRef = useRef<FilterExpression | null>(snapshot.state.filter);
  const appliedInlineFilterRef = useRef<string | null>(null);
  const orderedColumns = useMemo(() => orderColumns(snapshot), [snapshot]);
  const visibleColumns = useMemo(
    () => orderedColumns.filter((column) => snapshot.state.columnVisibility[column.id] !== false),
    [orderedColumns, snapshot.state.columnVisibility]
  );
  const viewportColumns = useMemo(
    () => createViewportColumns(
      visibleColumns,
      orderedColumns,
      snapshot,
      presentation.layout ?? "fixed",
      containerInlineSize
    ),
    [containerInlineSize, orderedColumns, presentation.layout, snapshot, visibleColumns]
  );
  const measuredRowHeights = useMemo(() => Object.fromEntries(snapshot.rows.map((row) => [
    row.id,
    Math.max(28, ...viewportColumns.map((column) => measuredCellHeights[row.id]?.[column.id] ?? 28))
  ])), [measuredCellHeights, snapshot.rows, viewportColumns]);
  const viewportRows = useMemo(
    () => createViewportRows(snapshot, presentation.rowHeight ?? 32, measuredRowHeights),
    [measuredRowHeights, presentation.rowHeight, snapshot]
  );
  const rowsById = useMemo(() => new Map(snapshot.rows.map((row) => [row.id, row])), [snapshot.rows]);
  const columnsById = useMemo(() => new Map(snapshot.columns.map((column) => [column.id, column])), [snapshot.columns]);
  const selectedCells = useMemo(
    () => cellsInSelection(snapshot.selection, viewportRows, viewportColumns, rowsById),
    [rowsById, snapshot.selection, viewportColumns, viewportRows]
  );
  const ariaRowCount = tableAriaRowCount(snapshot);
  const activeRows = useMemo(
    () => snapshot.rows.filter((row): row is Extract<QueryRow<TRow>, { kind: "data" }> => row.kind === "data"),
    [snapshot.rows]
  );
  const headerActionRows = useMemo(() => activeRows.map((row) => row.original), [activeRows]);
  const defaultActiveCell = viewportRows[0] && viewportColumns[0]
    ? { rowId: viewportRows[0].id, columnId: viewportColumns[0].id }
    : null;

  const reportDiagnostic = useCallback((event: TableDiagnosticEvent) => {
    try {
      presentation.onDiagnostic?.(event);
    } catch {
      // Host diagnostics stay isolated from rendering.
    }
  }, [presentation.onDiagnostic]);

  const measureCell = useCallback((rowId: string, columnId: string, height: number) => {
    if (presentation.rowHeight !== "auto") return;
    const current = measuredCellHeightsRef.current;
    if (current[rowId]?.[columnId] === height) return;
    const columnIds = viewportColumns.map((column) => column.id);
    const previousHeight = measuredHeightForRow(current, rowId, columnIds);
    const next = { ...current, [rowId]: { ...current[rowId], [columnId]: height } };
    const nextHeight = measuredHeightForRow(next, rowId, columnIds);
    const rowIndex = snapshot.rows.findIndex((row) => row.id === rowId);
    const scroller = scrollRef.current;
    if (scroller && rowIndex >= 0 && nextHeight !== previousHeight) {
      const rowTop = 28 + snapshot.rows.slice(0, rowIndex).reduce(
        (top, row) => top + measuredHeightForRow(current, row.id, columnIds),
        0
      );
      if (rowTop + previousHeight <= scroller.scrollTop) {
        pendingScrollAdjustmentRef.current += nextHeight - previousHeight;
      }
    }
    measuredCellHeightsRef.current = next;
    setMeasuredCellHeights(next);
  }, [presentation.rowHeight, snapshot.rows, viewportColumns]);

  useLayoutEffect(() => {
    const adjustment = pendingScrollAdjustmentRef.current;
    if (adjustment !== 0 && scrollRef.current) {
      scrollRef.current.scrollTop += adjustment;
      pendingScrollAdjustmentRef.current = 0;
    }
  }, [measuredCellHeights]);

  useLayoutEffect(() => {
    if (presentation.layout !== "ratio") return;
    const element = scrollRef.current;
    if (!element) return;
    const publish = (width: number) => {
      if (Number.isFinite(width) && width > 0) setContainerInlineSize(width);
    };
    publish(element.clientWidth);
    if (typeof ResizeObserver === "undefined") return;
    const observer = new ResizeObserver((entries) => {
      const entry = entries[0];
      publish(entry?.contentRect.width ?? element.clientWidth);
    });
    observer.observe(element);
    return () => observer.disconnect();
  }, [presentation.layout]);

  async function run(intent: TableIntent<TRow>) {
    const result = await session.dispatch(intent);
    if (result.status === "rejected") {
      const message = result.issues?.[0]?.message ?? "Table operation was rejected";
      setIssue(message);
      setAnnouncement(message);
    } else if (result.status === "conflict") {
      setIssue("Table data changed; refresh and try again");
    }
    return result;
  }

  async function handleInteraction(interaction: GridViewportInteraction) {
    switch (interaction.type) {
      case "selection-change":
        await run({ type: "set-selection", selection: interaction.selection });
        return;
      case "edit-start": {
        const cell = snapshot.getCell(interaction.cell.rowId, interaction.cell.columnId);
        setEditing({
          ...interaction.cell,
          rawText: cell.formula ?? (cell.storedValue === null || cell.storedValue === undefined ? "" : String(cell.storedValue))
        });
        setIssue("");
        return;
      }
      case "edit-change":
        setEditing((current) => current ? { ...current, rawText: interaction.rawText } : current);
        return;
      case "edit-cancel":
        setEditing(null);
        setIssue("");
        return;
      case "edit-commit":
        await commitEdit(interaction.cell, interaction.rawText, interaction.move);
        return;
      case "copy":
        await copySelection(interaction.selection);
        return;
      case "paste":
        await pasteMatrix(interaction.text);
        return;
      case "column-resize":
        await run({ type: "resize-column", columnId: interaction.columnId, width: interaction.width });
    }
  }

  async function commitEdit(
    cell: TableCellRef,
    rawText: string,
    move?: "up" | "down" | "left" | "right"
  ) {
    const result = await run({ type: "edit-cells", edits: [{ ...cell, rawText }] });
    if (result.status !== "committed") {
      queueMicrotask(() => {
        const editor = scrollRef.current?.querySelector<HTMLElement>('[data-grid-editor-overlay="true"] input, [data-grid-editor-overlay="true"] select');
        editor?.focus();
      });
      return;
    }
    setEditing(null);
    setIssue("");
    if (move) await moveAfterCommit(cell, move);
  }

  async function moveAfterCommit(cell: TableCellRef, move: "up" | "down" | "left" | "right") {
    const rowIndex = viewportRows.findIndex((row) => row.id === cell.rowId);
    const columnIndex = viewportColumns.findIndex((column) => column.id === cell.columnId);
    if (rowIndex < 0 || columnIndex < 0) return;
    const rowDelta = move === "down" ? 1 : move === "up" ? -1 : 0;
    const columnDelta = move === "right" ? 1 : move === "left" ? -1 : 0;
    const nextRow = viewportRows[Math.max(0, Math.min(viewportRows.length - 1, rowIndex + rowDelta))];
    const nextColumn = viewportColumns[Math.max(0, Math.min(viewportColumns.length - 1, columnIndex + columnDelta))];
    if (!nextRow || !nextColumn) return;
    const focus = { rowId: nextRow.id, columnId: nextColumn.id };
    await run({ type: "set-selection", selection: { anchor: focus, focus } });
  }

  async function copySelection(selection: TableSelection) {
    const cells = cellsInSelection(selection, viewportRows, viewportColumns, rowsById);
    if (cells.length === 0) return;
    const rowIds = unique(cells.map((cell) => cell.rowId));
    const columnIds = unique(cells.map((cell) => cell.columnId));
    const text = rowIds.map((rowId) =>
      columnIds.map((columnId) => snapshot.getCell(rowId, columnId).displayValue).join("\t")
    ).join("\n");
    try {
      if (!navigator.clipboard || typeof navigator.clipboard.writeText !== "function") {
        throw new Error("Clipboard API is unavailable");
      }
      await navigator.clipboard.writeText(text);
      setAnnouncement(`Copied ${cells.length} cells`);
    } catch {
      setIssue("Could not copy selection");
    }
  }

  async function pasteMatrix(text: string) {
    const selection = snapshot.selection;
    if (!selection) return;
    const matrix = parseClipboardMatrix(text);
    if (matrix.length === 0) return;
    const startRow = viewportRows.findIndex((row) => row.id === selection.focus.rowId);
    const startColumn = viewportColumns.findIndex((column) => column.id === selection.focus.columnId);
    if (startRow < 0 || startColumn < 0) return;
    const edits: Array<{ rowId: string; columnId: string; rawText: string }> = [];
    matrix.forEach((values, rowOffset) => {
      const row = viewportRows[startRow + rowOffset];
      if (!row || rowsById.get(row.id)?.kind !== "data") return;
      values.forEach((rawText, columnOffset) => {
        const column = viewportColumns[startColumn + columnOffset];
        if (column) edits.push({ rowId: row.id, columnId: column.id, rawText });
      });
    });
    if (edits.length === 0) return;
    const result = await run({ type: "edit-cells", edits });
    if (result.status === "committed") setAnnouncement(`Pasted ${edits.length} cells`);
  }

  async function toggleRowSelection(rowId: string, selected: boolean) {
    const current = snapshot.state.selectedRowIds;
    const rowIds = presentation.rowSelection === "single"
      ? selected ? [rowId] : []
      : selected ? [...new Set([...current, rowId])] : current.filter((id) => id !== rowId);
    const result = await run({ type: "set-row-selection", rowIds });
    if (result.status === "committed") {
      try {
        presentation.onRowSelectionChange?.(session.getSnapshot().state.selectedRowIds);
      } catch {
        reportDiagnostic({ category: "extension", metadata: { slot: "row-selection" } });
      }
    }
  }

  async function toggleAllRows(selected: boolean) {
    const rowIds = selected ? activeRows.map((row) => row.id) : [];
    const result = await run({ type: "set-row-selection", rowIds });
    if (result.status === "committed") {
      try {
        presentation.onRowSelectionChange?.(session.getSnapshot().state.selectedRowIds);
      } catch {
        reportDiagnostic({ category: "extension", metadata: { slot: "row-selection" } });
      }
    }
  }

  function reorderColumn(sourceId: string, targetId: string) {
    const order = orderedColumns.map((column) => column.id);
    const sourceIndex = order.indexOf(sourceId);
    const targetIndex = order.indexOf(targetId);
    if (sourceIndex < 0 || targetIndex < 0 || sourceIndex === targetIndex) return;
    order.splice(sourceIndex, 1);
    const insertionIndex = sourceIndex < targetIndex ? targetIndex - 1 : targetIndex;
    order.splice(insertionIndex, 0, sourceId);
    void run({ type: "set-column-order", columnIds: order });
    setAnnouncement(`${columnLabel(columnsById.get(sourceId)!)} moved to position ${insertionIndex + 1}`);
  }

  function updateInlineFilter(column: ColumnDef<TRow>, value: string) {
    const currentFilterKey = JSON.stringify(snapshot.state.filter);
    if (appliedInlineFilterRef.current !== currentFilterKey) {
      inlineFilterBaseRef.current = snapshot.state.filter;
    }
    const next = { ...inlineFilterValues, [column.id]: value };
    setInlineFilterValues(next);
    const expressions = visibleColumns.flatMap((candidate) => {
      const raw = next[candidate.id]?.trim();
      if (!raw || !candidate.filterable) return [];
      return [{
        kind: "comparison" as const,
        columnId: candidate.id,
        operator: candidate.dataType === "text" || candidate.dataType === undefined ? "contains" as const : "eq" as const,
        value: scalarFor(candidate, raw)
      }];
    });
    const filter = combineFilters(inlineFilterBaseRef.current, expressions);
    appliedInlineFilterRef.current = JSON.stringify(filter);
    void run({ type: "set-filter", filter });
  }

  useImperativeHandle(forwardedRef, (): DataTableHandle => ({
    focus() {
      const active = scrollRef.current?.querySelector<HTMLElement>('[role="gridcell"][tabindex="0"]');
      (active ?? scrollRef.current)?.focus();
    },
    refresh() {
      return session.refresh();
    },
    dispatch(intent) {
      return session.dispatch(intent as TableIntent<TRow>);
    },
    undo() {
      return session.undo();
    },
    redo() {
      return session.redo();
    },
    scrollToRow(rowId) {
      if (session.getSnapshot().getRowIndex(rowId) < 0) throw new Error(`Unknown row id: ${rowId}`);
      const firstColumn = visibleColumns[0];
      if (firstColumn) viewportApiRef.current?.ensureCellVisible(rowId, firstColumn.id);
    },
    getSelection() {
      return session.getSnapshot().selection;
    },
    async export(options) {
      return exportArtifactToBlob(await session.export(options));
    }
  }), [session, visibleColumns]);

  const layout = presentation.layout ?? "fixed";
  return (
    <div
      className={["js-spreadsheet-data-table-surface", presentation.className].filter(Boolean).join(" ")}
      data-table-layout={layout}
      data-row-height-mode={presentation.rowHeight === "auto" ? "auto" : "fixed"}
      style={presentation.style}
    >
      <DataTableToolbar
        session={session}
        snapshot={snapshot}
        selectedCells={selectedCells}
        onShowColumn={(columnId) => void run({ type: "set-column-visibility", columnId, visible: true })}
        onIssue={setIssue}
      />
      {presentation.inlineFilters ? (
        <div className="js-spreadsheet-data-table__inline-filters" aria-label="Inline filters">
          {visibleColumns.filter((column) => column.filterable).map((column) => (
            <label key={column.id}>
              <span>{columnLabel(column)}</span>
              <InlineFilterControl
                column={column}
                value={inlineFilterValues[column.id] ?? ""}
                onChange={(value) => updateInlineFilter(column, value)}
              />
            </label>
          ))}
        </div>
      ) : null}
      <GridViewport
        idPrefix={idPrefix}
        ariaLabel={presentation["aria-label"] ?? "Data table"}
        rows={viewportRows}
        ariaRowCount={ariaRowCount}
        columns={viewportColumns}
        ariaColumnCount={orderedColumns.length + 1}
        getCell={(rowId, columnId) => createViewportCell(snapshot, rowId, columnId, columnsById)}
        selection={snapshot.selection}
        activeCell={snapshot.selection?.focus ?? defaultActiveCell}
        editing={editing}
        onInteraction={(interaction) => void handleInteraction(interaction)}
        scrollRef={scrollRef}
        onRegisterApi={(api) => { viewportApiRef.current = api; }}
        onColumnHeaderDragOver={(_column, event) => event.preventDefault()}
        onColumnHeaderDrop={(column, event) => {
          event.preventDefault();
          if (draggedColumnId.current) reorderColumn(draggedColumnId.current, column.id);
          draggedColumnId.current = null;
        }}
        retainedRowIds={editing ? [editing.rowId] : []}
        retainedColumnIds={editing ? [editing.columnId] : []}
        rootClassName="js-spreadsheet-data-table__viewport"
        canvasClassName="js-spreadsheet-data-table__canvas"
        announce={announcement}
        renderCell={({ row, column }) => {
          const queryRow = rowsById.get(row.id)!;
          const columnDef = columnsById.get(column.id)!;
          return (
            <DataTableCell
              row={queryRow}
              column={columnDef}
              cell={snapshot.getCell(row.id, column.id)}
              isFirstColumn={column.id === visibleColumns[0]?.id}
              onToggleExpanded={(rowId, expanded) => void run({ type: "set-row-expanded", rowId, expanded })}
              onMeasure={presentation.rowHeight === "auto" ? measureCell : undefined}
              onDiagnostic={reportDiagnostic}
            />
          );
        }}
        renderEditor={({ row, column, editing: editorState }) => {
          const queryRow = rowsById.get(row.id);
          const columnDef = columnsById.get(column.id);
          if (!queryRow || queryRow.kind !== "data" || !columnDef) return null;
          return (
            <DataTableEditor
              row={queryRow}
              column={columnDef}
              cell={snapshot.getCell(row.id, column.id)}
              rawText={editorState.rawText}
              onChange={(rawText) => setEditing((current) => current ? { ...current, rawText } : current)}
              onCommit={(move) => void commitEdit({ rowId: row.id, columnId: column.id }, editorState.rawText, move)}
              onCancel={() => setEditing(null)}
              onDiagnostic={reportDiagnostic}
            />
          );
        }}
        renderColumnHeader={(viewportColumn) => {
          const column = columnsById.get(viewportColumn.id)!;
          const label = columnLabel(column);
          return (
            <div
              className="js-spreadsheet-data-table__column-header"
              onDragOver={(event) => event.preventDefault()}
              onDrop={(event) => {
                event.preventDefault();
                if (draggedColumnId.current) reorderColumn(draggedColumnId.current, column.id);
                draggedColumnId.current = null;
              }}
            >
              {typeof column.header === "function" ? (
                <DataTableExtensionBoundary columnId={column.id} slot="header" onDiagnostic={reportDiagnostic}>
                  <InvokeHeaderRenderer render={column.header} context={{ column }} />
                </DataTableExtensionBoundary>
              ) : <span>{column.header}</span>}
              <button
                type="button"
                draggable
                aria-label={`Reorder ${label}`}
                onDragStart={(event: DragEvent<HTMLButtonElement>) => {
                  draggedColumnId.current = column.id;
                  event.dataTransfer?.setData("text/plain", column.id);
                }}
              >↕</button>
              <button type="button" aria-label={`Column options for ${label}`} aria-expanded={openColumnId === column.id} onClick={() => setOpenColumnId(openColumnId === column.id ? null : column.id)}>
                ⋯
              </button>
              {openColumnId === column.id ? (
                <DataTableColumnMenu
                  column={column}
                  session={session}
                  snapshot={snapshot}
                  rows={headerActionRows}
                  onIssue={setIssue}
                  onDiagnostic={reportDiagnostic}
                  onAnnouncement={setAnnouncement}
                />
              ) : null}
            </div>
          );
        }}
        renderRowHeader={(row) => {
          const queryRow = rowsById.get(row.id);
          return (
            <span>
              {presentation.rowSelection && presentation.rowSelection !== "none" && queryRow?.kind === "data" ? (
                <input
                  type="checkbox"
                  aria-label={`Select row ${row.id}`}
                  checked={snapshot.state.selectedRowIds.includes(row.id)}
                  onChange={(event) => void toggleRowSelection(row.id, event.currentTarget.checked)}
                />
              ) : null}
              {row.label}
            </span>
          );
        }}
        renderCornerHeader={presentation.rowSelection === "multiple" ? () => (
          <input
            type="checkbox"
            aria-label="Select all rows"
            checked={activeRows.length > 0 && activeRows.every((row) => snapshot.state.selectedRowIds.includes(row.id))}
            onChange={(event) => void toggleAllRows(event.currentTarget.checked)}
          />
        ) : undefined}
      />
      {snapshot.rows.length === 0 ? (
        <div className="js-spreadsheet-data-table__empty">{presentation.noDataMessage ?? "No rows"}</div>
      ) : null}
      {issue ? <div role="alert" className="js-spreadsheet-data-table__issue">{issue}</div> : null}
    </div>
  );
}

function InvokeHeaderRenderer<TRow>({
  render,
  context
}: {
  render: (context: HeaderRenderContext<TRow, unknown>) => ReactNode;
  context: HeaderRenderContext<TRow, unknown>;
}) {
  return <>{render(context)}</>;
}

function orderColumns<TRow>(snapshot: TableViewSnapshot<TRow, ColumnDef<TRow>>): ColumnDef<TRow>[] {
  const byId = new Map(snapshot.columns.map((column) => [column.id, column]));
  const base = [
    ...snapshot.state.columnOrder.map((id) => byId.get(id)).filter((column): column is ColumnDef<TRow> => Boolean(column)),
    ...snapshot.columns.filter((column) => !snapshot.state.columnOrder.includes(column.id))
  ];
  const left = new Set(snapshot.state.columnPinning.left);
  const right = new Set(snapshot.state.columnPinning.right);
  return [
    ...base.filter((column) => left.has(column.id)),
    ...base.filter((column) => !left.has(column.id) && !right.has(column.id)),
    ...base.filter((column) => right.has(column.id))
  ];
}

function createViewportColumns<TRow>(
  columns: readonly ColumnDef<TRow>[],
  orderedColumns: readonly ColumnDef<TRow>[],
  snapshot: TableViewSnapshot<TRow, ColumnDef<TRow>>,
  layout: "fixed" | "fluid" | "ratio",
  containerInlineSize: number
): GridViewportColumn[] {
  const totalWeight = columns.reduce((sum, column) => sum + Math.max(1, column.width ?? 1), 0);
  const ratioWidth = containerInlineSize > 56 ? containerInlineSize - 56 : totalWeight;
  return columns.map((column) => {
    const minWidth = column.minWidth ?? 56;
    const maxWidth = column.maxWidth ?? 1200;
    const storedWidth = snapshot.state.columnWidths[column.id];
    const requested = storedWidth ?? column.width ?? (layout === "fluid" ? 160 : 120);
    const width = layout === "ratio" && storedWidth === undefined
      ? ratioWidth * Math.max(1, column.width ?? 1) / Math.max(1, totalWeight)
      : requested;
    const pin = snapshot.state.columnPinning.left.includes(column.id)
      ? "left"
      : snapshot.state.columnPinning.right.includes(column.id)
        ? "right"
        : undefined;
    return {
      id: column.id,
      label: columnLabel(column),
      width: Math.min(maxWidth, Math.max(minWidth, width)),
      minWidth,
      maxWidth,
      ariaColumnIndex: orderedColumns.findIndex((candidate) => candidate.id === column.id) + 2,
      ...(pin ? { pinned: pin } : {})
    };
  });
}

function createViewportRows<TRow>(
  snapshot: TableViewSnapshot<TRow, ColumnDef<TRow>>,
  rowHeight: number | "auto",
  measured: Readonly<Record<string, number>>
): GridViewportRow[] {
  const offset = snapshot.pageInfo.kind === "offset" ? snapshot.pageInfo.offset : 0;
  let dataPosition = 0;
  let aggregatePosition = 0;
  return snapshot.rows.map((row, index) => {
    const number = offset + dataPosition + 1;
    const ariaRowIndex = snapshot.state.grouping.length > 0
      ? index + 2
      : row.kind === "data"
        ? offset + dataPosition++ + 2
      : row.kind === "aggregate" && snapshot.totalRowCount.kind === "known"
        ? snapshot.totalRowCount.value + aggregatePosition++ + 2
        : index + 2;
    const label = row.kind === "data"
      ? String(number)
      : row.kind === "group"
        ? `${formatScalar(row.key)} (${row.count})`
        : "Totals";
    return {
      id: row.id,
      label,
      ariaLabel: row.kind === "data" ? `Row ${number}` : label,
      headerAriaLabel: row.kind === "data" ? `Row ${number}` : label,
      height: rowHeight === "auto" ? measured[row.id] ?? 32 : rowHeight,
      kind: row.kind,
      ariaRowIndex
    };
  });
}

function tableAriaRowCount<TRow>(snapshot: TableViewSnapshot<TRow, ColumnDef<TRow>>): number {
  if (snapshot.totalRowCount.kind === "unknown") return -1;
  if (snapshot.state.grouping.length > 0) {
    return snapshot.completeness === "completeDataset" ? snapshot.rows.length + 1 : -1;
  }
  const aggregateRows = snapshot.rows.filter((row) => row.kind === "aggregate").length;
  return snapshot.totalRowCount.value + aggregateRows + 1;
}

function createViewportCell<TRow>(
  snapshot: TableViewSnapshot<TRow, ColumnDef<TRow>>,
  rowId: string,
  columnId: string,
  columnsById: ReadonlyMap<string, ColumnDef<TRow>>
): GridViewportCell {
  const cell = snapshot.getCell(rowId, columnId);
  const format = cell.metadata.format;
  const column = columnsById.get(columnId)!;
  return {
    ref: { rowId, columnId },
    ariaLabel: `${rowId} ${columnLabel(column)}`,
    displayValue: cell.displayValue,
    editable: cell.editable,
    invalid: cell.issues.length > 0,
    className: cell.metadata.readOnly ? "js-spreadsheet-data-table__cell--readonly" : undefined,
    style: format ? {
      color: format.textColor,
      backgroundColor: format.backgroundColor,
      fontWeight: format.bold ? 700 : undefined,
      fontStyle: format.italic ? "italic" : undefined,
      fontFamily: format.fontFamily,
      fontSize: format.fontSize,
      textAlign: format.horizontalAlign,
      verticalAlign: format.verticalAlign,
      whiteSpace: format.wrapText ? "normal" : undefined
    } : undefined
  };
}

function cellsInSelection<TRow>(
  selection: TableSelection | null,
  rows: readonly GridViewportRow[],
  columns: readonly GridViewportColumn[],
  rowsById: ReadonlyMap<string, QueryRow<TRow>>
): TableCellRef[] {
  if (!selection) return [];
  const rowIndexes = [rows.findIndex((row) => row.id === selection.anchor.rowId), rows.findIndex((row) => row.id === selection.focus.rowId)];
  const columnIndexes = [
    columns.findIndex((column) => column.id === selection.anchor.columnId),
    columns.findIndex((column) => column.id === selection.focus.columnId)
  ];
  if ([...rowIndexes, ...columnIndexes].some((index) => index < 0)) return [];
  const result: TableCellRef[] = [];
  for (let rowIndex = Math.min(...rowIndexes); rowIndex <= Math.max(...rowIndexes); rowIndex += 1) {
    const row = rows[rowIndex];
    if (rowsById.get(row.id)?.kind !== "data") continue;
    for (let columnIndex = Math.min(...columnIndexes); columnIndex <= Math.max(...columnIndexes); columnIndex += 1) {
      result.push({ rowId: row.id, columnId: columns[columnIndex].id });
    }
  }
  return result;
}

function parseClipboardMatrix(text: string): string[][] {
  const normalized = text.replace(/\r\n?/g, "\n");
  const lines = normalized.split("\n");
  if (lines.at(-1) === "") lines.pop();
  return lines.filter((line, index) => line.length > 0 || index < lines.length - 1).map((line) => line.split("\t"));
}

function scalarFor<TRow>(column: ColumnDef<TRow>, raw: string): QueryScalar {
  if (column.dataType === "number") return { type: "number", value: Number(raw) };
  if (column.dataType === "boolean") return { type: "boolean", value: raw === "true" };
  if (column.dataType === "date") return { type: "date", value: raw };
  if (column.dataType === "datetime") return { type: "datetime", value: raw };
  return { type: "string", value: raw };
}

function formatScalar(value: QueryScalar): string {
  return value.type === "null" ? "Blank" : String(value.value);
}

function unique(values: readonly string[]): string[] {
  return [...new Set(values)];
}

function InlineFilterControl<TRow>({
  column,
  value,
  onChange
}: {
  column: ColumnDef<TRow>;
  value: string;
  onChange(value: string): void;
}) {
  const label = `Inline filter ${columnLabel(column)}`;
  if (column.dataType === "boolean") {
    return (
      <select aria-label={label} value={value} onChange={(event) => onChange(event.currentTarget.value)}>
        <option value="">Any</option>
        <option value="true">True</option>
        <option value="false">False</option>
      </select>
    );
  }
  const type = column.dataType === "number"
    ? "number"
    : column.dataType === "date"
      ? "date"
      : column.dataType === "datetime"
        ? "datetime-local"
        : "text";
  return <input type={type} aria-label={label} value={value} onChange={(event) => onChange(event.currentTarget.value)} />;
}

function measuredHeightForRow(
  measurements: Readonly<Record<string, Readonly<Record<string, number>>>>,
  rowId: string,
  columnIds: readonly string[]
): number {
  return Math.max(28, ...columnIds.map((columnId) => measurements[rowId]?.[columnId] ?? 28));
}

function combineFilters(
  base: FilterExpression | null,
  inline: readonly FilterExpression[]
): FilterExpression | null {
  const operands = [...(base ? [base] : []), ...inline];
  if (operands.length === 0) return null;
  if (operands.length === 1) return operands[0];
  return { kind: "logical", operator: "and", operands };
}
