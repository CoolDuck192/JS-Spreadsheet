import {
  useCallback,
  useLayoutEffect,
  useMemo,
  useRef,
  type KeyboardEvent as ReactKeyboardEvent,
  type RefObject
} from "react";
import type { TableCellRef, TableSelection } from "../../table/core/types";
import {
  createGridInteractionState,
  reduceGridInteraction,
  type GridInteractionModel
} from "./gridInteraction";
import type {
  GridEditorState,
  GridViewportColumn,
  GridViewportInteraction,
  GridViewportRow
} from "./types";

export type UseGridInteractionOptions = {
  idPrefix: string;
  rows: readonly GridViewportRow[];
  columns: readonly GridViewportColumn[];
  selection: TableSelection | null;
  activeCell?: TableCellRef | null;
  editing: GridEditorState | null;
  onInteraction(interaction: GridViewportInteraction): void;
  ensureCellVisible(rowIndex: number, columnIndex: number): void;
  rootRef: RefObject<HTMLDivElement | null>;
  getInitialRawText?(cell: TableCellRef): string;
  isCellEditable?(cell: TableCellRef): boolean;
};

export type GridInteractionBindings = {
  onKeyDown(event: ReactKeyboardEvent<HTMLDivElement>): void;
  requestFocusRestoration(ensureVisible?: boolean): void;
  onCellPointerDown(cell: TableCellRef, extend: boolean): void;
  onCellPointerEnter(cell: TableCellRef): void;
  onPointerUp(): void;
  activeDescendantId: string | undefined;
};

export function useGridInteraction({
  idPrefix,
  rows,
  columns,
  selection,
  activeCell: controlledActiveCell,
  editing,
  onInteraction,
  ensureCellVisible,
  rootRef,
  getInitialRawText = () => "",
  isCellEditable = () => true
}: UseGridInteractionOptions): GridInteractionBindings {
  const rowIds = useMemo(() => rows.map((row) => row.id), [rows]);
  const columnIds = useMemo(() => columns.map((column) => column.id), [columns]);
  const rowIndexById = useMemo(() => new Map(rowIds.map((id, index) => [id, index])), [rowIds]);
  const columnIndexById = useMemo(() => new Map(columnIds.map((id, index) => [id, index])), [columnIds]);
  const model = useMemo<GridInteractionModel>(() => ({ rowIds, columnIds }), [columnIds, rowIds]);
  const draggingRef = useRef(false);
  const dragAnchorRef = useRef<TableCellRef | null>(null);
  const focusPendingRef = useRef(false);
  const ensureFocusVisibleRef = useRef(false);
  const previousEditingRef = useRef<GridEditorState | null>(editing);
  const activeCell = resolveActiveCell(
    controlledActiveCell ?? selection?.focus ?? null,
    rowIndexById,
    columnIndexById
  );
  const activeDescendantId = activeCell ? gridCellDomId(idPrefix, activeCell) : undefined;

  const requestFocusRestoration = useCallback((ensureVisible = true) => {
    focusPendingRef.current = true;
    ensureFocusVisibleRef.current = ensureVisible;
  }, []);

  const emitSelection = useCallback(
    (nextSelection: TableSelection, ensureVisible = true) => {
      requestFocusRestoration(ensureVisible);
      const rowIndex = rowIndexById.get(nextSelection.focus.rowId);
      const columnIndex = columnIndexById.get(nextSelection.focus.columnId);
      if (ensureVisible && rowIndex !== undefined && columnIndex !== undefined) {
        ensureCellVisible(rowIndex, columnIndex);
      }
      const targetId = gridCellDomId(idPrefix, nextSelection.focus);
      if (!rootRef.current?.querySelector(`[id="${targetId}"]`)) {
        rootRef.current?.focus({ preventScroll: true });
      }
      onInteraction({ type: "selection-change", selection: nextSelection });
    },
    [columnIndexById, ensureCellVisible, idPrefix, onInteraction, requestFocusRestoration, rootRef, rowIndexById]
  );

  const moveSelection = useCallback(
    (rowDelta: number, columnDelta: number, extend: boolean) => {
      const initialCell = activeCell ?? firstCell(rowIds, columnIds);
      if (!initialCell) {
        return;
      }
      const state = selection
        ? { selection, editing, dragging: false }
        : createGridInteractionState(initialCell);
      const next = reduceGridInteraction(
        state,
        { type: "move", rowDelta, columnDelta, extend },
        model
      );
      emitSelection(next.selection);
    },
    [activeCell, columnIds, editing, emitSelection, model, rowIds, selection]
  );

  const startEdit = useCallback(
    (initialRawText: string) => {
      if (!activeCell || !isCellEditable(activeCell)) {
        return;
      }
      onInteraction({ type: "edit-start", cell: activeCell, initialRawText });
    },
    [activeCell, isCellEditable, onInteraction]
  );

  const onKeyDown = useCallback(
    (event: ReactKeyboardEvent<HTMLDivElement>) => {
      if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "c") {
        if (selection) {
          event.preventDefault();
          onInteraction({ type: "copy", selection });
        }
        return;
      }
      if (event.ctrlKey || event.metaKey || event.altKey) {
        return;
      }
      const pageSize = Math.max(
        1,
        Math.floor((rootRef.current?.clientHeight || 280) / averageRowHeight(rows))
      );
      switch (event.key) {
        case "ArrowUp":
          event.preventDefault();
          moveSelection(-1, 0, event.shiftKey);
          return;
        case "ArrowDown":
          event.preventDefault();
          moveSelection(1, 0, event.shiftKey);
          return;
        case "ArrowLeft":
          event.preventDefault();
          moveSelection(0, -1, event.shiftKey);
          return;
        case "ArrowRight":
          event.preventDefault();
          moveSelection(0, 1, event.shiftKey);
          return;
        case "Home":
          event.preventDefault();
          moveSelection(0, -columnIds.length, event.shiftKey);
          return;
        case "End":
          event.preventDefault();
          moveSelection(0, columnIds.length, event.shiftKey);
          return;
        case "PageUp":
          event.preventDefault();
          moveSelection(-pageSize, 0, event.shiftKey);
          return;
        case "PageDown":
          event.preventDefault();
          moveSelection(pageSize, 0, event.shiftKey);
          return;
        case "Tab":
          event.preventDefault();
          moveSelection(0, event.shiftKey ? -1 : 1, false);
          return;
        case "Enter":
          event.preventDefault();
          moveSelection(event.shiftKey ? -1 : 1, 0, false);
          return;
        case "F2":
          event.preventDefault();
          if (activeCell) {
            startEdit(getInitialRawText(activeCell));
          }
          return;
        case "Escape":
          if (editing) {
            event.preventDefault();
            onInteraction({ type: "edit-cancel" });
          }
          return;
        default:
          if (isPrintableKey(event)) {
            event.preventDefault();
            startEdit(event.key);
          }
      }
    },
    [activeCell, columnIds.length, editing, getInitialRawText, moveSelection, onInteraction, rootRef, rows, selection, startEdit]
  );

  const onCellPointerDown = useCallback(
    (cell: TableCellRef, extend: boolean) => {
      const initialCell = activeCell ?? cell;
      const state = selection
        ? { selection, editing, dragging: false }
        : createGridInteractionState(initialCell);
      const next = reduceGridInteraction(state, { type: "pointer-down", cell, extend }, model);
      draggingRef.current = true;
      dragAnchorRef.current = next.selection.anchor;
      emitSelection(next.selection, false);
    },
    [activeCell, editing, emitSelection, model, selection]
  );

  const onCellPointerEnter = useCallback(
    (cell: TableCellRef) => {
      const anchor = dragAnchorRef.current;
      if (!draggingRef.current || !anchor) {
        return;
      }
      emitSelection({ anchor, focus: { ...cell } }, false);
    },
    [emitSelection]
  );

  const onPointerUp = useCallback(() => {
    draggingRef.current = false;
    dragAnchorRef.current = null;
  }, []);

  useLayoutEffect(() => {
    window.addEventListener("pointerup", onPointerUp);
    window.addEventListener("pointercancel", onPointerUp);
    window.addEventListener("mouseup", onPointerUp);
    return () => {
      window.removeEventListener("pointerup", onPointerUp);
      window.removeEventListener("pointercancel", onPointerUp);
      window.removeEventListener("mouseup", onPointerUp);
    };
  }, [onPointerUp]);

  useLayoutEffect(() => {
    if (previousEditingRef.current && !editing) {
      focusPendingRef.current = true;
      ensureFocusVisibleRef.current = true;
    }
    previousEditingRef.current = editing;
    if (editing) {
      return;
    }
    if (!focusPendingRef.current || !activeCell) {
      return;
    }
    const rowIndex = rowIndexById.get(activeCell.rowId);
    const columnIndex = columnIndexById.get(activeCell.columnId);
    if (ensureFocusVisibleRef.current && rowIndex !== undefined && columnIndex !== undefined) {
      ensureCellVisible(rowIndex, columnIndex);
    }
    const activeId = gridCellDomId(idPrefix, activeCell);
    const activeElement = rootRef.current?.querySelector(`[id="${activeId}"]`);
    if (activeElement instanceof HTMLElement) {
      activeElement.focus({ preventScroll: true });
      focusPendingRef.current = false;
      ensureFocusVisibleRef.current = false;
    } else {
      rootRef.current?.focus({ preventScroll: true });
    }
  });

  return {
    onKeyDown,
    requestFocusRestoration,
    onCellPointerDown,
    onCellPointerEnter,
    onPointerUp,
    activeDescendantId
  };
}

export function gridCellDomId(idPrefix: string, cell: TableCellRef): string {
  return `${safePrefix(idPrefix)}-cell-${encodeIdPart(cell.rowId)}-${encodeIdPart(cell.columnId)}`;
}

export function gridLiveRegionDomId(idPrefix: string): string {
  return `${safePrefix(idPrefix)}-live`;
}

function resolveActiveCell(
  candidate: TableCellRef | null,
  rowIndexById: ReadonlyMap<string, number>,
  columnIndexById: ReadonlyMap<string, number>
): TableCellRef | null {
  if (candidate && rowIndexById.has(candidate.rowId) && columnIndexById.has(candidate.columnId)) {
    return candidate;
  }
  return null;
}

function firstCell(rowIds: readonly string[], columnIds: readonly string[]): TableCellRef | null {
  return rowIds[0] && columnIds[0] ? { rowId: rowIds[0], columnId: columnIds[0] } : null;
}

function averageRowHeight(rows: readonly GridViewportRow[]): number {
  if (rows.length === 0) {
    return 28;
  }
  return rows.reduce((total, row) => total + row.height, 0) / rows.length;
}

function isPrintableKey(event: ReactKeyboardEvent): boolean {
  return event.key.length === 1 && !event.ctrlKey && !event.metaKey && !event.altKey;
}

function safePrefix(prefix: string): string {
  const safe = prefix.replace(/[^A-Za-z0-9_-]/g, "-");
  return `${safe || "grid"}-${encodeIdPart(prefix)}`;
}

function encodeIdPart(value: string): string {
  if (value === "") {
    return "empty";
  }
  return Array.from(value, (character) => character.codePointAt(0)!.toString(36)).join("_");
}
