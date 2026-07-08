import { memo, useCallback, useEffect, useMemo, useRef, useState, type CSSProperties, type RefObject } from "react";
import { ChevronDown, ListFilter } from "lucide-react";
import { FormulaSuggestions, formulaSuggestionOptionId } from "./FormulaSuggestions";
import type { CellFormat, CellRange, ConditionalFormatRule, DataValidationRule, SheetFilter, SheetModel } from "../types";
import { columnIndexToName, formatCellAddress, getRangeAddresses, normalizeRange } from "../lib/addressing";
import { getConditionalDataBarForValue, getConditionalFormatForValue } from "../lib/conditionalFormatting";
import { formatDisplayValue } from "../lib/displayFormat";
import { getVisibleRows } from "../lib/filters";
import type { FormulaEngine } from "../lib/formulaEngine";
import { getFormulaSuggestions, insertFormulaSuggestion } from "../lib/formulaSuggestions";
import {
  DEFAULT_COLUMN_WIDTH,
  DEFAULT_ROW_HEIGHT,
  clampColumnWidth,
  clampRowHeight
} from "../lib/sheetDimensions";
import { validateCellValue } from "../lib/validation";

const DEFAULT_VIEWPORT_HEIGHT = 560;
const ROW_OVERSCAN = 16;
const ROW_HEADER_WIDTH = 48;
const COLUMN_HEADER_HEIGHT = 28;
const AUTO_SCROLL_MAX_STEP = 48;

export type CommitEditMove = "down" | "up" | "right" | "left";

export type GridScrollApi = {
  ensureCellVisible: (row: number, column: number) => void;
};

type GridProps = {
  sheet: SheetModel;
  formulaEngine: FormulaEngine;
  selection: CellRange;
  editingCell: { address: string; value: string } | null;
  copiedRange?: CellRange | null;
  zoomLevel?: number;
  showGridlines?: boolean;
  showHeaders?: boolean;
  showFormulas?: boolean;
  freezeTopRow?: boolean;
  freezeFirstColumn?: boolean;
  getCellFormat: (address: string) => CellFormat | undefined;
  getCellComment?: (address: string) => string | null | undefined;
  getCellHyperlink?: (address: string) => string | null | undefined;
  getCellReadOnly?: (address: string) => boolean;
  getCellValidation?: (address: string) => DataValidationRule | null | undefined;
  getCellConditionalFormatRules?: (address: string) => ConditionalFormatRule[];
  scrollRef?: RefObject<HTMLDivElement | null>;
  onSelectionChange: (range: CellRange) => void;
  onStartEdit: (address: string) => void;
  onEditValueChange: (value: string) => void;
  onCommitEdit: (address: string, value: string, move?: CommitEditMove) => void;
  onCancelEdit: () => void;
  onPasteText: (text: string) => void;
  onKeyCommand: (event: React.KeyboardEvent<HTMLDivElement>) => void;
  onAutoFill?: (sourceRange: CellRange, targetRange: CellRange) => void;
  onAutoFillDoubleClick?: () => void;
  onCellContextMenu?: (event: { address: string; row: number; column: number; x: number; y: number }) => void;
  onAutoFilterColumn?: (column: number, values: string[]) => void;
  onClearAutoFilterColumn?: (column: number) => void;
  onSortAutoFilterColumn?: (column: number, direction: "asc" | "desc") => void;
  onColumnResize?: (column: number, width: number) => void;
  onRowResize?: (row: number, height: number) => void;
  onColumnAutoFit?: (column: number) => void;
  onRowAutoFit?: (row: number) => void;
  onRegisterScrollApi?: (api: GridScrollApi) => void;
};

type ResizeDraft =
  | {
      kind: "column";
      index: number;
      startClient: number;
      startSize: number;
      size: number;
    }
  | {
      kind: "row";
      index: number;
      startClient: number;
      startSize: number;
      size: number;
    };

type DragMode =
  | { kind: "cells" }
  | { kind: "columns"; anchor: number }
  | { kind: "rows"; anchor: number };

type OverlayRect = { top: number; left: number; width: number; height: number };

export function Grid({
  sheet,
  formulaEngine,
  selection,
  editingCell,
  copiedRange = null,
  zoomLevel = 100,
  showGridlines = true,
  showHeaders = true,
  showFormulas = false,
  freezeTopRow = false,
  freezeFirstColumn = false,
  getCellFormat,
  getCellComment = () => null,
  getCellHyperlink = () => null,
  getCellReadOnly = () => false,
  getCellValidation = () => null,
  getCellConditionalFormatRules = () => [],
  scrollRef,
  onSelectionChange,
  onStartEdit,
  onEditValueChange,
  onCommitEdit,
  onCancelEdit,
  onPasteText,
  onKeyCommand,
  onAutoFill,
  onAutoFillDoubleClick,
  onCellContextMenu,
  onAutoFilterColumn,
  onClearAutoFilterColumn,
  onSortAutoFilterColumn,
  onColumnResize,
  onRowResize,
  onColumnAutoFit,
  onRowAutoFit,
  onRegisterScrollApi
}: GridProps) {
  const [dragMode, setDragMode] = useState<DragMode | null>(null);
  const [autoFillDrag, setAutoFillDrag] = useState<{ source: CellRange; target: CellRange } | null>(null);
  const [resizeDraft, setResizeDraft] = useState<ResizeDraft | null>(null);
  const [viewport, setViewport] = useState({ scrollTop: 0, height: DEFAULT_VIEWPORT_HEIGHT });
  const scrollFrameRef = useRef<number | null>(null);
  const dragModeRef = useRef<DragMode | null>(null);
  const autoFillDragRef = useRef<{ source: CellRange; target: CellRange } | null>(null);
  const pointerClientRef = useRef<{ x: number; y: number } | null>(null);
  const lastDragTargetRef = useRef<{ row: number; column: number } | null>(null);
  const normalizedSelection = useMemo(() => normalizeRange(selection), [selection]);
  // Everything derived from the sheet alone is memoized on the sheet snapshot's
  // identity: scroll/selection renders must stay O(visible rows), not O(rowCount).
  const columns = useMemo(
    () =>
      Array.from({ length: sheet.columnCount }, (_, column) => column).filter(
        (column) => !(sheet.hiddenColumns ?? {})[String(column)]
      ),
    [sheet]
  );
  const columnWidths = useMemo(
    () => Array.from({ length: sheet.columnCount }, (_, column) => columnWidth(sheet, column, resizeDraft)),
    [sheet, resizeDraft]
  );
  // Layout x-offsets of the visible columns, in unzoomed content pixels past the
  // row header. Drives pointer->cell math and the selection overlays.
  const columnLayout = useMemo(() => {
    let cursor = 0;
    return columns.map((column) => {
      const width = columnWidths[column];
      const entry = { column, left: cursor, width };
      cursor += width;
      return entry;
    });
  }, [columns, columnWidths]);
  const filteredRows = useMemo(
    () =>
      getVisibleRows(sheet.rowCount, sheet.filters ?? [], (row, column) =>
        formulaEngine.getDisplayValue(sheet.id, formatCellAddress({ row, column }))
      ).filter((row) => !(sheet.hiddenRows ?? {})[String(row)]),
    [formulaEngine, sheet]
  );
  const rowMeasurements = useMemo(() => measureRows(sheet, filteredRows, resizeDraft), [filteredRows, resizeDraft, sheet]);
  const firstVisibleIndex = Math.max(0, findFirstVisibleRow(rowMeasurements, viewport.scrollTop) - ROW_OVERSCAN);
  const lastVisibleIndex = Math.min(
    rowMeasurements.length,
    findLastVisibleRow(rowMeasurements, viewport.scrollTop + viewport.height) + ROW_OVERSCAN + 1
  );
  const visibleRowMeasurements = rowMeasurements.slice(firstVisibleIndex, Math.max(firstVisibleIndex + 1, lastVisibleIndex));
  const frozenTopRow =
    freezeTopRow &&
    filteredRows.includes(0) &&
    firstVisibleIndex > 0 &&
    !visibleRowMeasurements.some((measurement) => measurement.row === 0);
  const rows = frozenTopRow
    ? [{ row: 0, height: rowHeight(sheet, 0, resizeDraft) }, ...visibleRowMeasurements]
    : visibleRowMeasurements;
  const topSpacerHeight = Math.max(
    0,
    (rowMeasurements[firstVisibleIndex]?.start ?? 0) - (frozenTopRow ? rowHeight(sheet, 0, resizeDraft) : 0)
  );
  const bottomSpacerHeight = Math.max(0, (rowMeasurements.at(-1)?.end ?? 0) - (visibleRowMeasurements.at(-1)?.end ?? 0));
  const isWholeSheetSelected =
    normalizedSelection.start.row === 0 &&
    normalizedSelection.start.column === 0 &&
    normalizedSelection.end.row === sheet.rowCount - 1 &&
    normalizedSelection.end.column === sheet.columnCount - 1;
  const zoomStyle = {
    "--sheet-zoom": String(zoomLevel / 100),
    "--row-header-width": showHeaders ? `${ROW_HEADER_WIDTH}px` : "0px",
    "--column-header-height": showHeaders ? `${COLUMN_HEADER_HEIGHT}px` : "0px"
  } as CSSProperties;
  const gridClassName = ["grid-scroll", showGridlines ? "" : "grid-scroll--no-gridlines"].filter(Boolean).join(" ");
  const gridColumnCount = columns.length + (showHeaders ? 1 : 0);
  const conditionalRuleValuesCache = useMemo(() => new Map<string, string[]>(), [formulaEngine, sheet]);

  function getConditionalRuleValues(rule: ConditionalFormatRule): readonly string[] {
    const cachedValues = conditionalRuleValuesCache.get(rule.id);
    if (cachedValues) {
      return cachedValues;
    }

    const values = getRangeAddresses(rule.range).map((address) => formulaEngine.getDisplayValue(sheet.id, address));
    conditionalRuleValuesCache.set(rule.id, values);
    return values;
  }

  // Rows are memoized components; every callback handed to them must be
  // referentially stable or the memo never hits. The ref always holds the latest
  // implementations/state, and the stable wrappers read through it.
  const latestRef = useRef({
    sheet,
    selectionStart: selection.start,
    normalizedSelection,
    columnLayout,
    rowMeasurements,
    showHeaders,
    zoomLevel,
    freezeTopRow,
    freezeFirstColumn,
    onSelectionChange,
    onStartEdit,
    onEditValueChange,
    onCommitEdit,
    onCancelEdit,
    onAutoFill,
    onCellContextMenu,
    onAutoFilterColumn,
    onClearAutoFilterColumn,
    onSortAutoFilterColumn,
    onRowAutoFit,
    getCellFormat,
    getCellComment,
    getCellHyperlink,
    getCellReadOnly,
    getCellValidation,
    getCellConditionalFormatRules,
    getConditionalRuleValues
  });
  latestRef.current = {
    sheet,
    selectionStart: selection.start,
    normalizedSelection,
    columnLayout,
    rowMeasurements,
    showHeaders,
    zoomLevel,
    freezeTopRow,
    freezeFirstColumn,
    onSelectionChange,
    onStartEdit,
    onEditValueChange,
    onCommitEdit,
    onCancelEdit,
    onAutoFill,
    onCellContextMenu,
    onAutoFilterColumn,
    onClearAutoFilterColumn,
    onSortAutoFilterColumn,
    onRowAutoFit,
    getCellFormat,
    getCellComment,
    getCellHyperlink,
    getCellReadOnly,
    getCellValidation,
    getCellConditionalFormatRules,
    getConditionalRuleValues
  };

  const stableRowCallbacks = useMemo(
    () => ({
      onSelectionChange: (range: CellRange) => latestRef.current.onSelectionChange(range),
      onStartEdit: (address: string) => latestRef.current.onStartEdit(address),
      onEditValueChange: (value: string) => latestRef.current.onEditValueChange(value),
      onCommitEdit: (address: string, value: string, move?: CommitEditMove) =>
        latestRef.current.onCommitEdit(address, value, move),
      onCancelEdit: () => latestRef.current.onCancelEdit(),
      onCellContextMenu: (event: { address: string; row: number; column: number; x: number; y: number }) =>
        latestRef.current.onCellContextMenu?.(event),
      onAutoFilterColumn: (column: number, values: string[]) => latestRef.current.onAutoFilterColumn?.(column, values),
      onClearAutoFilterColumn: (column: number) => latestRef.current.onClearAutoFilterColumn?.(column),
      onSortAutoFilterColumn: (column: number, direction: "asc" | "desc") =>
        latestRef.current.onSortAutoFilterColumn?.(column, direction),
      onRowAutoFit: (row: number) => latestRef.current.onRowAutoFit?.(row),
      getCellFormat: (address: string) => latestRef.current.getCellFormat(address),
      getCellComment: (address: string) => latestRef.current.getCellComment(address),
      getCellHyperlink: (address: string) => latestRef.current.getCellHyperlink(address),
      getCellReadOnly: (address: string) => latestRef.current.getCellReadOnly(address),
      getCellValidation: (address: string) => latestRef.current.getCellValidation(address),
      getCellConditionalFormatRules: (address: string) => latestRef.current.getCellConditionalFormatRules(address),
      getConditionalRuleValues: (rule: ConditionalFormatRule) => latestRef.current.getConditionalRuleValues(rule)
    }),
    []
  );

  const setFillDrag = useCallback((next: { source: CellRange; target: CellRange } | null) => {
    autoFillDragRef.current = next;
    setAutoFillDrag(next);
  }, []);

  const beginDrag = useCallback((mode: DragMode) => {
    dragModeRef.current = mode;
    setDragMode(mode);
  }, []);

  // Applies the current drag (selection extend, header extend, or fill preview)
  // for the cell under the pointer. Shared by cell mouseenter events and the
  // window-level pointer tracking, so both agree on the target.
  const applyDragTarget = useCallback((coord: { row: number; column: number }) => {
    const current = latestRef.current;
    const fill = autoFillDragRef.current;
    lastDragTargetRef.current = coord;
    if (fill) {
      const target = createAutoFillTarget(fill.source, coord.row, coord.column);
      if (!rangesEqual(target, fill.target)) {
        setFillDrag({ source: fill.source, target });
      }
      return;
    }
    const mode = dragModeRef.current;
    if (!mode) {
      return;
    }
    if (mode.kind === "cells") {
      current.onSelectionChange({ start: current.selectionStart, end: { row: coord.row, column: coord.column } });
      return;
    }
    if (mode.kind === "columns") {
      current.onSelectionChange({
        start: { row: 0, column: mode.anchor },
        end: { row: current.sheet.rowCount - 1, column: coord.column }
      });
      return;
    }
    current.onSelectionChange({
      start: { row: mode.anchor, column: 0 },
      end: { row: coord.row, column: current.sheet.columnCount - 1 }
    });
  }, [setFillDrag]);

  // Maps a client-space point to the sheet cell underneath it, clamped to the
  // sheet bounds, accounting for scroll position, sticky headers, and zoom.
  const cellAtClientPoint = useCallback((clientX: number, clientY: number): { row: number; column: number } | null => {
    const element = scrollRef?.current;
    if (!element) {
      return null;
    }
    const current = latestRef.current;
    const { columnLayout: layout, rowMeasurements: measurements } = current;
    if (layout.length === 0 || measurements.length === 0) {
      return null;
    }
    const rect = element.getBoundingClientRect();
    const zoom = (current.zoomLevel || 100) / 100;
    const headerWidth = current.showHeaders ? ROW_HEADER_WIDTH : 0;
    const headerHeight = current.showHeaders ? COLUMN_HEADER_HEIGHT : 0;
    const contentX = (clientX - rect.left + element.scrollLeft) / zoom - headerWidth;
    const contentY = (clientY - rect.top + element.scrollTop) / zoom - headerHeight;
    // Frozen panes render sticky over scrolled content: a pointer inside the
    // frozen band must hit the frozen row/column, not the row underneath it.
    const viewportY = (clientY - rect.top) / zoom - headerHeight;
    const viewportX = (clientX - rect.left) / zoom - headerWidth;
    const frozenRowHit =
      current.freezeTopRow &&
      element.scrollTop > 0 &&
      measurements[0].row === 0 &&
      viewportY >= 0 &&
      viewportY < measurements[0].end - measurements[0].start;
    const frozenColumnHit =
      current.freezeFirstColumn &&
      element.scrollLeft > 0 &&
      layout[0].column === 0 &&
      viewportX >= 0 &&
      viewportX < layout[0].width;

    let low = 0;
    let high = layout.length - 1;
    let column = layout[high].column;
    if (contentX < layout[0].left + layout[0].width) {
      column = layout[0].column;
    } else if (contentX < layout[high].left) {
      while (low <= high) {
        const mid = (low + high) >> 1;
        const entry = layout[mid];
        if (contentX < entry.left) {
          high = mid - 1;
        } else if (contentX >= entry.left + entry.width) {
          low = mid + 1;
        } else {
          column = entry.column;
          break;
        }
      }
    }

    let rowLow = 0;
    let rowHigh = measurements.length - 1;
    let row = measurements[rowHigh].row;
    if (contentY < measurements[0].end) {
      row = measurements[0].row;
    } else if (contentY < measurements[rowHigh].start) {
      while (rowLow <= rowHigh) {
        const mid = (rowLow + rowHigh) >> 1;
        const measurement = measurements[mid];
        if (contentY < measurement.start) {
          rowHigh = mid - 1;
        } else if (contentY >= measurement.end) {
          rowLow = mid + 1;
        } else {
          row = measurement.row;
          break;
        }
      }
    }

    return { row: frozenRowHit ? 0 : row, column: frozenColumnHit ? 0 : column };
  }, [scrollRef]);

  const updateDragTargetFromPointer = useCallback(() => {
    const pointer = pointerClientRef.current;
    if (!pointer) {
      return;
    }
    const raw = cellAtClientPoint(pointer.x, pointer.y);
    if (!raw) {
      return;
    }
    // Header drags only consume one axis; collapse the other so crossing cells
    // perpendicular to the drag doesn't fire redundant selection updates.
    const mode = dragModeRef.current;
    const coord =
      mode?.kind === "columns"
        ? { row: 0, column: raw.column }
        : mode?.kind === "rows"
        ? { row: raw.row, column: 0 }
        : raw;
    const last = lastDragTargetRef.current;
    if (last && last.row === coord.row && last.column === coord.column) {
      return;
    }
    applyDragTarget(coord);
  }, [applyDragTarget, cellAtClientPoint]);

  const finalizeDrag = useCallback(() => {
    const fill = autoFillDragRef.current;
    if (fill) {
      autoFillDragRef.current = null;
      setAutoFillDrag(null);
      if (!rangesEqual(fill.source, fill.target)) {
        latestRef.current.onAutoFill?.(fill.source, fill.target);
      }
    }
    if (dragModeRef.current) {
      dragModeRef.current = null;
      setDragMode(null);
    }
    lastDragTargetRef.current = null;
    pointerClientRef.current = null;
  }, []);

  const handleCellMouseDown = useCallback((row: number, column: number, shiftKey: boolean) => {
    const current = latestRef.current;
    const anchor = shiftKey ? current.selectionStart : { row, column };
    beginDrag({ kind: "cells" });
    current.onSelectionChange({ start: anchor, end: { row, column } });
  }, [beginDrag]);

  const handleCellClick = useCallback((row: number, column: number, shiftKey: boolean) => {
    const current = latestRef.current;
    const anchor = shiftKey ? current.selectionStart : { row, column };
    current.onSelectionChange({ start: anchor, end: { row, column } });
  }, []);

  const handleCellPointerEnter = useCallback((row: number, column: number) => {
    if (autoFillDragRef.current || dragModeRef.current?.kind === "cells") {
      applyDragTarget({ row, column });
    }
  }, [applyDragTarget]);

  const handleStartAutoFill = useCallback(() => {
    dragModeRef.current = null;
    setDragMode(null);
    const source = latestRef.current.normalizedSelection;
    setFillDrag({ source, target: source });
  }, [setFillDrag]);

  const handleColumnHeaderMouseDown = useCallback((column: number, shiftKey: boolean) => {
    const current = latestRef.current;
    const anchor = shiftKey ? current.selectionStart.column : column;
    beginDrag({ kind: "columns", anchor });
    current.onSelectionChange({
      start: { row: 0, column: anchor },
      end: { row: current.sheet.rowCount - 1, column }
    });
  }, [beginDrag]);

  const handleColumnHeaderEnter = useCallback((column: number) => {
    const mode = dragModeRef.current;
    if (mode?.kind === "columns") {
      applyDragTarget({ row: 0, column });
    }
  }, [applyDragTarget]);

  const handleSelectColumn = useCallback((column: number, shiftKey: boolean) => {
    const current = latestRef.current;
    const anchor = shiftKey ? current.selectionStart.column : column;
    current.onSelectionChange({
      start: { row: 0, column: anchor },
      end: { row: current.sheet.rowCount - 1, column }
    });
  }, []);

  const handleRowHeaderMouseDown = useCallback((row: number, shiftKey: boolean) => {
    const current = latestRef.current;
    const anchor = shiftKey ? current.selectionStart.row : row;
    beginDrag({ kind: "rows", anchor });
    current.onSelectionChange({
      start: { row: anchor, column: 0 },
      end: { row, column: current.sheet.columnCount - 1 }
    });
  }, [beginDrag]);

  const handleRowHeaderEnter = useCallback((row: number) => {
    const mode = dragModeRef.current;
    if (mode?.kind === "rows") {
      applyDragTarget({ row, column: 0 });
    }
  }, [applyDragTarget]);

  const handleSelectRow = useCallback((selectedRow: number, shiftKey = false) => {
    const current = latestRef.current;
    const anchor = shiftKey ? current.selectionStart.row : selectedRow;
    current.onSelectionChange({
      start: { row: anchor, column: 0 },
      end: { row: selectedRow, column: current.sheet.columnCount - 1 }
    });
  }, []);

  const handleStartRowResize = useCallback((row: number, event: React.MouseEvent<HTMLButtonElement>) => {
    const currentSheet = latestRef.current.sheet;
    setResizeDraft({
      kind: "row",
      index: row,
      startClient: event.clientY,
      startSize: rowHeight(currentSheet, row, null),
      size: rowHeight(currentSheet, row, null)
    });
  }, []);

  const editingRow = editingCell ? addressToCoord(editingCell.address).row : -1;

  useEffect(() => {
    setViewport({ scrollTop: 0, height: scrollRef?.current?.clientHeight || DEFAULT_VIEWPORT_HEIGHT });
    if (scrollRef?.current) {
      scrollRef.current.scrollTop = 0;
      scrollRef.current.scrollLeft = 0;
    }
  }, [scrollRef, sheet.id]);

  useEffect(() => {
    const element = scrollRef?.current;
    if (!element || typeof ResizeObserver === "undefined") {
      return undefined;
    }
    const observer = new ResizeObserver(() => {
      const height = element.clientHeight || DEFAULT_VIEWPORT_HEIGHT;
      setViewport((current) => (current.height === height ? current : { ...current, height }));
    });
    observer.observe(element);
    return () => observer.disconnect();
  }, [scrollRef]);

  useEffect(
    () => () => {
      if (scrollFrameRef.current !== null) {
        cancelAnimationFrame(scrollFrameRef.current);
      }
    },
    []
  );

  useEffect(() => {
    if (!resizeDraft) {
      return undefined;
    }
    const activeResize = resizeDraft;

    function handleMouseMove(event: MouseEvent) {
      setResizeDraft((current) => {
        if (!current) {
          return current;
        }
        const delta = current.kind === "column" ? event.clientX - current.startClient : event.clientY - current.startClient;
        const size = current.kind === "column" ? clampColumnWidth(current.startSize + delta) : clampRowHeight(current.startSize + delta);
        return { ...current, size };
      });
    }

    function handleMouseUp() {
      if (activeResize.kind === "column") {
        onColumnResize?.(activeResize.index, activeResize.size);
      } else {
        onRowResize?.(activeResize.index, activeResize.size);
      }
      setResizeDraft(null);
    }

    window.addEventListener("mousemove", handleMouseMove);
    window.addEventListener("mouseup", handleMouseUp);
    return () => {
      window.removeEventListener("mousemove", handleMouseMove);
      window.removeEventListener("mouseup", handleMouseUp);
    };
  }, [onColumnResize, onRowResize, resizeDraft]);

  // Window-level drag tracking: keeps a selection or fill drag alive when the
  // pointer leaves the grid (Excel keeps dragging), and auto-scrolls the
  // viewport toward the pointer at the edges.
  const isPointerDragActive = Boolean(dragMode) || Boolean(autoFillDrag);
  useEffect(() => {
    if (!isPointerDragActive) {
      return undefined;
    }

    function handleWindowMouseMove(event: MouseEvent) {
      pointerClientRef.current = { x: event.clientX, y: event.clientY };
      updateDragTargetFromPointer();
    }

    function handleWindowMouseUp() {
      finalizeDrag();
    }

    let rafId: number | null = null;
    const autoScrollLoop = () => {
      const element = scrollRef?.current;
      const pointer = pointerClientRef.current;
      if (element && pointer) {
        const rect = element.getBoundingClientRect();
        const zoom = (latestRef.current.zoomLevel || 100) / 100;
        const innerLeft = rect.left + (latestRef.current.showHeaders ? ROW_HEADER_WIDTH * zoom : 0);
        const innerTop = rect.top + (latestRef.current.showHeaders ? COLUMN_HEADER_HEIGHT * zoom : 0);
        // Header drags keep the pointer inside a sticky header band the whole
        // time; only scroll along the axis that drag actually selects.
        const dragKind = dragModeRef.current?.kind;
        let deltaX = 0;
        let deltaY = 0;
        if (dragKind !== "columns") {
          if (pointer.y > rect.bottom) {
            deltaY = Math.min(AUTO_SCROLL_MAX_STEP, (pointer.y - rect.bottom) * 0.35 + 2);
          } else if (pointer.y < innerTop) {
            deltaY = -Math.min(AUTO_SCROLL_MAX_STEP, (innerTop - pointer.y) * 0.35 + 2);
          }
        }
        if (dragKind !== "rows") {
          if (pointer.x > rect.right) {
            deltaX = Math.min(AUTO_SCROLL_MAX_STEP, (pointer.x - rect.right) * 0.35 + 2);
          } else if (pointer.x < innerLeft) {
            deltaX = -Math.min(AUTO_SCROLL_MAX_STEP, (innerLeft - pointer.x) * 0.35 + 2);
          }
        }
        if (deltaX !== 0 || deltaY !== 0) {
          element.scrollLeft += deltaX;
          element.scrollTop += deltaY;
          updateDragTargetFromPointer();
        }
      }
      rafId = requestAnimationFrame(autoScrollLoop);
    };
    rafId = requestAnimationFrame(autoScrollLoop);

    window.addEventListener("mousemove", handleWindowMouseMove);
    window.addEventListener("mouseup", handleWindowMouseUp);
    return () => {
      window.removeEventListener("mousemove", handleWindowMouseMove);
      window.removeEventListener("mouseup", handleWindowMouseUp);
      if (rafId !== null) {
        cancelAnimationFrame(rafId);
      }
    };
  }, [finalizeDrag, isPointerDragActive, scrollRef, updateDragTargetFromPointer]);

  const ensureCellVisible = useCallback((row: number, column: number) => {
    const element = scrollRef?.current;
    if (!element) {
      return;
    }
    const current = latestRef.current;
    const zoom = (current.zoomLevel || 100) / 100;
    const headerWidth = current.showHeaders ? ROW_HEADER_WIDTH : 0;
    const headerHeight = current.showHeaders ? COLUMN_HEADER_HEIGHT : 0;
    const rowRect = rowRangeRect(current.rowMeasurements, row, row);
    const columnRect = columnRangeRect(current.columnLayout, column, column);
    if (rowRect) {
      const top = (headerHeight + rowRect.top) * zoom;
      const bottom = (headerHeight + rowRect.top + rowRect.height) * zoom;
      const viewTop = element.scrollTop + headerHeight * zoom;
      const viewBottom = element.scrollTop + element.clientHeight;
      if (top < viewTop) {
        element.scrollTop = Math.max(0, top - headerHeight * zoom);
      } else if (bottom > viewBottom) {
        element.scrollTop = bottom - element.clientHeight;
      }
    }
    if (columnRect) {
      const left = (headerWidth + columnRect.left) * zoom;
      const right = (headerWidth + columnRect.left + columnRect.width) * zoom;
      const viewLeft = element.scrollLeft + headerWidth * zoom;
      const viewRight = element.scrollLeft + element.clientWidth;
      if (left < viewLeft) {
        element.scrollLeft = Math.max(0, left - headerWidth * zoom);
      } else if (right > viewRight) {
        element.scrollLeft = right - element.clientWidth;
      }
    }
  }, [scrollRef]);

  useEffect(() => {
    onRegisterScrollApi?.({ ensureCellVisible });
  }, [ensureCellVisible, onRegisterScrollApi]);

  const headerWidth = showHeaders ? ROW_HEADER_WIDTH : 0;
  const headerHeight = showHeaders ? COLUMN_HEADER_HEIGHT : 0;
  // The outline covers any merge the selection touches, like Excel's border.
  const selectionRect = useMemo(
    () =>
      rangeOverlayRect(
        expandRangeToMerges(normalizedSelection, sheet.merges),
        rowMeasurements,
        columnLayout,
        headerWidth,
        headerHeight
      ),
    [columnLayout, headerHeight, headerWidth, normalizedSelection, rowMeasurements, sheet.merges]
  );
  const fillPreviewRange = autoFillDrag ? normalizeRange(autoFillDrag.target) : null;
  const fillPreviewRect = useMemo(
    () =>
      fillPreviewRange
        ? rangeOverlayRect(fillPreviewRange, rowMeasurements, columnLayout, headerWidth, headerHeight)
        : null,
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [
      columnLayout,
      headerHeight,
      headerWidth,
      rowMeasurements,
      fillPreviewRange?.start.row,
      fillPreviewRange?.start.column,
      fillPreviewRange?.end.row,
      fillPreviewRange?.end.column
    ]
  );
  const copiedRect = useMemo(() => {
    if (!copiedRange) {
      return null;
    }
    return rangeOverlayRect(normalizeRange(copiedRange), rowMeasurements, columnLayout, headerWidth, headerHeight);
  }, [columnLayout, copiedRange, headerHeight, headerWidth, rowMeasurements]);
  const showFillHandle = !editingCell && !autoFillDrag && Boolean(selectionRect) && Boolean(onAutoFill);

  return (
    <div
      ref={scrollRef}
      className={gridClassName}
      role="grid"
      aria-label="Spreadsheet grid"
      data-zoom-level={zoomLevel}
      data-gridlines={showGridlines ? "visible" : "hidden"}
      data-headers={showHeaders ? "visible" : "hidden"}
      style={zoomStyle}
      aria-rowcount={sheet.rowCount}
      aria-colcount={sheet.columnCount}
      tabIndex={0}
      onKeyDown={onKeyCommand}
      onScroll={(event) => {
        // Leading + trailing throttle: respond to the first scroll event of a frame
        // immediately, coalesce the rest into one trailing update. A fast fling
        // renders once per frame instead of once per scroll event.
        const element = event.currentTarget;
        const applyViewport = () => {
          const nextViewport = {
            scrollTop: element.scrollTop,
            height: element.clientHeight || DEFAULT_VIEWPORT_HEIGHT
          };
          setViewport((current) =>
            current.scrollTop === nextViewport.scrollTop && current.height === nextViewport.height ? current : nextViewport
          );
        };
        if (scrollFrameRef.current !== null) {
          return;
        }
        applyViewport();
        scrollFrameRef.current = requestAnimationFrame(() => {
          scrollFrameRef.current = null;
          applyViewport();
        });
      }}
      onPaste={(event) => {
        event.preventDefault();
        onPasteText(event.clipboardData.getData("text/plain") || event.clipboardData.getData("Text"));
      }}
      onMouseUp={finalizeDrag}
    >
      <div
        className="spreadsheet-grid"
        style={{
          gridTemplateColumns: [showHeaders ? `${ROW_HEADER_WIDTH}px` : "", ...columns.map((column) => `${columnWidths[column]}px`)]
            .filter(Boolean)
            .join(" ")
        }}
      >
        {showHeaders ? (
          <button
            type="button"
            className="corner-cell"
            aria-label="Select sheet"
            aria-pressed={isWholeSheetSelected}
            onClick={() =>
              onSelectionChange({
                start: { row: 0, column: 0 },
                end: { row: sheet.rowCount - 1, column: sheet.columnCount - 1 }
              })
            }
          />
        ) : null}
        {showHeaders
          ? columns.map((column) => {
          const columnName = columnIndexToName(column);
          const isColumnSelected =
            normalizedSelection.start.row === 0 &&
            normalizedSelection.end.row === sheet.rowCount - 1 &&
            column >= normalizedSelection.start.column &&
            column <= normalizedSelection.end.column;
          const isColumnHit =
            !isColumnSelected &&
            column >= normalizedSelection.start.column &&
            column <= normalizedSelection.end.column;

          return (
          <div
            key={column}
            className={[
              "column-header",
              isColumnSelected ? "selected-header" : "",
              isColumnHit ? "column-header--hit" : ""
            ]
              .filter(Boolean)
              .join(" ")}
            role="columnheader"
            aria-label={`Column ${columnName}`}
            aria-selected={isColumnSelected}
            tabIndex={0}
            style={{ width: columnWidths[column] }}
            onMouseDown={(event) => {
              if (event.button !== 0) {
                return;
              }
              event.preventDefault();
              handleColumnHeaderMouseDown(column, event.shiftKey);
            }}
            onMouseEnter={() => handleColumnHeaderEnter(column)}
            onClick={(event) => handleSelectColumn(column, event.shiftKey)}
            onKeyDown={(event) => {
              if (event.key === "Enter" || event.key === " ") {
                event.preventDefault();
                handleSelectColumn(column, event.shiftKey);
              }
            }}
          >
            {columnName}
            <button
              type="button"
              className="column-resize-handle"
              aria-label={`Resize column ${columnName}`}
              onClick={(event) => {
                event.preventDefault();
                event.stopPropagation();
              }}
              onDoubleClick={(event) => {
                event.preventDefault();
                event.stopPropagation();
                onColumnAutoFit?.(column);
              }}
              onMouseDown={(event) => {
                event.preventDefault();
                event.stopPropagation();
                setResizeDraft({
                  kind: "column",
                  index: column,
                  startClient: event.clientX,
                  startSize: columnWidth(sheet, column, null),
                  size: columnWidth(sheet, column, null)
                });
              }}
            />
          </div>
          );
        })
          : null}
        {topSpacerHeight > 0 ? (
          <div
            className="grid-row-spacer"
            style={{ gridColumn: `1 / span ${gridColumnCount}`, height: topSpacerHeight }}
          />
        ) : null}
        {rows.map(({ row, height }) => (
          <MemoRowFragment
            key={row}
            row={row}
            rowHeight={height}
            columns={columns}
            columnWidths={columnWidths}
            sheet={sheet}
            formulaEngine={formulaEngine}
            selection={normalizedSelection}
            activeRow={selection.start.row}
            activeColumn={selection.start.column}
            editingCell={editingRow === row ? editingCell : null}
            showHeaders={showHeaders}
            showFormulas={showFormulas}
            freezeTopRow={freezeTopRow}
            freezeFirstColumn={freezeFirstColumn}
            getCellFormat={stableRowCallbacks.getCellFormat}
            getCellComment={stableRowCallbacks.getCellComment}
            getCellHyperlink={stableRowCallbacks.getCellHyperlink}
            getCellReadOnly={stableRowCallbacks.getCellReadOnly}
            getCellValidation={stableRowCallbacks.getCellValidation}
            getCellConditionalFormatRules={stableRowCallbacks.getCellConditionalFormatRules}
            getConditionalRuleValues={stableRowCallbacks.getConditionalRuleValues}
            onSelectionChange={stableRowCallbacks.onSelectionChange}
            onCellContextMenu={stableRowCallbacks.onCellContextMenu}
            onAutoFilterColumn={stableRowCallbacks.onAutoFilterColumn}
            onClearAutoFilterColumn={stableRowCallbacks.onClearAutoFilterColumn}
            onSortAutoFilterColumn={stableRowCallbacks.onSortAutoFilterColumn}
            onCellMouseDown={handleCellMouseDown}
            onCellClick={handleCellClick}
            onCellPointerEnter={handleCellPointerEnter}
            onRowHeaderMouseDown={handleRowHeaderMouseDown}
            onRowHeaderEnter={handleRowHeaderEnter}
            onStartEdit={stableRowCallbacks.onStartEdit}
            onEditValueChange={stableRowCallbacks.onEditValueChange}
            onCommitEdit={stableRowCallbacks.onCommitEdit}
            onCancelEdit={stableRowCallbacks.onCancelEdit}
            onSelectRow={handleSelectRow}
            onStartRowResize={handleStartRowResize}
            onRowAutoFit={stableRowCallbacks.onRowAutoFit}
          />
        ))}
        {bottomSpacerHeight > 0 ? (
          <div
            className="grid-row-spacer"
            style={{ gridColumn: `1 / span ${gridColumnCount}`, height: bottomSpacerHeight }}
          />
        ) : null}
        {selectionRect && !editingCell ? (
          <div
            className="selection-outline"
            aria-hidden="true"
            style={{
              top: selectionRect.top,
              left: selectionRect.left,
              width: selectionRect.width,
              height: selectionRect.height
            }}
          />
        ) : null}
        {copiedRect ? (
          <div
            className="copy-marquee"
            aria-hidden="true"
            style={{ top: copiedRect.top, left: copiedRect.left, width: copiedRect.width, height: copiedRect.height }}
          />
        ) : null}
        {fillPreviewRect ? (
          <div
            className="fill-preview-outline"
            aria-hidden="true"
            style={{
              top: fillPreviewRect.top,
              left: fillPreviewRect.left,
              width: fillPreviewRect.width,
              height: fillPreviewRect.height
            }}
          />
        ) : null}
        {showFillHandle && selectionRect ? (
          <button
            type="button"
            className="auto-fill-handle"
            aria-label="AutoFill selection"
            title="AutoFill selection"
            style={{
              top: selectionRect.top + selectionRect.height,
              left: selectionRect.left + selectionRect.width
            }}
            onMouseDown={(event) => {
              if (event.button !== 0) {
                return;
              }
              event.preventDefault();
              event.stopPropagation();
              handleStartAutoFill();
            }}
            onClick={(event) => {
              event.preventDefault();
              event.stopPropagation();
            }}
            onDoubleClick={(event) => {
              event.preventDefault();
              event.stopPropagation();
              onAutoFillDoubleClick?.();
            }}
          />
        ) : null}
      </div>
    </div>
  );
}

// Rows re-render only when their own props change: scrolling mounts new rows,
// typing re-renders just the editing row, and an edit commit (new sheet
// snapshot) refreshes the visible window.
const MemoRowFragment = memo(RowFragment);

function RowFragment({
  row,
  rowHeight,
  columns,
  columnWidths,
  sheet,
  formulaEngine,
  selection,
  activeRow,
  activeColumn,
  editingCell,
  showHeaders,
  showFormulas,
  freezeTopRow,
  freezeFirstColumn,
  getCellFormat,
  getCellComment,
  getCellHyperlink,
  getCellReadOnly,
  getCellValidation,
  getCellConditionalFormatRules,
  getConditionalRuleValues,
  onSelectionChange,
  onCellContextMenu,
  onAutoFilterColumn,
  onClearAutoFilterColumn,
  onSortAutoFilterColumn,
  onCellMouseDown,
  onCellClick,
  onCellPointerEnter,
  onRowHeaderMouseDown,
  onRowHeaderEnter,
  onStartEdit,
  onEditValueChange,
  onCommitEdit,
  onCancelEdit,
  onSelectRow,
  onStartRowResize,
  onRowAutoFit
}: {
  row: number;
  rowHeight: number;
  columns: number[];
  columnWidths: number[];
  sheet: SheetModel;
  formulaEngine: FormulaEngine;
  selection: CellRange;
  activeRow: number;
  activeColumn: number;
  editingCell: { address: string; value: string } | null;
  showHeaders: boolean;
  showFormulas: boolean;
  freezeTopRow: boolean;
  freezeFirstColumn: boolean;
  getCellFormat: (address: string) => CellFormat | undefined;
  getCellComment: (address: string) => string | null | undefined;
  getCellHyperlink: (address: string) => string | null | undefined;
  getCellReadOnly: (address: string) => boolean;
  getCellValidation: (address: string) => DataValidationRule | null | undefined;
  getCellConditionalFormatRules: (address: string) => ConditionalFormatRule[];
  getConditionalRuleValues: (rule: ConditionalFormatRule) => readonly string[];
  onSelectionChange: (range: CellRange) => void;
  onCellContextMenu?: (event: { address: string; row: number; column: number; x: number; y: number }) => void;
  onAutoFilterColumn?: (column: number, values: string[]) => void;
  onClearAutoFilterColumn?: (column: number) => void;
  onSortAutoFilterColumn?: (column: number, direction: "asc" | "desc") => void;
  onCellMouseDown: (row: number, column: number, shiftKey: boolean) => void;
  onCellClick: (row: number, column: number, shiftKey: boolean) => void;
  onCellPointerEnter: (row: number, column: number) => void;
  onRowHeaderMouseDown: (row: number, shiftKey: boolean) => void;
  onRowHeaderEnter: (row: number) => void;
  onStartEdit: (address: string) => void;
  onEditValueChange: (value: string) => void;
  onCommitEdit: (address: string, value: string, move?: CommitEditMove) => void;
  onCancelEdit: () => void;
  onSelectRow: (row: number, shiftKey?: boolean) => void;
  onStartRowResize: (row: number, event: React.MouseEvent<HTMLButtonElement>) => void;
  onRowAutoFit?: (row: number) => void;
}) {
  const [activeSuggestionIndex, setActiveSuggestionIndex] = useState(0);
  const [dismissedEditingSuggestion, setDismissedEditingSuggestion] = useState<{ address: string; value: string } | null>(null);
  const [openValidationDropdownAddress, setOpenValidationDropdownAddress] = useState<string | null>(null);
  const [openAutoFilterColumn, setOpenAutoFilterColumn] = useState<number | null>(null);
  const [autoFilterDraft, setAutoFilterDraft] = useState<{ column: number; values: string[] } | null>(null);
  const editingSuggestionKey = useMemo(() => {
    if (!editingCell) {
      return "";
    }

    return getFormulaSuggestions(editingCell.value)
      .map((suggestion) => suggestion.name)
      .join("|");
  }, [editingCell?.address, editingCell?.value]);
  const isRowSelected =
    selection.start.column === 0 && selection.end.column === sheet.columnCount - 1 && row >= selection.start.row && row <= selection.end.row;
  const isRowHit = !isRowSelected && row >= selection.start.row && row <= selection.end.row;

  useEffect(() => {
    setActiveSuggestionIndex(0);
  }, [editingCell?.address, editingCell?.value, editingSuggestionKey]);

  useEffect(() => {
    setDismissedEditingSuggestion(null);
  }, [editingCell?.address]);

  useEffect(() => {
    setOpenValidationDropdownAddress(null);
    setOpenAutoFilterColumn(null);
    setAutoFilterDraft(null);
  }, [selection.start.row, selection.start.column, selection.end.row, selection.end.column]);

  return (
    <>
      {showHeaders ? (
        <div
          className={[
            "row-header",
            isRowSelected ? "selected-header" : "",
            isRowHit ? "row-header--hit" : ""
          ]
            .filter(Boolean)
            .join(" ")}
          role="rowheader"
          aria-label={`Row ${row + 1}`}
          aria-selected={isRowSelected}
          tabIndex={0}
          style={{ height: rowHeight }}
          onMouseDown={(event) => {
            if (event.button !== 0) {
              return;
            }
            event.preventDefault();
            onRowHeaderMouseDown(row, event.shiftKey);
          }}
          onMouseEnter={() => onRowHeaderEnter(row)}
          onClick={(event) => onSelectRow(row, event.shiftKey)}
          onKeyDown={(event) => {
            if (event.key === "Enter" || event.key === " ") {
              event.preventDefault();
              onSelectRow(row, event.shiftKey);
            }
          }}
        >
          {row + 1}
          <button
            type="button"
            className="row-resize-handle"
            aria-label={`Resize row ${row + 1}`}
            onClick={(event) => {
              event.preventDefault();
              event.stopPropagation();
            }}
            onDoubleClick={(event) => {
              event.preventDefault();
              event.stopPropagation();
              onRowAutoFit?.(row);
            }}
            onMouseDown={(event) => {
              event.preventDefault();
              event.stopPropagation();
              onStartRowResize(row, event);
            }}
          />
        </div>
      ) : null}
      {columns.map((column) => {
        const address = formatCellAddress({ row, column });
        const mergeInfo = getMergeInfo(sheet, row, column);
        const isMergeAnchor = mergeInfo?.role === "anchor";
        const isMergeCovered = mergeInfo?.role === "covered";
        const format = getCellFormat(address);
        const comment = getCellComment(address)?.trim() ?? "";
        const hyperlink = isMergeCovered ? "" : getCellHyperlink(address)?.trim() ?? "";
        const isReadOnly = !isMergeCovered && getCellReadOnly(address);
        const validation = getCellValidation(address);
        const rawContent = isMergeCovered ? null : formulaEngine.getRawContent(sheet.id, address);
        const rawDisplayValue = isMergeCovered ? "" : formulaEngine.getDisplayValue(sheet.id, address);
        const formulaText = typeof rawContent === "string" && rawContent.startsWith("=") ? rawContent : "";
        const conditionalRules = getCellConditionalFormatRules(address);
        const conditionalFormat = getConditionalFormatForValue(rawDisplayValue, conditionalRules, {
          getRuleValues: getConditionalRuleValues
        });
        const conditionalDataBar = getConditionalDataBarForValue(rawDisplayValue, conditionalRules, {
          getRuleValues: getConditionalRuleValues
        });
        const mergedFormat = conditionalFormat ? { ...(format ?? {}), ...conditionalFormat } : format;
        const displayValue = formatDisplayValue(rawDisplayValue, mergedFormat);
        const visibleValue = showFormulas && formulaText ? formulaText : hyperlink && displayValue === "" ? hyperlink : displayValue;
        const isSelected = isInSelection(row, column, selection);
        const isActiveCell = row === activeRow && column === activeColumn;
        const isEditing = editingCell?.address === address && !isMergeCovered;
        const suggestions = isEditing ? getFormulaSuggestions(editingCell.value) : [];
        const visibleSuggestions =
          dismissedEditingSuggestion?.address === address && dismissedEditingSuggestion.value === editingCell?.value ? [] : suggestions;
        const suggestionListId = `cell-editor-${address.toLowerCase()}-formula-suggestions`;
        const activeSuggestion = visibleSuggestions[wrapSuggestionIndex(activeSuggestionIndex, visibleSuggestions.length)];
        const rawValidationValue = isEditing ? editingCell.value : rawContent;
        const validationValue = rawValidationValue === null ? "" : String(rawValidationValue);
        const isInvalidValidation =
          !isMergeCovered && validationValue.trim() !== "" && !validateCellValue(validationValue, validation).valid;
        const hasValidationDropdown =
          !isMergeCovered && !isReadOnly && !isEditing && isSelected && validation?.type === "list" && validation.values.length > 0;
        const isValidationDropdownOpen = hasValidationDropdown && openValidationDropdownAddress === address;
        const autoFilterRange = sheet.autoFilterRange ? normalizeRange(sheet.autoFilterRange) : null;
        const isAutoFilterHeader =
          !isMergeCovered &&
          !isEditing &&
          Boolean(autoFilterRange) &&
          row === autoFilterRange!.start.row &&
          column >= autoFilterRange!.start.column &&
          column <= autoFilterRange!.end.column;
        const autoFilterLabel = visibleValue || columnIndexToName(column);
        const activeAutoFilter =
          autoFilterRange && isAutoFilterHeader ? findAutoFilterForColumn(sheet.filters ?? [], autoFilterRange, column) : null;
        const isAutoFilterMenuOpen = isAutoFilterHeader && openAutoFilterColumn === column;
        const autoFilterChoices =
          autoFilterRange && isAutoFilterHeader ? getAutoFilterColumnChoices(sheet, formulaEngine, autoFilterRange, column) : [];
        const selectedAutoFilterValues =
          autoFilterDraft?.column === column ? autoFilterDraft.values : activeAutoFilterValues(activeAutoFilter);
        const cellWidth =
          isMergeAnchor && mergeInfo ? sumColumnWidths(columnWidths, mergeInfo.range.start.column, mergeInfo.range.end.column) : columnWidths[column];
        const cellHeight =
          isMergeAnchor && mergeInfo ? sumRowHeights(sheet, mergeInfo.range.start.row, mergeInfo.range.end.row) : rowHeight;
        const borderStyle = getBorderStyle(mergedFormat?.borders);
        const style =
          mergedFormat?.bold ||
          mergedFormat?.italic ||
          mergedFormat?.fontFamily ||
          mergedFormat?.fontSize ||
          mergedFormat?.textColor ||
          mergedFormat?.backgroundColor ||
          mergedFormat?.horizontalAlign ||
          mergedFormat?.verticalAlign ||
          mergedFormat?.wrapText ||
          Object.keys(borderStyle).length > 0
            ? {
                fontWeight: mergedFormat?.bold ? 700 : undefined,
                fontStyle: mergedFormat?.italic ? "italic" : undefined,
                fontFamily: mergedFormat?.fontFamily,
                // fontSize is in POINTS (Excel's unit, matching xlsx round-trip);
                // CSS pt renders it at the same visual size Excel does.
                fontSize: mergedFormat?.fontSize ? `${mergedFormat.fontSize}pt` : undefined,
                color: mergedFormat?.textColor,
                backgroundColor: mergedFormat?.backgroundColor,
                textAlign: mergedFormat?.horizontalAlign,
                alignItems: verticalAlignToFlex(mergedFormat?.verticalAlign),
                whiteSpace: mergedFormat?.wrapText ? "normal" : undefined,
                ...borderStyle,
                width: cellWidth,
                minWidth: cellWidth,
                maxWidth: cellWidth,
                height: cellHeight,
                minHeight: cellHeight
              }
            : {
                width: cellWidth,
                minWidth: cellWidth,
                maxWidth: cellWidth,
                height: cellHeight,
                minHeight: cellHeight
              };

        return (
          <div
            key={address}
            role="gridcell"
            aria-label={visibleValue ? `${address} ${visibleValue}` : address}
            aria-selected={isSelected}
            aria-colspan={isMergeAnchor ? mergeInfo.columnSpan : undefined}
            aria-rowspan={isMergeAnchor ? mergeInfo.rowSpan : undefined}
            className={[
              isSelected ? "cell selected-cell" : "cell",
              isActiveCell ? "active-cell" : "",
              isEditing ? "editing-cell" : "",
              isMergeAnchor ? "merged-cell" : "",
              isMergeCovered ? "merge-covered-cell" : "",
              isInvalidValidation ? "invalid-validation-cell" : "",
              comment && !isMergeCovered ? "commented-cell" : "",
              hyperlink ? "hyperlink-cell" : "",
              isAutoFilterHeader ? "auto-filter-header-cell" : "",
              activeAutoFilter ? "filtered-header-cell" : "",
              isAutoFilterMenuOpen ? "auto-filter-menu-open" : "",
              hasValidationDropdown ? "validation-list-cell" : "",
              isValidationDropdownOpen ? "validation-dropdown-open" : "",
              mergedFormat?.wrapText ? "wrapped-cell" : "",
              isReadOnly ? "read-only-cell" : "",
              conditionalFormat ? "conditional-format-cell" : "",
              conditionalDataBar ? "conditional-data-bar-cell" : "",
              freezeTopRow && row === 0 ? "frozen-top-row-cell" : "",
              freezeFirstColumn && column === 0 ? "frozen-first-column-cell" : ""
            ]
              .filter(Boolean)
              .join(" ")}
            style={style}
            aria-readonly={isReadOnly || undefined}
            title={[isReadOnly ? "Read only" : "", comment ? `Comment: ${comment}` : "", hyperlink ? `Link: ${hyperlink}` : ""]
              .filter(Boolean)
              .join("\n") || undefined}
            onMouseDown={(event) => {
              if (event.button !== 0) {
                return;
              }
              onCellMouseDown(row, column, event.shiftKey);
            }}
            onMouseEnter={() => onCellPointerEnter(row, column)}
            onClick={(event) => onCellClick(row, column, event.shiftKey)}
            onContextMenu={(event) => {
              event.preventDefault();
              if (!isInSelection(row, column, selection)) {
                onSelectionChange({ start: { row, column }, end: { row, column } });
              }
              onCellContextMenu?.({ address, row, column, x: event.clientX, y: event.clientY });
            }}
            onDoubleClick={() => {
              if (!isMergeCovered) {
                onStartEdit(address);
              }
            }}
          >
            {conditionalDataBar ? (
              <span
                className="cell-data-bar"
                aria-hidden="true"
                style={{ width: `${conditionalDataBar.percent}%`, backgroundColor: conditionalDataBar.color }}
              />
            ) : null}
            {isEditing ? (
              <div className="cell-editor-shell" onMouseDown={(event) => event.stopPropagation()}>
                {validation?.type === "list" ? (
                  <select
                    className="cell-editor"
                    aria-label={`Cell editor ${address}`}
                    autoFocus
                    value={editingCell.value}
                    onChange={(event) => onEditValueChange(event.currentTarget.value)}
                    onBlur={() => onCommitEdit(address, editingCell.value)}
                    onKeyDown={(event) => {
                      if (event.key === "Enter") {
                        event.preventDefault();
                        onCommitEdit(address, editingCell.value, event.shiftKey ? "up" : "down");
                      }
                      if (event.key === "Escape") {
                        event.preventDefault();
                        onCancelEdit();
                      }
                    }}
                  >
                    {validation.allowBlank === false ? null : <option value="" />}
                    {validation.values.map((value) => (
                      <option key={value} value={value}>
                        {value}
                      </option>
                    ))}
                  </select>
                ) : (
                  <>
                    <input
                      className="cell-editor"
                      aria-label={`Cell editor ${address}`}
                      aria-autocomplete="list"
                      aria-controls={visibleSuggestions.length > 0 ? suggestionListId : undefined}
                      aria-expanded={visibleSuggestions.length > 0}
                      aria-activedescendant={activeSuggestion ? formulaSuggestionOptionId(suggestionListId, activeSuggestion.name) : undefined}
                      autoFocus
                      value={editingCell.value}
                      onChange={(event) => {
                        setDismissedEditingSuggestion(null);
                        onEditValueChange(event.currentTarget.value);
                      }}
                      onBlur={() => onCommitEdit(address, editingCell.value)}
                      onKeyDown={(event) => {
                        // Excel semantics: Up/Down navigate the autocomplete list, Tab accepts
                        // the highlighted suggestion (or commits and moves right when there is
                        // none), and Enter ALWAYS commits the cell, then moves down.
                        // Left/Right are left alone so the user can move the text caret.
                        if (event.key === "ArrowDown" && visibleSuggestions.length > 0) {
                          event.preventDefault();
                          setActiveSuggestionIndex((current) => wrapSuggestionIndex(current + 1, visibleSuggestions.length));
                          return;
                        }

                        if (event.key === "ArrowUp" && visibleSuggestions.length > 0) {
                          event.preventDefault();
                          setActiveSuggestionIndex((current) => wrapSuggestionIndex(current - 1, visibleSuggestions.length));
                          return;
                        }

                        if (event.key === "Tab" && visibleSuggestions.length > 0) {
                          event.preventDefault();
                          const suggestion = visibleSuggestions[wrapSuggestionIndex(activeSuggestionIndex, visibleSuggestions.length)];
                          if (suggestion) {
                            setDismissedEditingSuggestion(null);
                            onEditValueChange(insertFormulaSuggestion(editingCell.value, suggestion.name));
                          }
                          return;
                        }

                        if (event.key === "Tab") {
                          event.preventDefault();
                          onCommitEdit(address, editingCell.value, event.shiftKey ? "left" : "right");
                          return;
                        }

                        if (event.key === "Enter") {
                          event.preventDefault();
                          onCommitEdit(address, editingCell.value, event.shiftKey ? "up" : "down");
                        }
                        if (event.key === "Escape") {
                          event.preventDefault();
                          if (visibleSuggestions.length > 0) {
                            setDismissedEditingSuggestion({ address, value: editingCell.value });
                            return;
                          }
                          onCancelEdit();
                        }
                      }}
                    />
                    <div
                      className="cell-suggestion-layer"
                      onMouseDown={(event) => event.stopPropagation()}
                      onClick={(event) => event.stopPropagation()}
                    >
                      <FormulaSuggestions
                        suggestions={visibleSuggestions}
                        activeIndex={activeSuggestionIndex}
                        listId={suggestionListId}
                        onActiveIndexChange={setActiveSuggestionIndex}
                        onSelect={(name) => {
                          setDismissedEditingSuggestion(null);
                          onEditValueChange(insertFormulaSuggestion(editingCell.value, name));
                        }}
                      />
                    </div>
                  </>
                )}
              </div>
            ) : (
              <>
                {hyperlink ? (
                  <a
                    href={hyperlink}
                    target="_blank"
                    rel="noreferrer"
                    onMouseDown={(event) => event.stopPropagation()}
                    onClick={(event) => event.stopPropagation()}
                  >
                    {visibleValue}
                  </a>
                ) : (
                  <span>{visibleValue}</span>
                )}
                {isAutoFilterHeader ? (
                  <>
                    <button
                      type="button"
                      className="auto-filter-toggle"
                      aria-label={`Open AutoFilter menu for ${autoFilterLabel}`}
                      aria-haspopup="menu"
                      aria-expanded={isAutoFilterMenuOpen}
                      title={`AutoFilter menu for ${autoFilterLabel}`}
                      onMouseDown={(event) => {
                        event.preventDefault();
                        event.stopPropagation();
                      }}
                      onClick={(event) => {
                        event.stopPropagation();
                        setOpenAutoFilterColumn((current) => {
                          const nextColumn = current === column ? null : column;
                          setAutoFilterDraft(nextColumn === null ? null : { column, values: activeAutoFilterValues(activeAutoFilter) });
                          return nextColumn;
                        });
                      }}
                      onKeyDown={(event) => {
                        if (event.key === "Escape") {
                          event.preventDefault();
                          setOpenAutoFilterColumn(null);
                          setAutoFilterDraft(null);
                        }
                      }}
                    >
                      <ListFilter aria-hidden="true" />
                    </button>
                    {isAutoFilterMenuOpen ? (
                      <div
                        className="auto-filter-menu"
                        role="menu"
                        aria-label={`AutoFilter menu for ${autoFilterLabel}`}
                        onMouseDown={(event) => event.stopPropagation()}
                        onClick={(event) => event.stopPropagation()}
                        onKeyDown={(event) => {
                          if (event.key === "Escape") {
                            event.preventDefault();
                            setOpenAutoFilterColumn(null);
                            setAutoFilterDraft(null);
                          }
                        }}
                      >
                        <button
                          type="button"
                          role="menuitem"
                          onClick={() => {
                            onSortAutoFilterColumn?.(column, "asc");
                            setOpenAutoFilterColumn(null);
                            setAutoFilterDraft(null);
                          }}
                        >
                          Sort A to Z
                        </button>
                        <button
                          type="button"
                          role="menuitem"
                          onClick={() => {
                            onSortAutoFilterColumn?.(column, "desc");
                            setOpenAutoFilterColumn(null);
                            setAutoFilterDraft(null);
                          }}
                        >
                          Sort Z to A
                        </button>
                        <div className="auto-filter-menu-divider" />
                        <button
                          type="button"
                          role="menuitem"
                          disabled={!activeAutoFilter}
                          onClick={() => {
                            onClearAutoFilterColumn?.(column);
                            setOpenAutoFilterColumn(null);
                            setAutoFilterDraft(null);
                          }}
                        >
                          Clear filter from {autoFilterLabel}
                        </button>
                        <div className="auto-filter-menu-divider" />
                        <div className="auto-filter-choice-list">
                          {autoFilterChoices.map((choice) => (
                            <button
                              key={choice.key}
                              type="button"
                              role="menuitemcheckbox"
                              aria-checked={selectedAutoFilterValues.includes(choice.value)}
                              onClick={() => {
                                setAutoFilterDraft((current) => {
                                  const values = current?.column === column ? current.values : selectedAutoFilterValues;
                                  return { column, values: toggleFilterValue(values, choice.value) };
                                });
                              }}
                            >
                              {choice.label}
                            </button>
                          ))}
                        </div>
                        <button
                          type="button"
                          role="menuitem"
                          onClick={() => {
                            onAutoFilterColumn?.(column, selectedAutoFilterValues);
                            setOpenAutoFilterColumn(null);
                            setAutoFilterDraft(null);
                          }}
                        >
                          Apply selected values
                        </button>
                      </div>
                    ) : null}
                  </>
                ) : null}
                {hasValidationDropdown ? (
                  <>
                    <button
                      type="button"
                      className="validation-dropdown-toggle"
                      aria-label={`Open validation choices for ${address}`}
                      aria-haspopup="listbox"
                      aria-expanded={isValidationDropdownOpen}
                      title={`Validation choices for ${address}`}
                      onMouseDown={(event) => {
                        event.preventDefault();
                        event.stopPropagation();
                      }}
                      onClick={(event) => {
                        event.stopPropagation();
                        setOpenValidationDropdownAddress((current) => (current === address ? null : address));
                      }}
                      onKeyDown={(event) => {
                        if (event.key === "Escape") {
                          event.preventDefault();
                          setOpenValidationDropdownAddress(null);
                        }
                      }}
                    >
                      <ChevronDown aria-hidden="true" />
                    </button>
                    {isValidationDropdownOpen ? (
                      <div
                        className="validation-dropdown"
                        role="listbox"
                        aria-label={`Validation choices for ${address}`}
                        onMouseDown={(event) => event.stopPropagation()}
                        onClick={(event) => event.stopPropagation()}
                        onKeyDown={(event) => {
                          if (event.key === "Escape") {
                            event.preventDefault();
                            setOpenValidationDropdownAddress(null);
                          }
                        }}
                      >
                        {validation.values.map((value, index) => (
                          <button
                            key={`${value}-${index}`}
                            type="button"
                            role="option"
                            aria-selected={validationValue === value}
                            onClick={() => {
                              onCommitEdit(address, value);
                              setOpenValidationDropdownAddress(null);
                            }}
                          >
                            {value}
                          </button>
                        ))}
                      </div>
                    ) : null}
                  </>
                ) : null}
              </>
            )}
          </div>
        );
      })}
    </>
  );
}

function isInSelection(row: number, column: number, selection: CellRange): boolean {
  return row >= selection.start.row && row <= selection.end.row && column >= selection.start.column && column <= selection.end.column;
}

// Excel fill-handle semantics: dragging outside the source extends along the
// dominant axis (down/up/right/left); dragging back inside the source shrinks
// the range, which clears the cells left behind on release.
export function createAutoFillTarget(source: CellRange, row: number, column: number): CellRange {
  const normalized = normalizeRange(source);
  const rowOvershoot =
    row > normalized.end.row ? row - normalized.end.row : row < normalized.start.row ? row - normalized.start.row : 0;
  const columnOvershoot =
    column > normalized.end.column
      ? column - normalized.end.column
      : column < normalized.start.column
      ? column - normalized.start.column
      : 0;

  if (rowOvershoot === 0 && columnOvershoot === 0) {
    // Pointer inside the source: shrink (vertical first, matching Excel's bias).
    if (row < normalized.end.row) {
      return { start: normalized.start, end: { row, column: normalized.end.column } };
    }
    if (column < normalized.end.column) {
      return { start: normalized.start, end: { row: normalized.end.row, column } };
    }
    return normalized;
  }

  if (Math.abs(rowOvershoot) >= Math.abs(columnOvershoot)) {
    if (rowOvershoot > 0) {
      return { start: normalized.start, end: { row, column: normalized.end.column } };
    }
    return { start: { row, column: normalized.start.column }, end: normalized.end };
  }

  if (columnOvershoot > 0) {
    return { start: normalized.start, end: { row: normalized.end.row, column } };
  }
  return { start: { row: normalized.start.row, column }, end: normalized.end };
}

// Grows a range until it fully covers every merged range it intersects,
// looping because absorbing one merge can bring the range into contact with
// another (chained merges).
function expandRangeToMerges(range: CellRange, merges: SheetModel["merges"]): CellRange {
  if (!merges || merges.length === 0) {
    return range;
  }

  let current = normalizeRange(range);
  let changed = true;
  while (changed) {
    changed = false;
    for (const merge of merges) {
      const mergeRange = normalizeRange(merge.range);
      const intersects =
        mergeRange.start.row <= current.end.row &&
        mergeRange.end.row >= current.start.row &&
        mergeRange.start.column <= current.end.column &&
        mergeRange.end.column >= current.start.column;
      if (!intersects) {
        continue;
      }
      const next = {
        start: {
          row: Math.min(current.start.row, mergeRange.start.row),
          column: Math.min(current.start.column, mergeRange.start.column)
        },
        end: {
          row: Math.max(current.end.row, mergeRange.end.row),
          column: Math.max(current.end.column, mergeRange.end.column)
        }
      };
      if (
        next.start.row !== current.start.row ||
        next.start.column !== current.start.column ||
        next.end.row !== current.end.row ||
        next.end.column !== current.end.column
      ) {
        current = next;
        changed = true;
      }
    }
  }
  return current;
}

function rangeOverlayRect(
  range: CellRange,
  measurements: ReturnType<typeof measureRows>,
  columnLayout: Array<{ column: number; left: number; width: number }>,
  headerWidth: number,
  headerHeight: number
): OverlayRect | null {
  const rowRect = rowRangeRect(measurements, range.start.row, range.end.row);
  const columnRect = columnRangeRect(columnLayout, range.start.column, range.end.column);
  if (!rowRect || !columnRect) {
    return null;
  }
  return {
    top: headerHeight + rowRect.top,
    left: headerWidth + columnRect.left,
    width: columnRect.width,
    height: rowRect.height
  };
}

function rowRangeRect(
  measurements: ReturnType<typeof measureRows>,
  startRow: number,
  endRow: number
): { top: number; height: number } | null {
  if (measurements.length === 0) {
    return null;
  }
  // measurements are sorted by row; find the visible slice inside [startRow, endRow].
  let low = 0;
  let high = measurements.length - 1;
  while (low <= high) {
    const mid = (low + high) >> 1;
    if (measurements[mid].row < startRow) {
      low = mid + 1;
    } else {
      high = mid - 1;
    }
  }
  const firstIndex = low;
  low = 0;
  high = measurements.length - 1;
  while (low <= high) {
    const mid = (low + high) >> 1;
    if (measurements[mid].row <= endRow) {
      low = mid + 1;
    } else {
      high = mid - 1;
    }
  }
  const lastIndex = high;
  if (firstIndex > lastIndex) {
    return null;
  }
  return {
    top: measurements[firstIndex].start,
    height: measurements[lastIndex].end - measurements[firstIndex].start
  };
}

function columnRangeRect(
  columnLayout: Array<{ column: number; left: number; width: number }>,
  startColumn: number,
  endColumn: number
): { left: number; width: number } | null {
  if (columnLayout.length === 0) {
    return null;
  }
  let low = 0;
  let high = columnLayout.length - 1;
  while (low <= high) {
    const mid = (low + high) >> 1;
    if (columnLayout[mid].column < startColumn) {
      low = mid + 1;
    } else {
      high = mid - 1;
    }
  }
  const firstIndex = low;
  low = 0;
  high = columnLayout.length - 1;
  while (low <= high) {
    const mid = (low + high) >> 1;
    if (columnLayout[mid].column <= endColumn) {
      low = mid + 1;
    } else {
      high = mid - 1;
    }
  }
  const lastIndex = high;
  if (firstIndex > lastIndex) {
    return null;
  }
  return {
    left: columnLayout[firstIndex].left,
    width: columnLayout[lastIndex].left + columnLayout[lastIndex].width - columnLayout[firstIndex].left
  };
}

function findAutoFilterForColumn(filters: readonly SheetFilter[], range: CellRange, column: number): SheetFilter | null {
  return filters.find((filter) => filter.column === column && rangesEqual(filter.range, range)) ?? null;
}

function activeAutoFilterValues(filter: SheetFilter | null): string[] {
  if (!filter) {
    return [];
  }

  return filter.values && filter.values.length > 0 ? [...filter.values] : [filter.value];
}

function toggleFilterValue(values: readonly string[], value: string): string[] {
  return values.includes(value) ? values.filter((current) => current !== value) : [...values, value];
}

function getAutoFilterColumnChoices(
  sheet: SheetModel,
  formulaEngine: FormulaEngine,
  range: CellRange,
  column: number
): Array<{ key: string; label: string; value: string }> {
  const choices = new Map<string, { key: string; label: string; value: string }>();
  for (let row = range.start.row + 1; row <= range.end.row; row += 1) {
    const value = String(formulaEngine.getDisplayValue(sheet.id, formatCellAddress({ row, column })) ?? "");
    const key = value.trim().toLocaleLowerCase();
    if (!choices.has(key)) {
      choices.set(key, {
        key: key || "__blank__",
        label: value.trim() === "" ? "(Blanks)" : value,
        value
      });
    }
  }

  return [...choices.values()].sort((left, right) => left.label.localeCompare(right.label, undefined, { numeric: true, sensitivity: "base" }));
}

function rangesEqual(left: CellRange, right: CellRange): boolean {
  const normalizedLeft = normalizeRange(left);
  const normalizedRight = normalizeRange(right);
  return (
    normalizedLeft.start.row === normalizedRight.start.row &&
    normalizedLeft.start.column === normalizedRight.start.column &&
    normalizedLeft.end.row === normalizedRight.end.row &&
    normalizedLeft.end.column === normalizedRight.end.column
  );
}

function columnWidth(sheet: SheetModel, column: number, resizeDraft: ResizeDraft | null): number {
  if (resizeDraft?.kind === "column" && resizeDraft.index === column) {
    return resizeDraft.size;
  }
  return clampColumnWidth((sheet.columnWidths ?? {})[String(column)] ?? DEFAULT_COLUMN_WIDTH);
}

function rowHeight(sheet: SheetModel, row: number, resizeDraft: ResizeDraft | null): number {
  if (resizeDraft?.kind === "row" && resizeDraft.index === row) {
    return resizeDraft.size;
  }
  return clampRowHeight((sheet.rowHeights ?? {})[String(row)] ?? DEFAULT_ROW_HEIGHT);
}

function getMergeInfo(sheet: SheetModel, row: number, column: number) {
  const merge = (sheet.merges ?? []).find((candidate) =>
    isInSelection(row, column, normalizeRange(candidate.range))
  );
  if (!merge) {
    return null;
  }

  const range = normalizeRange(merge.range);
  const columnSpan = range.end.column - range.start.column + 1;
  const rowSpan = range.end.row - range.start.row + 1;
  const base = { range, columnSpan, rowSpan };
  return row === range.start.row && column === range.start.column
    ? { role: "anchor" as const, ...base }
    : { role: "covered" as const, ...base, anchor: { ...range.start } };
}

function getBorderStyle(borders: CellFormat["borders"]): CSSProperties {
  if (!borders) {
    return {};
  }

  return {
    borderTop: borders.top ? `${borderWidth(borders.top.style)} solid ${borders.top.color}` : undefined,
    borderRight: borders.right ? `${borderWidth(borders.right.style)} solid ${borders.right.color}` : undefined,
    borderBottom: borders.bottom ? `${borderWidth(borders.bottom.style)} solid ${borders.bottom.color}` : undefined,
    borderLeft: borders.left ? `${borderWidth(borders.left.style)} solid ${borders.left.color}` : undefined
  };
}

function borderWidth(style: NonNullable<NonNullable<CellFormat["borders"]>["top"]>["style"]): string {
  return style === "thin" ? "1px" : "1px";
}

function verticalAlignToFlex(align: CellFormat["verticalAlign"]): CSSProperties["alignItems"] | undefined {
  if (align === "top") {
    return "flex-start";
  }
  if (align === "bottom") {
    return "flex-end";
  }
  return align === "middle" ? "center" : undefined;
}

function wrapSuggestionIndex(index: number, suggestionCount: number): number {
  if (suggestionCount <= 0) {
    return 0;
  }

  return ((index % suggestionCount) + suggestionCount) % suggestionCount;
}

function sumColumnWidths(columnWidths: number[], startColumn: number, endColumn: number): number {
  let width = 0;
  for (let column = startColumn; column <= endColumn; column += 1) {
    width += columnWidths[column] ?? DEFAULT_COLUMN_WIDTH;
  }
  return width;
}

function sumRowHeights(sheet: SheetModel, startRow: number, endRow: number): number {
  let height = 0;
  for (let row = startRow; row <= endRow; row += 1) {
    height += rowHeight(sheet, row, null);
  }
  return height;
}

function measureRows(sheet: SheetModel, rows: number[], resizeDraft: ResizeDraft | null) {
  let cursor = 0;
  return rows.map((row) => {
    const height = rowHeight(sheet, row, resizeDraft);
    const measurement = {
      row,
      height,
      start: cursor,
      end: cursor + height
    };
    cursor += height;
    return measurement;
  });
}

function findFirstVisibleRow(rows: ReturnType<typeof measureRows>, scrollTop: number): number {
  const index = rows.findIndex((row) => row.end >= scrollTop);
  return index < 0 ? Math.max(rows.length - 1, 0) : index;
}

function findLastVisibleRow(rows: ReturnType<typeof measureRows>, viewportBottom: number): number {
  const index = rows.findIndex((row) => row.start > viewportBottom);
  return index < 0 ? rows.length - 1 : Math.max(0, index - 1);
}

function addressToCoord(address: string) {
  const match = address.match(/^([A-Z]+)(\d+)$/);
  if (!match) {
    return { row: 0, column: 0 };
  }

  const column = match[1].split("").reduce((total, character) => total * 26 + character.charCodeAt(0) - 64, 0) - 1;
  return { row: Number(match[2]) - 1, column };
}
