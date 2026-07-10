import {
  useCallback,
  useEffect,
  useMemo,
  useRef,
  useState,
  type CSSProperties,
  type RefObject
} from "react";
import { measureAxis } from "../core/viewport/axis";
import { columnIndexToName, formatCellAddress, getRangeAddresses, normalizeRange } from "../lib/addressing";
import { getConditionalDataBarForValue, getConditionalFormatForValue } from "../lib/conditionalFormatting";
import { formatDisplayValue } from "../lib/displayFormat";
import {
  BLANK_FILTER_VALUE,
  EMPTY_RESULT_FILTER_VALUE,
  compareDeterministicText,
  foldDeterministicText,
  getVisibleRows
} from "../lib/filters";
import type { FormulaEngine } from "../lib/formulaEngine";
import {
  DEFAULT_COLUMN_WIDTH,
  DEFAULT_ROW_HEIGHT,
  clampColumnWidth,
  clampRowHeight
} from "../lib/sheetDimensions";
import { validateCellCandidate } from "../lib/validation";
import { GridViewport } from "../react/viewport/GridViewport";
import type {
  GridEditorState,
  GridViewportApi,
  GridViewportCell,
  GridViewportColumn,
  GridViewportInteraction,
  GridViewportRenderContext,
  GridViewportRow
} from "../react/viewport/types";
import type { TableCellRef, TableSelection } from "../table/core/types";
import type {
  CellFormat,
  CellRange,
  ConditionalFormatRule,
  DataValidationRule,
  SheetFilter,
  SheetModel,
  TableStyle
} from "../types";
import {
  SpreadsheetCell,
  type SpreadsheetAutoFilter,
  type SpreadsheetAutoFilterChoice,
  type SpreadsheetMergeInfo
} from "./grid/SpreadsheetCell";
import { SpreadsheetColumnHeader, SpreadsheetRowHeader } from "./grid/SpreadsheetGridHeaders";

const ROW_OVERSCAN = 16;
const COLUMN_OVERSCAN = 2;
const ROW_HEADER_WIDTH = 48;
const COLUMN_HEADER_HEIGHT = 28;
const AUTO_SCROLL_MAX_STEP = 48;

export type CommitEditMove = "down" | "up" | "right" | "left";

export type GridScrollApi = {
  ensureCellVisible: (row: number, column: number) => void;
  focusCell: (row: number, column: number) => void;
};

export type StructuredTableCellProjection = {
  tableId: string;
  columnId: string;
  rowId?: string;
  role: "header" | "body" | "totals";
  style?: TableStyle;
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
  getStructuredTableCell?: (address: string) => StructuredTableCellProjection | null;
  isStructuredTableRowVisible?: (row: number) => boolean;
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
  onColumnHeaderContextMenu?: (event: {
    column: number;
    x: number;
    y: number;
    opener: HTMLElement;
  }) => void;
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
  | { kind: "column"; index: number; startClient: number; startSize: number; size: number }
  | { kind: "row"; index: number; startClient: number; startSize: number; size: number };

type HeaderDrag = { kind: "columns"; anchor: number } | { kind: "rows"; anchor: number };
type OverlayRect = { top: number; left: number; width: number; height: number };
type RowMeasurement = { row: number; height: number; start: number; end: number };
type ColumnLayout = { column: number; left: number; width: number };

type FullMergeInfo =
  | ({ role: "anchor" } & SpreadsheetMergeInfo & { range: CellRange })
  | ({ role: "covered" } & SpreadsheetMergeInfo & { range: CellRange; anchor: { row: number; column: number } });

type ResolvedSpreadsheetCell = {
  viewportCell: GridViewportCell;
  address: string;
  row: number;
  column: number;
  visibleValue: string;
  hyperlink: string;
  conditionalDataBar: { percent: number; color: string } | null;
  format: CellFormat | undefined;
  validation: DataValidationRule | null | undefined;
  validationValue: string;
  mergeInfo: FullMergeInfo | null;
  ruleValues: readonly (readonly string[])[];
  showValidationDropdown: boolean;
  autoFilter: SpreadsheetAutoFilter | null;
  structuredTableCell: StructuredTableCellProjection | null;
};

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
  getStructuredTableCell = () => null,
  isStructuredTableRowVisible = () => true,
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
  onColumnHeaderContextMenu,
  onAutoFilterColumn,
  onClearAutoFilterColumn,
  onSortAutoFilterColumn,
  onColumnResize,
  onRowResize,
  onColumnAutoFit,
  onRowAutoFit,
  onRegisterScrollApi
}: GridProps) {
  const [headerDrag, setHeaderDrag] = useState<HeaderDrag | null>(null);
  const [autoFillDrag, setAutoFillDrag] = useState<{ source: CellRange; target: CellRange } | null>(null);
  const [resizeDraft, setResizeDraft] = useState<ResizeDraft | null>(null);
  const internalScrollRef = useRef<HTMLDivElement>(null);
  const resolvedScrollRef = scrollRef ?? internalScrollRef;
  const headerDragRef = useRef<HeaderDrag | null>(null);
  const autoFillDragRef = useRef<{ source: CellRange; target: CellRange } | null>(null);
  const pointerClientRef = useRef<{ x: number; y: number } | null>(null);
  const lastDragTargetRef = useRef<{ row: number; column: number } | null>(null);
  const normalizedSelection = useMemo(() => normalizeRange(selection), [selection]);
  const selectionKey = `${selection.start.row}:${selection.start.column}:${selection.end.row}:${selection.end.column}`;
  const interactionResetKey = useMemo(() => ({ sheet, selectionKey }), [selectionKey, sheet]);
  const columnWidths = useMemo(
    () => Array.from({ length: sheet.columnCount }, (_, column) => columnWidth(sheet, column, resizeDraft)),
    [resizeDraft, sheet]
  );
  const filteredRows = useMemo(
    () =>
      getVisibleRows(sheet.rowCount, sheet.filters ?? [], (row, column) =>
        formulaEngine.getComputedValue(sheet.id, formatCellAddress({ row, column }))
      ).filter((row) => !(sheet.hiddenRows ?? {})[String(row)] && isStructuredTableRowVisible(row)),
    [formulaEngine, isStructuredTableRowVisible, sheet]
  );
  const viewportRows = useMemo<readonly GridViewportRow[]>(
    () =>
      filteredRows.map((row) => ({
        id: sheetRowId(row),
        label: String(row + 1),
        ariaLabel: `Row ${row + 1}`,
        headerAriaLabel: `Row ${row + 1}`,
        height: rowHeight(sheet, row, resizeDraft),
        kind: "data",
        ariaRowIndex: row + 2,
        pinned: freezeTopRow && row === 0 ? "top" : undefined
      })),
    [filteredRows, freezeTopRow, resizeDraft, sheet]
  );
  const viewportColumns = useMemo<readonly GridViewportColumn[]>(
    () =>
      Array.from({ length: sheet.columnCount }, (_, column) => column)
        .filter((column) => !(sheet.hiddenColumns ?? {})[String(column)])
        .map((column) => ({
          id: sheetColumnId(column),
          label: `Column ${columnIndexToName(column)}`,
          width: columnWidths[column],
          minWidth: 24,
          maxWidth: 640,
          pinned: freezeFirstColumn && column === 0 ? "left" : undefined,
          ariaColumnIndex: column + (showHeaders ? 2 : 1)
        })),
    [columnWidths, freezeFirstColumn, sheet.columnCount, sheet.hiddenColumns, showHeaders]
  );
  const rowAxisMeasurements = useMemo(
    () => measureAxis(viewportRows.length, (index) => viewportRows[index].id, (index) => viewportRows[index].height),
    [viewportRows]
  );
  const columnAxisMeasurements = useMemo(
    () =>
      measureAxis(
        viewportColumns.length,
        (index) => viewportColumns[index].id,
        (index) => viewportColumns[index].width
      ),
    [viewportColumns]
  );
  const rowMeasurements = useMemo<RowMeasurement[]>(
    () =>
      rowAxisMeasurements.map((measurement) => ({
        row: parseSheetIndex(measurement.key, "row:"),
        height: measurement.size,
        start: measurement.start,
        end: measurement.end
      })),
    [rowAxisMeasurements]
  );
  const columnLayout = useMemo<ColumnLayout[]>(
    () =>
      columnAxisMeasurements.map((measurement) => ({
        column: parseSheetIndex(measurement.key, "column:"),
        left: measurement.start,
        width: measurement.size
      })),
    [columnAxisMeasurements]
  );
  const viewportSelection = useMemo<TableSelection>(
    () => ({ anchor: sheetCellRef(selection.start), focus: sheetCellRef(selection.end) }),
    [selection]
  );
  const activeViewportCell = useMemo(() => sheetCellRef(selection.start), [selection.start.column, selection.start.row]);
  const viewportEditing = useMemo<GridEditorState | null>(() => {
    if (!editingCell) {
      return null;
    }
    const coordinate = addressToCoord(editingCell.address);
    return { ...sheetCellRef(coordinate), rawText: editingCell.value };
  }, [editingCell]);
  const retainedRowIds = useMemo(
    () => [
      ...new Set([
        ...(sheet.merges ?? []).map((merge) => sheetRowId(normalizeRange(merge.range).start.row)),
        sheetRowId(selection.start.row),
        sheetRowId(selection.end.row)
      ])
    ],
    [selection.end.row, selection.start.row, sheet.merges]
  );
  const retainedColumnIds = useMemo(
    () => [
      ...new Set([
        ...(sheet.merges ?? []).map((merge) => sheetColumnId(normalizeRange(merge.range).start.column)),
        sheetColumnId(selection.start.column),
        sheetColumnId(selection.end.column)
      ])
    ],
    [selection.end.column, selection.start.column, sheet.merges]
  );
  const conditionalRuleValuesCache = useMemo(() => new Map<string, string[]>(), [formulaEngine, sheet]);
  const resolvedCellCache = useMemo(
    () => new Map<string, ResolvedSpreadsheetCell>(),
    [
      editingCell,
      formulaEngine,
      getCellComment,
      getCellConditionalFormatRules,
      getCellFormat,
      getCellHyperlink,
      getCellReadOnly,
      getCellValidation,
      getStructuredTableCell,
      normalizedSelection,
      onAutoFilterColumn,
      onClearAutoFilterColumn,
      onSortAutoFilterColumn,
      sheet,
      showFormulas,
      freezeFirstColumn,
      freezeTopRow,
      columnWidths
    ]
  );

  function getConditionalRuleValues(rule: ConditionalFormatRule): readonly string[] {
    const cached = conditionalRuleValuesCache.get(rule.id);
    if (cached) {
      return cached;
    }
    const values = getRangeAddresses(rule.range).map((address) => formulaEngine.getDisplayValue(sheet.id, address));
    conditionalRuleValuesCache.set(rule.id, values);
    return values;
  }

  function resolveCell(rowId: string, columnId: string): ResolvedSpreadsheetCell {
    const cacheKey = `${rowId}|${columnId}`;
    const cached = resolvedCellCache.get(cacheKey);
    if (cached) {
      return cached;
    }
    const row = parseSheetIndex(rowId, "row:");
    const column = parseSheetIndex(columnId, "column:");
    const address = formatCellAddress({ row, column });
    const mergeInfo = getMergeInfo(sheet, row, column);
    const isMergeAnchor = mergeInfo?.role === "anchor";
    const isMergeCovered = mergeInfo?.role === "covered";
    const structuredTableCell = isMergeCovered ? null : getStructuredTableCell(address);
    const isStructuredTableStriped = structuredTableCell?.role === "body"
      && structuredTableCell.style?.showRowStripes !== false
      && row % 2 === 1;
    const format = getCellFormat(address);
    const comment = getCellComment(address)?.trim() ?? "";
    const hyperlink = isMergeCovered ? "" : getCellHyperlink(address)?.trim() ?? "";
    const isReadOnly = !isMergeCovered && getCellReadOnly(address);
    const validation = getCellValidation(address);
    const rawContent = isMergeCovered ? null : formulaEngine.getRawContent(sheet.id, address);
    const rawDisplayValue = isMergeCovered ? "" : formulaEngine.getDisplayValue(sheet.id, address);
    const formulaText = typeof rawContent === "string" && rawContent.startsWith("=") ? rawContent : "";
    const conditionalRules = getCellConditionalFormatRules(address);
    const ruleValues = conditionalRules.map((rule) => getConditionalRuleValues(rule));
    const conditionalFormat = getConditionalFormatForValue(rawDisplayValue, conditionalRules, {
      getRuleValues: getConditionalRuleValues
    });
    const conditionalDataBar = getConditionalDataBarForValue(rawDisplayValue, conditionalRules, {
      getRuleValues: getConditionalRuleValues
    });
    const mergedFormat = conditionalFormat ? { ...(format ?? {}), ...conditionalFormat } : format;
    const displayValue = formatDisplayValue(rawDisplayValue, mergedFormat);
    const visibleValue = showFormulas && formulaText ? formulaText : hyperlink && displayValue === "" ? hyperlink : displayValue;
    const isSelected = isInSelection(row, column, normalizedSelection);
    const isActive = row === selection.start.row && column === selection.start.column;
    const isEditing = editingCell?.address === address && !isMergeCovered;
    const validationCandidate =
      validation && !isMergeCovered && !isEditing
        ? {
            raw: rawContent === null ? "" : String(rawContent),
            parsed: rawContent,
            evaluated: formulaEngine.getComputedValue(sheet.id, address),
            formula: formulaText || undefined
          }
        : null;
    const isInvalidValidation =
      validationCandidate !== null && !validateCellCandidate(validationCandidate, validation).valid;
    const validationValue = rawContent === null ? "" : String(rawContent);
    const showValidationDropdown =
      !isMergeCovered &&
      !isReadOnly &&
      !isEditing &&
      isSelected &&
      validation?.type === "list" &&
      validation.values.length > 0;
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
      autoFilterRange && isAutoFilterHeader
        ? findAutoFilterForColumn(sheet.filters ?? [], autoFilterRange, column)
        : null;
    const autoFilter: SpreadsheetAutoFilter | null =
      autoFilterRange && isAutoFilterHeader
        ? {
            column,
            label: autoFilterLabel,
            filtered: Boolean(activeAutoFilter),
            activeValues: activeAutoFilterValues(activeAutoFilter),
            loadChoices: () => getAutoFilterColumnChoices(sheet, formulaEngine, autoFilterRange, column),
            onApply: (values) => onAutoFilterColumn?.(column, values),
            onClear: () => onClearAutoFilterColumn?.(column),
            onSort: (direction) => onSortAutoFilterColumn?.(column, direction)
          }
        : null;
    const cellWidth =
      isMergeAnchor && mergeInfo
        ? sumColumnWidths(columnWidths, mergeInfo.range.start.column, mergeInfo.range.end.column, sheet.hiddenColumns)
        : columnWidths[column];
    const cellHeight =
      isMergeAnchor && mergeInfo
        ? sumRowHeights(sheet, mergeInfo.range.start.row, mergeInfo.range.end.row)
        : rowHeight(sheet, row, resizeDraft);
    const borderStyle = getBorderStyle(mergedFormat?.borders);
    const style: CSSProperties = {
      fontWeight: mergedFormat?.bold ? 700 : undefined,
      fontStyle: mergedFormat?.italic ? "italic" : undefined,
      fontFamily: mergedFormat?.fontFamily,
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
    };
    const className = [
      "cell",
      isSelected ? "selected-cell" : "",
      isActive ? "active-cell" : "",
      isEditing ? "editing-cell" : "",
      isMergeAnchor ? "merged-cell" : "",
      isMergeCovered ? "merge-covered-cell" : "",
      isInvalidValidation ? "invalid-validation-cell" : "",
      comment && !isMergeCovered ? "commented-cell" : "",
      hyperlink ? "hyperlink-cell" : "",
      isAutoFilterHeader ? "auto-filter-header-cell" : "",
      activeAutoFilter ? "filtered-header-cell" : "",
      showValidationDropdown ? "validation-list-cell" : "",
      mergedFormat?.wrapText ? "wrapped-cell" : "",
      isReadOnly ? "read-only-cell" : "",
      conditionalFormat ? "conditional-format-cell" : "",
      conditionalDataBar ? "conditional-data-bar-cell" : "",
      structuredTableCell ? "structured-table-cell" : "",
      structuredTableCell ? `structured-table-cell--${structuredTableCell.role}` : "",
      isStructuredTableStriped ? "structured-table-cell--striped" : "",
      freezeTopRow && row === 0 ? "frozen-top-row-cell" : "",
      freezeFirstColumn && column === 0 ? "frozen-first-column-cell" : ""
    ]
      .filter(Boolean)
      .join(" ");
    const title = [
      isReadOnly ? "Read only" : "",
      comment ? `Comment: ${comment}` : "",
      hyperlink ? `Link: ${hyperlink}` : ""
    ]
      .filter(Boolean)
      .join("\n");
    const resolved: ResolvedSpreadsheetCell = {
      viewportCell: {
        ref: { rowId, columnId },
        ariaLabel: visibleValue ? `${address} ${visibleValue}` : address,
        displayValue: visibleValue,
        editable: !isReadOnly && !isMergeCovered,
        invalid: isInvalidValidation,
        className,
        style,
        title: title || undefined,
        dataAttributes: structuredTableCell ? {
          "data-structured-table-id": structuredTableCell.tableId,
          "data-structured-table-column-id": structuredTableCell.columnId,
          "data-structured-table-row-id": structuredTableCell.rowId,
          "data-structured-table-role": structuredTableCell.role,
          "data-structured-table-style": structuredTableCell.style?.theme,
          "data-striped": isStructuredTableStriped ? "true" : undefined
        } : undefined,
        columnSpan: isMergeAnchor ? mergeInfo.columnSpan : undefined,
        rowSpan: isMergeAnchor ? mergeInfo.rowSpan : undefined
      },
      address,
      row,
      column,
      visibleValue,
      hyperlink,
      conditionalDataBar,
      format: mergedFormat,
      validation,
      validationValue,
      mergeInfo,
      ruleValues,
      showValidationDropdown,
      autoFilter,
      structuredTableCell
    };
    resolvedCellCache.set(cacheKey, resolved);
    return resolved;
  }

  const handleViewportInteraction = useCallback(
    (interaction: GridViewportInteraction) => {
      switch (interaction.type) {
        case "selection-change":
          onSelectionChange({
            start: viewportRefToCoordinate(interaction.selection.anchor),
            end: viewportRefToCoordinate(interaction.selection.focus)
          });
          return;
        case "edit-start":
          onStartEdit(formatCellAddress(viewportRefToCoordinate(interaction.cell)));
          return;
        case "edit-change":
          onEditValueChange(interaction.rawText);
          return;
        case "edit-commit":
          onCommitEdit(
            formatCellAddress(viewportRefToCoordinate(interaction.cell)),
            interaction.rawText,
            interaction.move
          );
          return;
        case "edit-cancel":
          onCancelEdit();
          return;
        case "paste":
          onPasteText(interaction.text);
          return;
        case "column-resize":
          onColumnResize?.(parseSheetIndex(interaction.columnId, "column:"), interaction.width);
          return;
        case "copy":
          return;
      }
    },
    [onCancelEdit, onColumnResize, onCommitEdit, onEditValueChange, onPasteText, onSelectionChange, onStartEdit]
  );

  const renderSpreadsheetCell = useCallback(
    (context: GridViewportRenderContext, mode: "display" | "editor") => {
      const resolved = resolveCell(context.row.id, context.column.id);
      return (
        <SpreadsheetCell
          context={context}
          mode={mode}
          address={resolved.address}
          visibleValue={resolved.visibleValue}
          hyperlink={resolved.hyperlink}
          conditionalDataBar={resolved.conditionalDataBar}
          format={resolved.format}
          validation={resolved.validation}
          validationValue={resolved.validationValue}
          mergeInfo={resolved.mergeInfo}
          ruleValues={resolved.ruleValues}
          showValidationDropdown={resolved.showValidationDropdown}
          editorValue={mode === "editor" ? editingCell?.value ?? "" : undefined}
          autoFilter={mode === "display" ? resolved.autoFilter : null}
          structuredTableCell={mode === "display" ? resolved.structuredTableCell : null}
          interactionResetKey={interactionResetKey}
          onEditValueChange={onEditValueChange}
          onCommitEdit={onCommitEdit}
          onCancelEdit={onCancelEdit}
        />
      );
    },
    [editingCell?.value, interactionResetKey, onCancelEdit, onCommitEdit, onEditValueChange, resolvedCellCache]
  );

  const setFillDrag = useCallback((next: { source: CellRange; target: CellRange } | null) => {
    autoFillDragRef.current = next;
    setAutoFillDrag(next);
  }, []);

  const setCurrentHeaderDrag = useCallback((next: HeaderDrag | null) => {
    headerDragRef.current = next;
    setHeaderDrag(next);
  }, []);

  const latestRef = useRef({
    sheet,
    selectionStart: selection.start,
    normalizedSelection,
    rowMeasurements,
    columnLayout,
    showHeaders,
    zoomLevel,
    freezeTopRow,
    freezeFirstColumn,
    onSelectionChange,
    onAutoFill
  });
  latestRef.current = {
    sheet,
    selectionStart: selection.start,
    normalizedSelection,
    rowMeasurements,
    columnLayout,
    showHeaders,
    zoomLevel,
    freezeTopRow,
    freezeFirstColumn,
    onSelectionChange,
    onAutoFill
  };

  const applySpecialDragTarget = useCallback(
    (coordinate: { row: number; column: number }) => {
      const current = latestRef.current;
      const fill = autoFillDragRef.current;
      lastDragTargetRef.current = coordinate;
      if (fill) {
        const target = createAutoFillTarget(fill.source, coordinate.row, coordinate.column);
        if (!rangesEqual(target, fill.target)) {
          setFillDrag({ source: fill.source, target });
        }
        return;
      }
      const drag = headerDragRef.current;
      if (drag?.kind === "columns") {
        current.onSelectionChange({
          start: { row: 0, column: drag.anchor },
          end: { row: current.sheet.rowCount - 1, column: coordinate.column }
        });
      } else if (drag?.kind === "rows") {
        current.onSelectionChange({
          start: { row: drag.anchor, column: 0 },
          end: { row: coordinate.row, column: current.sheet.columnCount - 1 }
        });
      }
    },
    [setFillDrag]
  );

  const cellAtClientPoint = useCallback(
    (clientX: number, clientY: number): { row: number; column: number } | null => {
      const element = resolvedScrollRef.current;
      if (!element) {
        return null;
      }
      const current = latestRef.current;
      const layout = current.columnLayout;
      const measurements = current.rowMeasurements;
      if (layout.length === 0 || measurements.length === 0) {
        return null;
      }
      const rect = element.getBoundingClientRect();
      const zoom = (current.zoomLevel || 100) / 100;
      const headerWidth = current.showHeaders ? ROW_HEADER_WIDTH : 0;
      const headerHeight = current.showHeaders ? COLUMN_HEADER_HEIGHT : 0;
      const contentX = (clientX - rect.left + element.scrollLeft) / zoom - headerWidth;
      const contentY = (clientY - rect.top + element.scrollTop) / zoom - headerHeight;
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
      return {
        row: frozenRowHit ? 0 : rowAtOffset(measurements, contentY),
        column: frozenColumnHit ? 0 : columnAtOffset(layout, contentX)
      };
    },
    [resolvedScrollRef]
  );

  const updateDragTargetFromPointer = useCallback(() => {
    const pointer = pointerClientRef.current;
    if (!pointer) {
      return;
    }
    const raw = cellAtClientPoint(pointer.x, pointer.y);
    if (!raw) {
      return;
    }
    const drag = headerDragRef.current;
    const coordinate =
      drag?.kind === "columns"
        ? { row: 0, column: raw.column }
        : drag?.kind === "rows"
          ? { row: raw.row, column: 0 }
          : raw;
    const last = lastDragTargetRef.current;
    if (last && last.row === coordinate.row && last.column === coordinate.column) {
      return;
    }
    applySpecialDragTarget(coordinate);
  }, [applySpecialDragTarget, cellAtClientPoint]);

  const finalizeDrag = useCallback(() => {
    const fill = autoFillDragRef.current;
    if (fill) {
      autoFillDragRef.current = null;
      setAutoFillDrag(null);
      if (!rangesEqual(fill.source, fill.target)) {
        latestRef.current.onAutoFill?.(fill.source, fill.target);
      }
    }
    if (headerDragRef.current) {
      setCurrentHeaderDrag(null);
    }
    lastDragTargetRef.current = null;
    pointerClientRef.current = null;
  }, [setCurrentHeaderDrag]);

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
        const size =
          current.kind === "column"
            ? clampColumnWidth(current.startSize + delta)
            : clampRowHeight(current.startSize + delta);
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

  const specialDragActive = Boolean(headerDrag) || Boolean(autoFillDrag);
  useEffect(() => {
    if (!specialDragActive) {
      return undefined;
    }
    function handleWindowMouseMove(event: MouseEvent) {
      pointerClientRef.current = { x: event.clientX, y: event.clientY };
      updateDragTargetFromPointer();
    }
    function handleWindowMouseUp() {
      finalizeDrag();
    }
    let frame: number | null = null;
    const autoScroll = () => {
      const element = resolvedScrollRef.current;
      const pointer = pointerClientRef.current;
      if (element && pointer) {
        const rect = element.getBoundingClientRect();
        const zoom = (latestRef.current.zoomLevel || 100) / 100;
        const innerLeft = rect.left + (latestRef.current.showHeaders ? ROW_HEADER_WIDTH * zoom : 0);
        const innerTop = rect.top + (latestRef.current.showHeaders ? COLUMN_HEADER_HEIGHT * zoom : 0);
        const dragKind = headerDragRef.current?.kind;
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
      frame = requestAnimationFrame(autoScroll);
    };
    frame = requestAnimationFrame(autoScroll);
    window.addEventListener("mousemove", handleWindowMouseMove);
    window.addEventListener("mouseup", handleWindowMouseUp);
    return () => {
      window.removeEventListener("mousemove", handleWindowMouseMove);
      window.removeEventListener("mouseup", handleWindowMouseUp);
      if (frame !== null) {
        cancelAnimationFrame(frame);
      }
    };
  }, [finalizeDrag, resolvedScrollRef, specialDragActive, updateDragTargetFromPointer]);

  const registerViewportApi = useCallback(
    (api: GridViewportApi) => {
      onRegisterScrollApi?.({
        ensureCellVisible: (row, column) => api.ensureCellVisible(sheetRowId(row), sheetColumnId(column)),
        focusCell: (row, column) => api.focusCell(sheetRowId(row), sheetColumnId(column))
      });
    },
    [onRegisterScrollApi]
  );

  const headerWidth = showHeaders ? ROW_HEADER_WIDTH : 0;
  const headerHeight = showHeaders ? COLUMN_HEADER_HEIGHT : 0;
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
  const fillPreviewRect = useMemo(
    () =>
      autoFillDrag
        ? rangeOverlayRect(
            normalizeRange(autoFillDrag.target),
            rowMeasurements,
            columnLayout,
            headerWidth,
            headerHeight
          )
        : null,
    [autoFillDrag, columnLayout, headerHeight, headerWidth, rowMeasurements]
  );
  const copiedRect = useMemo(
    () =>
      copiedRange
        ? rangeOverlayRect(normalizeRange(copiedRange), rowMeasurements, columnLayout, headerWidth, headerHeight)
        : null,
    [columnLayout, copiedRange, headerHeight, headerWidth, rowMeasurements]
  );
  const showFillHandle = !editingCell && !autoFillDrag && Boolean(selectionRect) && Boolean(onAutoFill);
  const isWholeSheetSelected =
    normalizedSelection.start.row === 0 &&
    normalizedSelection.start.column === 0 &&
    normalizedSelection.end.row === sheet.rowCount - 1 &&
    normalizedSelection.end.column === sheet.columnCount - 1;
  const zoomStyle = {
    "--js-spreadsheet-sheet-zoom": String(zoomLevel / 100),
    "--js-spreadsheet-row-header-width": showHeaders ? `${ROW_HEADER_WIDTH}px` : "0px",
    "--js-spreadsheet-column-header-height": showHeaders ? `${COLUMN_HEADER_HEIGHT}px` : "0px"
  } as CSSProperties;
  const rootClassName = ["grid-scroll", showGridlines ? "" : "grid-scroll--no-gridlines"]
    .filter(Boolean)
    .join(" ");

  function openColumnHeaderContextMenu(column: number, x: number, y: number, opener: HTMLElement) {
    const preservesWholeColumnSelection =
      normalizedSelection.start.row === 0 &&
      normalizedSelection.end.row === sheet.rowCount - 1 &&
      column >= normalizedSelection.start.column &&
      column <= normalizedSelection.end.column;
    if (!preservesWholeColumnSelection) {
      onSelectionChange({
        start: { row: 0, column },
        end: { row: sheet.rowCount - 1, column }
      });
    }
    onColumnHeaderContextMenu?.({ column, x, y, opener });
  }

  return (
    <GridViewport
      idPrefix={`spreadsheet-${sheet.id}`}
      ariaLabel="Spreadsheet grid"
      rows={viewportRows}
      ariaRowCount={sheet.rowCount + 1}
      columns={viewportColumns}
      ariaColumnCount={sheet.columnCount + (showHeaders ? 1 : 0)}
      getCell={(rowId, columnId) => resolveCell(rowId, columnId).viewportCell}
      selection={viewportSelection}
      activeCell={activeViewportCell}
      editing={viewportEditing}
      onInteraction={handleViewportInteraction}
      renderCell={(context) => renderSpreadsheetCell(context, "display")}
      renderEditor={(context) => renderSpreadsheetCell(context, "editor")}
      renderColumnHeader={(column) => {
        const columnIndex = parseSheetIndex(column.id, "column:");
        return (
          <SpreadsheetColumnHeader
            column={columnIndex}
            label={columnIndexToName(columnIndex)}
            onStartResize={(index, event) =>
              setResizeDraft({
                kind: "column",
                index,
                startClient: event.clientX,
                startSize: columnWidth(sheet, index, null),
                size: columnWidth(sheet, index, null)
              })
            }
            onAutoFit={onColumnAutoFit}
          />
        );
      }}
      renderRowHeader={
        showHeaders
          ? (row) => {
              const rowIndex = parseSheetIndex(row.id, "row:");
              return (
                <SpreadsheetRowHeader
                  row={rowIndex}
                  label={String(rowIndex + 1)}
                  onStartResize={(index, event) =>
                    setResizeDraft({
                      kind: "row",
                      index,
                      startClient: event.clientY,
                      startSize: rowHeight(sheet, index, null),
                      size: rowHeight(sheet, index, null)
                    })
                  }
                  onAutoFit={onRowAutoFit}
                />
              );
            }
          : undefined
      }
      renderCornerHeader={
        showHeaders
          ? () => (
              <button
                type="button"
                className="corner-cell"
                aria-label="Select sheet"
                aria-pressed={isWholeSheetSelected}
                style={{ position: "relative", width: "100%", height: "100%" }}
                onClick={() =>
                  onSelectionChange({
                    start: { row: 0, column: 0 },
                    end: { row: sheet.rowCount - 1, column: sheet.columnCount - 1 }
                  })
                }
              />
            )
          : undefined
      }
      renderOverlay={() => (
        <>
          {selectionRect && !editingCell ? (
            <div className="selection-outline" aria-hidden="true" style={selectionRect} />
          ) : null}
          {copiedRect ? <div className="copy-marquee" aria-hidden="true" style={copiedRect} /> : null}
          {fillPreviewRect ? (
            <div className="fill-preview-outline" aria-hidden="true" style={fillPreviewRect} />
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
                setCurrentHeaderDrag(null);
                const source = latestRef.current.normalizedSelection;
                setFillDrag({ source, target: source });
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
        </>
      )}
      scrollRef={resolvedScrollRef}
      showColumnHeaders={showHeaders}
      rowHeaderWidth={ROW_HEADER_WIDTH}
      columnHeaderHeight={COLUMN_HEADER_HEIGHT}
      rowOverscan={ROW_OVERSCAN}
      columnOverscan={COLUMN_OVERSCAN}
      scale={(zoomLevel || 100) / 100}
      resetKey={sheet.id}
      retainedRowIds={retainedRowIds}
      retainedColumnIds={retainedColumnIds}
      rootClassName={rootClassName}
      rootStyle={zoomStyle}
      rootDataAttributes={{
        "data-zoom-level": zoomLevel,
        "data-gridlines": showGridlines ? "visible" : "hidden",
        "data-headers": showHeaders ? "visible" : "hidden"
      }}
      canvasClassName="spreadsheet-grid"
      interactionEventMode="mouse"
      onBeforeKeyDown={onKeyCommand}
      onCellMouseEnter={(cell) => {
        if (!autoFillDragRef.current) {
          return false;
        }
        applySpecialDragTarget(viewportRefToCoordinate(cell));
        return true;
      }}
      onCellContextMenu={(cell, event) => {
        event.preventDefault();
        const coordinate = viewportRefToCoordinate(cell);
        const address = formatCellAddress(coordinate);
        if (!isInSelection(coordinate.row, coordinate.column, normalizedSelection)) {
          onSelectionChange({ start: coordinate, end: coordinate });
        }
        onCellContextMenu?.({
          address,
          row: coordinate.row,
          column: coordinate.column,
          x: event.clientX,
          y: event.clientY
        });
      }}
      onColumnHeaderContextMenu={onColumnHeaderContextMenu
        ? (column, event) => {
            event.preventDefault();
            openColumnHeaderContextMenu(
              parseSheetIndex(column.id, "column:"),
              event.clientX,
              event.clientY,
              event.currentTarget
            );
          }
        : undefined}
      onReadOnlyCellEditAttempt={(cell) => {
        const resolved = resolveCell(cell.rowId, cell.columnId);
        if (resolved.mergeInfo?.role !== "covered") {
          onStartEdit(resolved.address);
        }
      }}
      getColumnHeaderState={(column) => {
        const index = parseSheetIndex(column.id, "column:");
        const selected =
          normalizedSelection.start.row === 0 &&
          normalizedSelection.end.row === sheet.rowCount - 1 &&
          index >= normalizedSelection.start.column &&
          index <= normalizedSelection.end.column;
        const hit =
          !selected && index >= normalizedSelection.start.column && index <= normalizedSelection.end.column;
        return {
          className: ["column-header", selected ? "selected-header" : "", hit ? "column-header--hit" : ""]
            .filter(Boolean)
            .join(" "),
          ariaSelected: selected,
          tabIndex: 0
        };
      }}
      getRowHeaderState={(row) => {
        const index = parseSheetIndex(row.id, "row:");
        const selected =
          normalizedSelection.start.column === 0 &&
          normalizedSelection.end.column === sheet.columnCount - 1 &&
          index >= normalizedSelection.start.row &&
          index <= normalizedSelection.end.row;
        const hit = !selected && index >= normalizedSelection.start.row && index <= normalizedSelection.end.row;
        return {
          className: ["row-header", selected ? "selected-header" : "", hit ? "row-header--hit" : ""]
            .filter(Boolean)
            .join(" "),
          ariaSelected: selected,
          tabIndex: 0
        };
      }}
      onColumnHeaderMouseDown={(column, event) => {
        if (event.button !== 0) {
          return;
        }
        event.preventDefault();
        const index = parseSheetIndex(column.id, "column:");
        const anchor = event.shiftKey ? selection.start.column : index;
        setCurrentHeaderDrag({ kind: "columns", anchor });
        onSelectionChange({
          start: { row: 0, column: anchor },
          end: { row: sheet.rowCount - 1, column: index }
        });
      }}
      onColumnHeaderMouseEnter={(column) => {
        if (headerDragRef.current?.kind === "columns") {
          applySpecialDragTarget({ row: 0, column: parseSheetIndex(column.id, "column:") });
        }
      }}
      onColumnHeaderClick={(column, event) => {
        const index = parseSheetIndex(column.id, "column:");
        const anchor = event.shiftKey ? selection.start.column : index;
        onSelectionChange({
          start: { row: 0, column: anchor },
          end: { row: sheet.rowCount - 1, column: index }
        });
      }}
      onColumnHeaderKeyDown={(column, event) => {
        if (
          onColumnHeaderContextMenu &&
          ((event.key === "F10" && event.shiftKey) || event.key === "ContextMenu")
        ) {
          event.preventDefault();
          event.stopPropagation();
          const rect = event.currentTarget.getBoundingClientRect();
          openColumnHeaderContextMenu(
            parseSheetIndex(column.id, "column:"),
            rect.left,
            rect.bottom,
            event.currentTarget
          );
          return;
        }
        if (event.key === "Enter" || event.key === " ") {
          event.preventDefault();
          const index = parseSheetIndex(column.id, "column:");
          const anchor = event.shiftKey ? selection.start.column : index;
          onSelectionChange({
            start: { row: 0, column: anchor },
            end: { row: sheet.rowCount - 1, column: index }
          });
        }
      }}
      onRowHeaderMouseDown={(row, event) => {
        if (event.button !== 0) {
          return;
        }
        event.preventDefault();
        const index = parseSheetIndex(row.id, "row:");
        const anchor = event.shiftKey ? selection.start.row : index;
        setCurrentHeaderDrag({ kind: "rows", anchor });
        onSelectionChange({
          start: { row: anchor, column: 0 },
          end: { row: index, column: sheet.columnCount - 1 }
        });
      }}
      onRowHeaderMouseEnter={(row) => {
        if (headerDragRef.current?.kind === "rows") {
          applySpecialDragTarget({ row: parseSheetIndex(row.id, "row:"), column: 0 });
        }
      }}
      onRowHeaderClick={(row, event) => {
        const index = parseSheetIndex(row.id, "row:");
        const anchor = event.shiftKey ? selection.start.row : index;
        onSelectionChange({
          start: { row: anchor, column: 0 },
          end: { row: index, column: sheet.columnCount - 1 }
        });
      }}
      onRowHeaderKeyDown={(row, event) => {
        if (event.key === "Enter" || event.key === " ") {
          event.preventDefault();
          const index = parseSheetIndex(row.id, "row:");
          const anchor = event.shiftKey ? selection.start.row : index;
          onSelectionChange({
            start: { row: anchor, column: 0 },
            end: { row: index, column: sheet.columnCount - 1 }
          });
        }
      }}
      onRootMouseUp={finalizeDrag}
      onRegisterApi={registerViewportApi}
    />
  );
}

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
    if (row < normalized.end.row) {
      return { start: normalized.start, end: { row, column: normalized.end.column } };
    }
    if (column < normalized.end.column) {
      return { start: normalized.start, end: { row: normalized.end.row, column } };
    }
    return normalized;
  }
  if (Math.abs(rowOvershoot) >= Math.abs(columnOvershoot)) {
    return rowOvershoot > 0
      ? { start: normalized.start, end: { row, column: normalized.end.column } }
      : { start: { row, column: normalized.start.column }, end: normalized.end };
  }
  return columnOvershoot > 0
    ? { start: normalized.start, end: { row: normalized.end.row, column } }
    : { start: { row: normalized.start.row, column }, end: normalized.end };
}

function sheetRowId(row: number): string {
  return `row:${row}`;
}

function sheetColumnId(column: number): string {
  return `column:${column}`;
}

function parseSheetIndex(id: string, prefix: "row:" | "column:"): number {
  const value = id.startsWith(prefix) ? id.slice(prefix.length) : "";
  if (!/^\d+$/.test(value) || !Number.isInteger(Number(value))) {
    throw new Error(`Invalid spreadsheet viewport id: ${id}`);
  }
  return Number(value);
}

function sheetCellRef(coordinate: { row: number; column: number }): TableCellRef {
  return { rowId: sheetRowId(coordinate.row), columnId: sheetColumnId(coordinate.column) };
}

function viewportRefToCoordinate(cell: TableCellRef): { row: number; column: number } {
  return {
    row: parseSheetIndex(cell.rowId, "row:"),
    column: parseSheetIndex(cell.columnId, "column:")
  };
}

function isInSelection(row: number, column: number, selection: CellRange): boolean {
  return (
    row >= selection.start.row &&
    row <= selection.end.row &&
    column >= selection.start.column &&
    column <= selection.end.column
  );
}

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
      if (!rangesEqual(next, current)) {
        current = next;
        changed = true;
      }
    }
  }
  return current;
}

function rangeOverlayRect(
  range: CellRange,
  rows: readonly RowMeasurement[],
  columns: readonly ColumnLayout[],
  headerWidth: number,
  headerHeight: number
): OverlayRect | null {
  const rowRect = rowRangeRect(rows, range.start.row, range.end.row);
  const columnRect = columnRangeRect(columns, range.start.column, range.end.column);
  return rowRect && columnRect
    ? {
        top: headerHeight + rowRect.top,
        left: headerWidth + columnRect.left,
        width: columnRect.width,
        height: rowRect.height
      }
    : null;
}

function rowRangeRect(
  measurements: readonly RowMeasurement[],
  startRow: number,
  endRow: number
): { top: number; height: number } | null {
  let low = 0;
  let high = measurements.length - 1;
  while (low <= high) {
    const middle = (low + high) >> 1;
    if (measurements[middle].row < startRow) {
      low = middle + 1;
    } else {
      high = middle - 1;
    }
  }
  const firstIndex = low;
  const first = measurements[firstIndex];
  low = 0;
  high = measurements.length - 1;
  while (low <= high) {
    const middle = (low + high) >> 1;
    if (measurements[middle].row <= endRow) {
      low = middle + 1;
    } else {
      high = middle - 1;
    }
  }
  const lastIndex = high;
  const last = measurements[lastIndex];
  return firstIndex <= lastIndex && first && last ? { top: first.start, height: last.end - first.start } : null;
}

function columnRangeRect(
  layout: readonly ColumnLayout[],
  startColumn: number,
  endColumn: number
): { left: number; width: number } | null {
  let low = 0;
  let high = layout.length - 1;
  while (low <= high) {
    const middle = (low + high) >> 1;
    if (layout[middle].column < startColumn) {
      low = middle + 1;
    } else {
      high = middle - 1;
    }
  }
  const firstIndex = low;
  const first = layout[firstIndex];
  low = 0;
  high = layout.length - 1;
  while (low <= high) {
    const middle = (low + high) >> 1;
    if (layout[middle].column <= endColumn) {
      low = middle + 1;
    } else {
      high = middle - 1;
    }
  }
  const lastIndex = high;
  const last = layout[lastIndex];
  return firstIndex <= lastIndex && first && last
    ? { left: first.left, width: last.left + last.width - first.left }
    : null;
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

function getAutoFilterColumnChoices(
  sheet: SheetModel,
  formulaEngine: FormulaEngine,
  range: CellRange,
  column: number
): SpreadsheetAutoFilterChoice[] {
  const choices = new Map<string, SpreadsheetAutoFilterChoice>();
  for (let row = range.start.row + 1; row <= range.end.row; row += 1) {
    const address = formatCellAddress({ row, column });
    const raw = formulaEngine.getRawContent(sheet.id, address);
    const computed = formulaEngine.getComputedValue(sheet.id, address);
    let choice: SpreadsheetAutoFilterChoice;
    if (raw === null) {
      choice = { key: BLANK_FILTER_VALUE, label: "Blank", value: BLANK_FILTER_VALUE };
    } else if (typeof raw === "string" && raw.startsWith("=") && computed === "") {
      choice = { key: EMPTY_RESULT_FILTER_VALUE, label: "Empty result", value: EMPTY_RESULT_FILTER_VALUE };
    } else {
      const value = String(formulaEngine.getDisplayValue(sheet.id, address) ?? "");
      choice = { key: `value:${foldDeterministicText(value.trim())}`, label: value, value };
    }
    if (!choices.has(choice.key)) {
      choices.set(choice.key, choice);
    }
  }
  return [...choices.values()].sort((left, right) => compareDeterministicText(left.label, right.label));
}

function rangesEqual(left: CellRange, right: CellRange): boolean {
  const a = normalizeRange(left);
  const b = normalizeRange(right);
  return (
    a.start.row === b.start.row &&
    a.start.column === b.start.column &&
    a.end.row === b.end.row &&
    a.end.column === b.end.column
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

function getMergeInfo(sheet: SheetModel, row: number, column: number): FullMergeInfo | null {
  const merge = (sheet.merges ?? []).find((candidate) => isInSelection(row, column, normalizeRange(candidate.range)));
  if (!merge) {
    return null;
  }
  const range = normalizeRange(merge.range);
  const columnSpan = range.end.column - range.start.column + 1;
  const rowSpan = range.end.row - range.start.row + 1;
  const base = { range, columnSpan, rowSpan };
  return row === range.start.row && column === range.start.column
    ? { role: "anchor", ...base }
    : { role: "covered", ...base, anchor: { ...range.start } };
}

function getBorderStyle(borders: CellFormat["borders"]): CSSProperties {
  if (!borders) {
    return {};
  }
  return {
    borderTop: borders.top ? `1px solid ${borders.top.color}` : undefined,
    borderRight: borders.right ? `1px solid ${borders.right.color}` : undefined,
    borderBottom: borders.bottom ? `1px solid ${borders.bottom.color}` : undefined,
    borderLeft: borders.left ? `1px solid ${borders.left.color}` : undefined
  };
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

function sumColumnWidths(
  columnWidths: readonly number[],
  startColumn: number,
  endColumn: number,
  hiddenColumns: SheetModel["hiddenColumns"]
): number {
  let width = 0;
  for (let column = startColumn; column <= endColumn; column += 1) {
    if (!hiddenColumns?.[String(column)]) {
      width += columnWidths[column] ?? DEFAULT_COLUMN_WIDTH;
    }
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

function addressToCoord(address: string): { row: number; column: number } {
  const match = address.match(/^([A-Z]+)(\d+)$/);
  if (!match) {
    return { row: 0, column: 0 };
  }
  return {
    row: Number(match[2]) - 1,
    column: match[1].split("").reduce((total, character) => total * 26 + character.charCodeAt(0) - 64, 0) - 1
  };
}

function rowAtOffset(measurements: readonly RowMeasurement[], offset: number): number {
  if (offset <= measurements[0].end) {
    return measurements[0].row;
  }
  let low = 0;
  let high = measurements.length - 1;
  while (low <= high) {
    const middle = (low + high) >> 1;
    const measurement = measurements[middle];
    if (offset < measurement.start) {
      high = middle - 1;
    } else if (offset >= measurement.end) {
      low = middle + 1;
    } else {
      return measurement.row;
    }
  }
  return measurements.at(-1)!.row;
}

function columnAtOffset(layout: readonly ColumnLayout[], offset: number): number {
  if (offset <= layout[0].left + layout[0].width) {
    return layout[0].column;
  }
  let low = 0;
  let high = layout.length - 1;
  while (low <= high) {
    const middle = (low + high) >> 1;
    const entry = layout[middle];
    if (offset < entry.left) {
      high = middle - 1;
    } else if (offset >= entry.left + entry.width) {
      low = middle + 1;
    } else {
      return entry.column;
    }
  }
  return layout.at(-1)!.column;
}
