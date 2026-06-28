import { useEffect, useMemo, useState, type CSSProperties, type RefObject } from "react";
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
const ROW_OVERSCAN = 8;

type GridProps = {
  sheet: SheetModel;
  formulaEngine: FormulaEngine;
  selection: CellRange;
  editingCell: { address: string; value: string } | null;
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
  onCommitEdit: (address: string, value: string) => void;
  onCancelEdit: () => void;
  onPasteText: (text: string) => void;
  onKeyCommand: (event: React.KeyboardEvent<HTMLDivElement>) => void;
  onAutoFill?: (sourceRange: CellRange, targetRange: CellRange) => void;
  onCellContextMenu?: (event: { address: string; row: number; column: number; x: number; y: number }) => void;
  onAutoFilterColumn?: (column: number, values: string[]) => void;
  onClearAutoFilterColumn?: (column: number) => void;
  onSortAutoFilterColumn?: (column: number, direction: "asc" | "desc") => void;
  onColumnResize?: (column: number, width: number) => void;
  onRowResize?: (row: number, height: number) => void;
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

export function Grid({
  sheet,
  formulaEngine,
  selection,
  editingCell,
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
  onCellContextMenu,
  onAutoFilterColumn,
  onClearAutoFilterColumn,
  onSortAutoFilterColumn,
  onColumnResize,
  onRowResize
}: GridProps) {
  const [isDragging, setIsDragging] = useState(false);
  const [autoFillDrag, setAutoFillDrag] = useState<{ source: CellRange; target: CellRange } | null>(null);
  const [resizeDraft, setResizeDraft] = useState<ResizeDraft | null>(null);
  const [viewport, setViewport] = useState({ scrollTop: 0, height: DEFAULT_VIEWPORT_HEIGHT });
  const normalizedSelection = normalizeRange(selection);
  const columns = Array.from({ length: sheet.columnCount }, (_, column) => column).filter(
    (column) => !(sheet.hiddenColumns ?? {})[String(column)]
  );
  const columnWidths = Array.from({ length: sheet.columnCount }, (_, column) => columnWidth(sheet, column, resizeDraft));
  const filteredRows = getVisibleRows(sheet.rowCount, sheet.filters ?? [], (row, column) =>
    formulaEngine.getDisplayValue(sheet.id, formatCellAddress({ row, column }))
  ).filter((row) => !(sheet.hiddenRows ?? {})[String(row)]);
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
    "--row-header-width": showHeaders ? "48px" : "0px",
    "--column-header-height": showHeaders ? "28px" : "0px"
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

  useEffect(() => {
    setViewport({ scrollTop: 0, height: scrollRef?.current?.clientHeight || DEFAULT_VIEWPORT_HEIGHT });
    if (scrollRef?.current) {
      scrollRef.current.scrollTop = 0;
      scrollRef.current.scrollLeft = 0;
    }
  }, [scrollRef, sheet.id]);

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
        const nextViewport = {
          scrollTop: event.currentTarget.scrollTop,
          height: event.currentTarget.clientHeight || DEFAULT_VIEWPORT_HEIGHT
        };
        setViewport((current) =>
          current.scrollTop === nextViewport.scrollTop && current.height === nextViewport.height ? current : nextViewport
        );
      }}
      onPaste={(event) => {
        event.preventDefault();
        onPasteText(event.clipboardData.getData("text/plain") || event.clipboardData.getData("Text"));
      }}
      onMouseLeave={() => {
        setIsDragging(false);
        setAutoFillDrag(null);
      }}
      onMouseUp={() => {
        setIsDragging(false);
        if (autoFillDrag && !rangesEqual(autoFillDrag.source, autoFillDrag.target)) {
          onAutoFill?.(autoFillDrag.source, autoFillDrag.target);
        }
        setAutoFillDrag(null);
      }}
    >
      <div
        className="spreadsheet-grid"
        style={{
          gridTemplateColumns: [showHeaders ? "48px" : "", ...columns.map((column) => `${columnWidths[column]}px`)]
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

          return (
          <div
            key={column}
            className={["column-header", isColumnSelected ? "selected-header" : ""].filter(Boolean).join(" ")}
            role="columnheader"
            aria-label={`Column ${columnName}`}
            aria-selected={isColumnSelected}
            tabIndex={0}
            style={{ width: columnWidths[column] }}
            onClick={() =>
              onSelectionChange({
                start: { row: 0, column },
                end: { row: sheet.rowCount - 1, column }
              })
            }
            onKeyDown={(event) => {
              if (event.key === "Enter" || event.key === " ") {
                event.preventDefault();
                onSelectionChange({
                  start: { row: 0, column },
                  end: { row: sheet.rowCount - 1, column }
                });
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
          <RowFragment
            key={row}
            row={row}
            rowHeight={height}
            columns={columns}
            columnWidths={columnWidths}
            sheet={sheet}
            formulaEngine={formulaEngine}
            selection={normalizedSelection}
            editingCell={editingCell}
            showHeaders={showHeaders}
            showFormulas={showFormulas}
            freezeTopRow={freezeTopRow}
            freezeFirstColumn={freezeFirstColumn}
            getCellFormat={getCellFormat}
            getCellComment={getCellComment}
            getCellHyperlink={getCellHyperlink}
            getCellReadOnly={getCellReadOnly}
            getCellValidation={getCellValidation}
            getCellConditionalFormatRules={getCellConditionalFormatRules}
            getConditionalRuleValues={getConditionalRuleValues}
            isDragging={isDragging}
            onSelectionChange={onSelectionChange}
            onCellContextMenu={onCellContextMenu}
            onAutoFilterColumn={onAutoFilterColumn}
            onClearAutoFilterColumn={onClearAutoFilterColumn}
            onSortAutoFilterColumn={onSortAutoFilterColumn}
            onStartDrag={() => setIsDragging(true)}
            onExtendDrag={(address) => {
              if (isDragging) {
                onSelectionChange({ start: selection.start, end: addressToCoord(address) });
              }
            }}
            isAutoFillDragging={Boolean(autoFillDrag)}
            onStartAutoFill={() => {
              setIsDragging(false);
              setAutoFillDrag({ source: normalizedSelection, target: normalizedSelection });
            }}
            onPreviewAutoFill={(row, column) => {
              setAutoFillDrag((current) => {
                if (!current) {
                  return current;
                }
                return { ...current, target: createAutoFillTarget(current.source, row, column) };
              });
            }}
            onStartEdit={onStartEdit}
            onEditValueChange={onEditValueChange}
            onCommitEdit={onCommitEdit}
            onCancelEdit={onCancelEdit}
            onSelectRow={(selectedRow) =>
              onSelectionChange({
                start: { row: selectedRow, column: 0 },
                end: { row: selectedRow, column: sheet.columnCount - 1 }
              })
            }
            onStartRowResize={(event) => {
              setResizeDraft({
                kind: "row",
                index: row,
                startClient: event.clientY,
                startSize: rowHeight(sheet, row, null),
                size: rowHeight(sheet, row, null)
              });
            }}
          />
        ))}
        {bottomSpacerHeight > 0 ? (
          <div
            className="grid-row-spacer"
            style={{ gridColumn: `1 / span ${gridColumnCount}`, height: bottomSpacerHeight }}
          />
        ) : null}
      </div>
    </div>
  );
}

function RowFragment({
  row,
  rowHeight,
  columns,
  columnWidths,
  sheet,
  formulaEngine,
  selection,
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
  isDragging,
  onSelectionChange,
  onCellContextMenu,
  onAutoFilterColumn,
  onClearAutoFilterColumn,
  onSortAutoFilterColumn,
  onStartDrag,
  onExtendDrag,
  isAutoFillDragging,
  onStartAutoFill,
  onPreviewAutoFill,
  onStartEdit,
  onEditValueChange,
  onCommitEdit,
  onCancelEdit,
  onSelectRow,
  onStartRowResize
}: {
  row: number;
  rowHeight: number;
  columns: number[];
  columnWidths: number[];
  sheet: SheetModel;
  formulaEngine: FormulaEngine;
  selection: CellRange;
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
  isDragging: boolean;
  onSelectionChange: (range: CellRange) => void;
  onCellContextMenu?: (event: { address: string; row: number; column: number; x: number; y: number }) => void;
  onAutoFilterColumn?: (column: number, values: string[]) => void;
  onClearAutoFilterColumn?: (column: number) => void;
  onSortAutoFilterColumn?: (column: number, direction: "asc" | "desc") => void;
  onStartDrag: () => void;
  onExtendDrag: (address: string) => void;
  isAutoFillDragging: boolean;
  onStartAutoFill: () => void;
  onPreviewAutoFill: (row: number, column: number) => void;
  onStartEdit: (address: string) => void;
  onEditValueChange: (value: string) => void;
  onCommitEdit: (address: string, value: string) => void;
  onCancelEdit: () => void;
  onSelectRow: (row: number) => void;
  onStartRowResize: (event: React.MouseEvent<HTMLButtonElement>) => void;
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
          className={["row-header", isRowSelected ? "selected-header" : ""].filter(Boolean).join(" ")}
          role="rowheader"
          aria-label={`Row ${row + 1}`}
          aria-selected={isRowSelected}
          tabIndex={0}
          style={{ height: rowHeight }}
          onClick={() => onSelectRow(row)}
          onKeyDown={(event) => {
            if (event.key === "Enter" || event.key === " ") {
              event.preventDefault();
              onSelectRow(row);
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
            onMouseDown={(event) => {
              event.preventDefault();
              event.stopPropagation();
              onStartRowResize(event);
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
        const isEditing = editingCell?.address === address && !isMergeCovered;
        const isAutoFillHandleCell =
          !isAutoFillDragging &&
          !isEditing &&
          !isMergeCovered &&
          row === selection.end.row &&
          column === selection.end.column &&
          isSelected;
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
          mergedFormat?.textColor ||
          mergedFormat?.backgroundColor ||
          mergedFormat?.horizontalAlign ||
          mergedFormat?.verticalAlign ||
          mergedFormat?.wrapText ||
          Object.keys(borderStyle).length > 0
            ? {
                fontWeight: mergedFormat?.bold ? 700 : undefined,
                fontStyle: mergedFormat?.italic ? "italic" : undefined,
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
              onSelectionChange({ start: { row, column }, end: { row, column } });
              onStartDrag();
            }}
            onMouseEnter={() => {
              if (isAutoFillDragging) {
                onPreviewAutoFill(row, column);
                return;
              }
              if (isDragging) {
                onExtendDrag(address);
              }
            }}
            onClick={() => onSelectionChange({ start: { row, column }, end: { row, column } })}
            onContextMenu={(event) => {
              event.preventDefault();
              onSelectionChange({ start: { row, column }, end: { row, column } });
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
            {isAutoFillHandleCell ? (
              <button
                type="button"
                className="auto-fill-handle"
                aria-label="AutoFill selection"
                title="AutoFill selection"
                onMouseDown={(event) => {
                  event.preventDefault();
                  event.stopPropagation();
                  onStartAutoFill();
                }}
                onClick={(event) => {
                  event.preventDefault();
                  event.stopPropagation();
                }}
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
                        onCommitEdit(address, editingCell.value);
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
                        if ((event.key === "ArrowRight" || event.key === "ArrowDown") && visibleSuggestions.length > 0) {
                          event.preventDefault();
                          setActiveSuggestionIndex((current) => wrapSuggestionIndex(current + 1, visibleSuggestions.length));
                          return;
                        }

                        if ((event.key === "ArrowLeft" || event.key === "ArrowUp") && visibleSuggestions.length > 0) {
                          event.preventDefault();
                          setActiveSuggestionIndex((current) => wrapSuggestionIndex(current - 1, visibleSuggestions.length));
                          return;
                        }

                        if ((event.key === "Tab" || event.key === "Enter") && visibleSuggestions.length > 0) {
                          event.preventDefault();
                          const suggestion = visibleSuggestions[wrapSuggestionIndex(activeSuggestionIndex, visibleSuggestions.length)];
                          if (suggestion) {
                            setDismissedEditingSuggestion(null);
                            onEditValueChange(insertFormulaSuggestion(editingCell.value, suggestion.name));
                          }
                          return;
                        }

                        if (event.key === "Enter") {
                          event.preventDefault();
                          onCommitEdit(address, editingCell.value);
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

function createAutoFillTarget(source: CellRange, row: number, column: number): CellRange {
  const normalized = normalizeRange(source);
  if (row > normalized.end.row && column >= normalized.start.column && column <= normalized.end.column) {
    return {
      start: normalized.start,
      end: { row, column: normalized.end.column }
    };
  }

  if (column > normalized.end.column && row >= normalized.start.row && row <= normalized.end.row) {
    return {
      start: normalized.start,
      end: { row: normalized.end.row, column }
    };
  }

  return normalized;
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
