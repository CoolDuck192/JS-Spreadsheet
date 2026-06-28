import { useEffect, useMemo, useRef, useState } from "react";
import { ChartPanel } from "./components/ChartPanel";
import { CellContextMenu } from "./components/CellContextMenu";
import { ConditionalFormattingPanel } from "./components/ConditionalFormattingPanel";
import { DataValidationPanel } from "./components/DataValidationPanel";
import { FilterPanel } from "./components/FilterPanel";
import { FindReplacePanel } from "./components/FindReplacePanel";
import { FormulaAuditPanel, type FormulaAuditReference } from "./components/FormulaAuditPanel";
import { FormulaBar } from "./components/FormulaBar";
import { FunctionLibraryPanel } from "./components/FunctionLibraryPanel";
import { GoToPanel } from "./components/GoToPanel";
import { Grid } from "./components/Grid";
import { NamedRangesPanel } from "./components/NamedRangesPanel";
import { PivotPanel } from "./components/PivotPanel";
import { SheetCharts } from "./components/SheetCharts";
import { SheetTabs } from "./components/SheetTabs";
import { StatusBar } from "./components/StatusBar";
import { Toolbar } from "./components/Toolbar";
import type {
  BorderPreset,
  CellContent,
  CellFormat,
  CellRange,
  DataValidationRule,
  FilterOperator,
  HistoryState,
  SheetChartType,
  SheetModel,
  NamedRange,
  PivotSheetMetadata,
  WorkbookModel
} from "./types";
import {
  columnIndexToName,
  formatCellAddress,
  getRangeAddresses,
  normalizeRange,
  parseCellAddress,
  parseRangeAddress
} from "./lib/addressing";
import { createAutoFitColumnPlan, createAutoFitRowPlan } from "./lib/autoFit";
import { createAutoSumPlan, formatRangeAddress, type AutoFunctionName } from "./lib/autoSum";
import { createChartData } from "./lib/charts";
import { parseCsv, serializeCsv } from "./lib/csv";
import { formatDisplayValue } from "./lib/displayFormat";
import { summarizeDataValidationRules, type DataValidationSummary } from "./lib/dataValidationSummary";
import { createFormulaEngine } from "./lib/formulaEngine";
import { extractFormulaReferences } from "./lib/formulaReferences";
import { getFormulaSuggestions, insertFormulaSuggestion } from "./lib/formulaSuggestions";
import { loadWorkbook, saveWorkbook } from "./lib/persistence";
import {
  createPivotTableWithDrilldowns,
  getPivotMaterializedRowKind,
  materializePivotRows,
  togglePivotDrilldown,
  type PivotConfig
} from "./lib/pivot";
import { validateCellValue } from "./lib/validation";
import { exportWorkbookToXlsx, importWorkbookFromXlsx } from "./lib/xlsx";
import {
  addSheet,
  addSheetChart,
  addConditionalFormatRule,
  autoFillRange,
  clearCellComments,
  clearDirectCellFormats,
  clearCellFormats,
  clearCellHyperlinks,
  addSheetFilter,
  clearRangeAll,
  clearRange,
  clearConditionalFormatRules,
  clearHiddenRowsAndColumns,
  clearSheetFilter,
  clearSheetFilters,
  commitHistory,
  copyRange,
  copyRichRange,
  createBlankWorkbook,
  createHistory,
  defineNamedRange,
  deleteSheetChart,
  deleteSheet,
  deleteColumns,
  deleteRows,
  duplicateSheet,
  fillDown,
  fillRight,
  getActiveSheet,
  getCellComment,
  getCellConditionalFormatRules,
  getCellContent,
  getCellFormat,
  getCellHyperlink,
  getCellReadOnly,
  getCellValidation,
  getNamedRangeForSelection,
  insertColumns,
  insertRows,
  isValidNamedRangeName,
  mergeCells,
  moveRichRange,
  moveSheet,
  pasteMatrix,
  pasteRichRange,
  previewRichPaste,
  redoHistory,
  removeConditionalFormatRule,
  removeDuplicateRows,
  removeNamedRange,
  renameSheet,
  setActiveSheet,
  setCellContent,
  setCellComment,
  setColumnWidth,
  setCellBorders,
  setCellFormat,
  setCellHyperlink,
  setCellValidation,
  setColumnsHidden,
  setRangeReadOnly,
  setSheetFreezePanes,
  setSheetHidden,
  setSheetProtection,
  setSheetTabColor,
  setRowsHidden,
  setRowHeight,
  sortRange,
  unhideAllSheets,
  unmergeCells,
  undoHistory
} from "./lib/workbook";
import type { FormulaEngine } from "./lib/formulaEngine";
import type { PasteRichRangeOptions, RichClipboardRange, RichPasteMode } from "./lib/workbook";

const INITIAL_SELECTION: CellRange = {
  start: { row: 0, column: 0 },
  end: { row: 0, column: 0 }
};
const MIN_ZOOM = 50;
const MAX_ZOOM = 200;
const ZOOM_STEP = 25;
const DEFAULT_SHEET_TAB_COLOR = "#2f7d9f";

type RichClipboardState = {
  text: string;
  range: RichClipboardRange;
  operation: "copy" | "cut";
  sourceSheetId: string;
};

type FormatPainterState = {
  format: CellFormat;
};

export default function App() {
  const [history, setHistory] = useState(() => createHistory(readInitialWorkbook()));
  const [selection, setSelection] = useState<CellRange>(INITIAL_SELECTION);
  const [editingCell, setEditingCell] = useState<{ address: string; value: string } | null>(null);
  const [formulaDraft, setFormulaDraft] = useState("");
  const [nameBoxDraft, setNameBoxDraft] = useState("");
  const [status, setStatus] = useState("Ready");
  const [isPivotPanelOpen, setPivotPanelOpen] = useState(false);
  const [isFindPanelOpen, setFindPanelOpen] = useState(false);
  const [isFilterPanelOpen, setFilterPanelOpen] = useState(false);
  const [isValidationPanelOpen, setValidationPanelOpen] = useState(false);
  const [isConditionalPanelOpen, setConditionalPanelOpen] = useState(false);
  const [isChartPanelOpen, setChartPanelOpen] = useState(false);
  const [isFunctionLibraryOpen, setFunctionLibraryOpen] = useState(false);
  const [isNamedRangesOpen, setNamedRangesOpen] = useState(false);
  const [isGoToPanelOpen, setGoToPanelOpen] = useState(false);
  const [isFormulaAuditOpen, setFormulaAuditOpen] = useState(false);
  const [goToDraft, setGoToDraft] = useState("");
  const [findDraft, setFindDraft] = useState("");
  const [replaceDraft, setReplaceDraft] = useState("");
  const [richClipboard, setRichClipboard] = useState<RichClipboardState | null>(null);
  const [formatPainter, setFormatPainter] = useState<FormatPainterState | null>(null);
  const [cellContextMenu, setCellContextMenu] = useState<{ address: string; x: number; y: number } | null>(null);
  const [zoomLevel, setZoomLevel] = useState(100);
  const [showGridlines, setShowGridlines] = useState(true);
  const [showHeaders, setShowHeaders] = useState(true);
  const [showFormulaBar, setShowFormulaBar] = useState(true);
  const [showFormulas, setShowFormulas] = useState(false);
  const [showSheetTabs, setShowSheetTabs] = useState(true);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const xlsxInputRef = useRef<HTMLInputElement>(null);
  const gridScrollRef = useRef<HTMLDivElement>(null);

  const workbook = history.present;
  const activeSheet = getActiveSheet(workbook);
  const freezeTopRow = Boolean(activeSheet.freezeTopRow);
  const freezeFirstColumn = Boolean(activeSheet.freezeFirstColumn);
  const activeAddress = formatCellAddress(selection.start);
  const selectionName = useMemo(
    () => getNamedRangeForSelection(workbook, activeSheet.id, selection)?.name ?? formatSelectionAddress(selection),
    [activeSheet.id, selection, workbook]
  );
  const formulaEngine = useMemo(() => createFormulaEngine(workbook), [workbook]);
  const activeFormat = getCellFormat(workbook, activeSheet.id, activeAddress);
  const formulaSuggestions = useMemo(() => getFormulaSuggestions(formulaDraft), [formulaDraft]);
  const activeFormulaContent = getCellContent(workbook, activeSheet.id, activeAddress);
  const activeFormula = typeof activeFormulaContent === "string" && activeFormulaContent.startsWith("=") ? activeFormulaContent : "";
  const formulaAuditPrecedents = useMemo(() => extractFormulaReferences(activeFormula), [activeFormula]);
  const formulaAuditDependents = useMemo(
    () => getFormulaDependents(activeSheet, selection),
    [activeSheet, selection]
  );
  const pivotSourceRows = useMemo(
    () => selectedRangeToDisplayRows(activeSheet, selection, formulaEngine),
    [activeSheet, formulaEngine, selection]
  );
  const pivotHeaders = useMemo(() => getPivotHeaders(pivotSourceRows), [pivotSourceRows]);
  const dataValidationRules = useMemo(
    () => summarizeDataValidationRules(activeSheet.validations ?? {}),
    [activeSheet.validations]
  );
  const selectionSummary = useMemo(
    () => summarizeSelection(activeSheet, selection, formulaEngine),
    [activeSheet, formulaEngine, selection]
  );

  useEffect(() => {
    saveWorkbook(window.localStorage, workbook);
  }, [workbook]);

  useEffect(() => {
    const raw = getCellContent(workbook, activeSheet.id, activeAddress);
    setFormulaDraft(raw === null ? "" : String(raw));
  }, [activeAddress, activeSheet.id, workbook]);

  useEffect(() => {
    setNameBoxDraft(selectionName);
  }, [selectionName]);

  useEffect(() => {
    if (!cellContextMenu) {
      return;
    }

    function closeMenu() {
      setCellContextMenu(null);
    }

    function closeMenuOnEscape(event: KeyboardEvent) {
      if (event.key === "Escape") {
        closeMenu();
      }
    }

    window.addEventListener("pointerdown", closeMenu);
    window.addEventListener("keydown", closeMenuOnEscape);
    return () => {
      window.removeEventListener("pointerdown", closeMenu);
      window.removeEventListener("keydown", closeMenuOnEscape);
    };
  }, [cellContextMenu]);

  function commitWorkbook(nextWorkbook: WorkbookModel, nextStatus = "Saved") {
    const nextHistory = commitHistory(history, nextWorkbook);
    if (nextHistory === history) {
      return;
    }
    setHistory(nextHistory);
    setStatus(nextStatus);
  }

  function applyHistoryTransition(nextHistory: HistoryState, nextStatus: string) {
    if (nextHistory === history) {
      return;
    }
    setHistory(nextHistory);
    setStatus(nextStatus);
  }

  function closeFloatingPanels() {
    setFindPanelOpen(false);
    setFilterPanelOpen(false);
    setPivotPanelOpen(false);
    setValidationPanelOpen(false);
    setConditionalPanelOpen(false);
    setChartPanelOpen(false);
    setFunctionLibraryOpen(false);
    setNamedRangesOpen(false);
    setGoToPanelOpen(false);
    setFormulaAuditOpen(false);
  }

  function commitCell(address: string, value: CellContent): boolean {
    if (!ensureEditableAddress(address)) {
      return false;
    }
    if (!validateCellCommit(address, value)) {
      return false;
    }
    commitWorkbook(setCellContent(workbook, activeSheet.id, address, value));
    return true;
  }

  function commitFormulaBar() {
    commitCell(activeAddress, formulaDraft);
  }

  function commitNameBox() {
    const nextName = nameBoxDraft.trim();
    if (!nextName || nextName === selectionName) {
      setNameBoxDraft(selectionName);
      return;
    }

    const targetRange = parseNameBoxRange(nextName, activeSheet);
    if (targetRange) {
      setSelection(targetRange);
      const targetLabel = formatSelectionAddress(targetRange);
      setNameBoxDraft(targetLabel);
      setStatus(`Selected ${targetLabel}`);
      return;
    }

    if (!isValidNamedRangeName(nextName)) {
      setNameBoxDraft(selectionName);
      setStatus("Names must start with a letter or underscore and use letters, numbers, or underscores");
      return;
    }

    const nextWorkbook = defineNamedRange(workbook, activeSheet.id, nextName, selection);
    const rangeLabel = formatSelectionAddress(selection);
    commitWorkbook(nextWorkbook, `Named ${rangeLabel} as ${nextName}`);
  }

  function selectRangeReference(targetRange: CellRange) {
    const normalizedTarget = normalizeRange(targetRange);
    const targetLabel = formatSelectionAddress(normalizedTarget);
    setSelection(normalizedTarget);
    setStatus(`Selected ${targetLabel}`);
  }

  function selectNamedRange(namedRange: NamedRange) {
    const nextWorkbook = setActiveSheet(workbook, namedRange.sheetId);
    setHistory({ ...history, present: nextWorkbook });
    setSelection(normalizeRange(namedRange.range));
    setStatus(`Selected ${namedRange.name}`);
  }

  function handleGoToReference() {
    const nextReference = goToDraft.trim();
    if (!nextReference) {
      setStatus("Enter a cell, range, or named range");
      return;
    }

    const targetRange = parseNameBoxRange(nextReference, activeSheet);
    if (targetRange) {
      selectRangeReference(targetRange);
      setGoToDraft("");
      setGoToPanelOpen(false);
      return;
    }

    const namedRange = findNamedRangeByName(workbook.namedRanges ?? [], nextReference);
    if (namedRange) {
      selectNamedRange(namedRange);
      setGoToDraft("");
      setGoToPanelOpen(false);
      return;
    }

    setStatus(`No cell, range, or named range named ${nextReference}`);
  }

  function validateCellCommit(address: string, value: CellContent): boolean {
    const rule = getCellValidation(workbook, activeSheet.id, address);
    const result = validateCellValue(value === null ? "" : String(value), rule);
    if (result.valid) {
      return true;
    }

    const coord = parseCellAddress(address);
    setSelection({ start: coord, end: coord });
    setStatus(result.message);
    return false;
  }

  function isAddressReadOnly(address: string): boolean {
    return isAddressReadOnlyOnSheet(activeSheet.id, address);
  }

  function isAddressReadOnlyOnSheet(sheetId: string, address: string): boolean {
    return getCellReadOnly(workbook, sheetId, address);
  }

  function ensureEditableAddress(address: string): boolean {
    return ensureEditableAddressOnSheet(activeSheet.id, address);
  }

  function ensureEditableAddressOnSheet(sheetId: string, address: string): boolean {
    if (!isAddressReadOnlyOnSheet(sheetId, address)) {
      return true;
    }

    if (sheetId === activeSheet.id) {
      const coord = parseCellAddress(address);
      setSelection({ start: coord, end: coord });
    }
    setStatus(`${address} is read-only`);
    return false;
  }

  function ensureEditableRange(range: CellRange): boolean {
    const readOnlyAddress = getRangeAddresses(range).find((address) => isAddressReadOnly(address));
    return readOnlyAddress ? ensureEditableAddress(readOnlyAddress) : true;
  }

  function ensureSheetStructureEditable(): boolean {
    if (!activeSheet.protection?.isProtected) {
      return true;
    }

    setStatus("Sheet is protected");
    return false;
  }

  function getVisibleCellText(address: string): string {
    return formatDisplayValue(
      formulaEngine.getDisplayValue(activeSheet.id, address),
      getCellFormat(workbook, activeSheet.id, address)
    );
  }

  function applyFormat(format: CellFormat, nextStatus = "Formatted selection") {
    if (!ensureEditableRange(selection)) {
      return;
    }
    commitWorkbook(setCellFormat(workbook, activeSheet.id, selection, format), nextStatus);
  }

  function handleFormatPainter() {
    if (formatPainter) {
      setFormatPainter(null);
      setStatus("Format painter canceled");
      return;
    }

    setEditingCell(null);
    setFormatPainter({ format: { ...activeFormat } });
    setStatus(`Copied format from ${activeAddress}`);
  }

  function handleSelectionChange(nextSelection: CellRange) {
    if (!formatPainter) {
      setSelection(nextSelection);
      return;
    }

    applyFormatPainter(nextSelection);
  }

  function applyFormatPainter(targetSelection: CellRange) {
    if (!formatPainter) {
      return;
    }

    const painter = formatPainter;
    const normalizedTarget = normalizeRange(targetSelection);
    if (!ensureEditableRange(normalizedTarget)) {
      return;
    }

    const targetLabel = formatSelectionAddress(normalizedTarget);
    const clearedWorkbook = clearDirectCellFormats(workbook, activeSheet.id, normalizedTarget);
    const paintedWorkbook = setCellFormat(clearedWorkbook, activeSheet.id, normalizedTarget, painter.format);
    setSelection(normalizedTarget);
    setFormatPainter(null);

    if (paintedWorkbook === workbook) {
      setStatus(`Painted format to ${targetLabel}`);
      return;
    }

    commitWorkbook(paintedWorkbook, `Painted format to ${targetLabel}`);
  }

  function applyBorders(preset: BorderPreset) {
    if (!ensureEditableRange(selection)) {
      return;
    }
    const nextWorkbook = setCellBorders(workbook, activeSheet.id, selection, preset);
    const selectionLabel = formatSelectionAddress(selection);
    const nextStatus =
      preset === "none"
        ? `Cleared borders from ${selectionLabel}`
        : `Applied ${borderPresetLabel(preset)} borders to ${selectionLabel}`;
    commitWorkbook(nextWorkbook, nextStatus);
  }

  function applyValidation(rule: DataValidationRule) {
    if (!ensureEditableRange(selection)) {
      return;
    }
    commitWorkbook(setCellValidation(workbook, activeSheet.id, selection, rule), "Applied data validation");
    setValidationPanelOpen(false);
  }

  function clearValidation() {
    if (!ensureEditableRange(selection)) {
      return;
    }
    commitWorkbook(setCellValidation(workbook, activeSheet.id, selection, null), "Cleared data validation");
    setValidationPanelOpen(false);
  }

  function deleteDataValidationRule(summary: DataValidationSummary) {
    if (!ensureEditableRange(summary.range)) {
      return;
    }

    commitWorkbook(
      setCellValidation(workbook, activeSheet.id, summary.range, null),
      "Deleted data validation rule"
    );
  }

  function applyConditionalFormatting(rule: Parameters<typeof addConditionalFormatRule>[3]) {
    if (!ensureEditableRange(selection)) {
      return;
    }
    commitWorkbook(
      addConditionalFormatRule(workbook, activeSheet.id, selection, rule),
      "Applied conditional formatting"
    );
    setConditionalPanelOpen(false);
  }

  function clearConditionalFormatting() {
    if (!ensureEditableRange(selection)) {
      return;
    }
    commitWorkbook(
      clearConditionalFormatRules(workbook, activeSheet.id, selection),
      "Cleared conditional formatting"
    );
    setConditionalPanelOpen(false);
  }

  function deleteConditionalFormatRule(ruleId: string) {
    const rule = (activeSheet.conditionalFormats ?? []).find((candidate) => candidate.id === ruleId);
    if (!rule || !ensureEditableRange(rule.range)) {
      return;
    }

    commitWorkbook(
      removeConditionalFormatRule(workbook, activeSheet.id, ruleId),
      "Deleted conditional format rule"
    );
  }

  function applyFilter(filter: { operator: FilterOperator; value: string }) {
    const normalized = normalizeRange(selection);
    commitWorkbook(
      addSheetFilter(workbook, activeSheet.id, normalized, {
        ...filter,
        column: normalized.start.column,
        hasHeader: normalized.start.row !== normalized.end.row
      }),
      "Filter applied"
    );
    setFilterPanelOpen(false);
  }

  function clearFilters() {
    commitWorkbook(clearSheetFilters(workbook, activeSheet.id), "Filters cleared");
    setFilterPanelOpen(false);
  }

  function toggleFilterPanel() {
    setFilterPanelOpen((isOpen) => !isOpen);
    setFindPanelOpen(false);
    setPivotPanelOpen(false);
    setValidationPanelOpen(false);
    setConditionalPanelOpen(false);
    setChartPanelOpen(false);
    setFunctionLibraryOpen(false);
    setNamedRangesOpen(false);
    setGoToPanelOpen(false);
    setFormulaAuditOpen(false);
  }

  function applyAutoFilterColumn(column: number, values: string[]) {
    const range = activeSheet.autoFilterRange ? normalizeRange(activeSheet.autoFilterRange) : null;
    if (!range) {
      return;
    }

    if (values.length === 0) {
      commitWorkbook(clearSheetFilter(workbook, activeSheet.id, range, column), `Cleared filter from ${columnIndexToName(column)}`);
      return;
    }

    commitWorkbook(
      addSheetFilter(workbook, activeSheet.id, range, {
        column,
        operator: "equals",
        value: values[0] ?? "",
        values,
        hasHeader: true
      }),
      `Filtered ${columnIndexToName(column)} by ${values.join(", ")}`
    );
  }

  function clearAutoFilterColumn(column: number) {
    const range = activeSheet.autoFilterRange ? normalizeRange(activeSheet.autoFilterRange) : null;
    if (!range) {
      return;
    }

    commitWorkbook(clearSheetFilter(workbook, activeSheet.id, range, column), `Cleared filter from ${columnIndexToName(column)}`);
  }

  function sortAutoFilterColumn(column: number, direction: "asc" | "desc") {
    const range = activeSheet.autoFilterRange ? normalizeRange(activeSheet.autoFilterRange) : null;
    if (!range || !ensureEditableRange(range)) {
      return;
    }

    commitWorkbook(
      sortRange(workbook, activeSheet.id, range, direction, column),
      direction === "asc" ? `Sorted ${columnIndexToName(column)} A to Z` : `Sorted ${columnIndexToName(column)} Z to A`
    );
  }

  function handleAutoSum(functionName: AutoFunctionName = "SUM") {
    const plan = createAutoSumPlan(
      activeSheet,
      selection,
      (address) => formulaEngine.getDisplayValue(activeSheet.id, address),
      functionName
    );
    if (!plan) {
      setStatus(`Select numbers or a blank cell beside numbers for ${autoFunctionLabel(functionName)}`);
      return;
    }

    const targetAddress = formatCellAddress(plan.target);
    if (!ensureEditableAddress(targetAddress)) {
      return;
    }
    if (!validateCellCommit(targetAddress, plan.formula)) {
      return;
    }

    commitWorkbook(
      setCellContent(workbook, activeSheet.id, targetAddress, plan.formula),
      functionName === "SUM"
        ? `Inserted AutoSum for ${formatRangeAddress(plan.source)}`
        : `Inserted ${autoFunctionLabel(functionName)} for ${formatRangeAddress(plan.source)}`
    );
    setSelection({ start: plan.target, end: plan.target });
  }

  function handleFunctionLibrary() {
    const nextOpen = !isFunctionLibraryOpen;
    closeFloatingPanels();
    setFunctionLibraryOpen(nextOpen);
    if (nextOpen) {
      setShowFormulaBar(true);
      setStatus("Function library opened");
    }
  }

  function handleInsertFunction(functionName: string) {
    setFormulaDraft(insertFormulaSuggestion("=", functionName));
    setShowFormulaBar(true);
    setFunctionLibraryOpen(false);
    setStatus(`Inserted ${functionName} function`);
  }

  function handleNamedRanges() {
    const nextOpen = !isNamedRangesOpen;
    closeFloatingPanels();
    setNamedRangesOpen(nextOpen);
    if (nextOpen) {
      setStatus("Named ranges opened");
    }
  }

  function handleFormulaAudit() {
    const nextOpen = !isFormulaAuditOpen;
    closeFloatingPanels();
    setFormulaAuditOpen(nextOpen);
    if (nextOpen) {
      setStatus("Formula audit opened");
    }
  }

  function handleSelectNamedRange(namedRange: NamedRange) {
    selectNamedRange(namedRange);
  }

  function handleDeleteNamedRange(name: string) {
    commitWorkbook(removeNamedRange(workbook, name), `Deleted named range ${name}`);
  }

  function handleComment() {
    if (!ensureEditableAddress(activeAddress)) {
      return;
    }
    const existingComment = getCellComment(workbook, activeSheet.id, activeAddress) ?? "";
    const nextComment = window.prompt("Cell comment", existingComment);
    if (nextComment === null) {
      return;
    }

    const hadComment = existingComment.trim() !== "";
    const hasComment = nextComment.trim() !== "";
    const nextWorkbook = setCellComment(workbook, activeSheet.id, activeAddress, nextComment);
    if (nextWorkbook === workbook) {
      return;
    }

    const action = hasComment ? (hadComment ? "Updated" : "Added") : "Removed";
    commitWorkbook(nextWorkbook, `${action} comment ${hasComment ? "to" : "from"} ${activeAddress}`);
  }

  function handleLink() {
    if (!ensureEditableAddress(activeAddress)) {
      return;
    }
    const existingLink = getCellHyperlink(workbook, activeSheet.id, activeAddress) ?? "";
    const promptedLink = window.prompt("Cell link URL", existingLink);
    if (promptedLink === null) {
      return;
    }

    const normalizedUrl = normalizeHyperlinkUrl(promptedLink);
    if (normalizedUrl === null) {
      setStatus("Links must use http, https, or mailto");
      return;
    }

    const hadLink = existingLink.trim() !== "";
    const hasLink = normalizedUrl !== "";
    const nextWorkbook = setCellHyperlink(workbook, activeSheet.id, activeAddress, normalizedUrl);
    if (nextWorkbook === workbook) {
      return;
    }

    const action = hasLink ? (hadLink ? "Updated" : "Added") : "Removed";
    commitWorkbook(nextWorkbook, `${action} link ${hasLink ? "to" : "from"} ${activeAddress}`);
  }

  function handleUnlink() {
    if (!ensureEditableAddress(activeAddress)) {
      return;
    }
    const existingLink = getCellHyperlink(workbook, activeSheet.id, activeAddress) ?? "";
    if (!existingLink.trim()) {
      setStatus(`No link on ${activeAddress}`);
      return;
    }

    commitWorkbook(setCellHyperlink(workbook, activeSheet.id, activeAddress, null), `Removed link from ${activeAddress}`);
  }

  function handleMergeCells() {
    const normalized = normalizeRange(selection);
    if (!ensureEditableRange(normalized)) {
      return;
    }
    const isSingleCell =
      normalized.start.row === normalized.end.row && normalized.start.column === normalized.end.column;
    if (isSingleCell) {
      setStatus("Select multiple cells to merge");
      return;
    }

    commitWorkbook(mergeCells(workbook, activeSheet.id, normalized), `Merged ${formatSelectionAddress(normalized)}`);
  }

  function handleUnmergeCells() {
    const normalized = normalizeRange(selection);
    if (!ensureEditableRange(normalized)) {
      return;
    }
    const nextWorkbook = unmergeCells(workbook, activeSheet.id, normalized);
    if (nextWorkbook === workbook) {
      setStatus("No merged cells in selection");
      return;
    }

    commitWorkbook(nextWorkbook, `Unmerged ${formatSelectionAddress(normalized)}`);
  }

  function handleCreateChart(config: { type: SheetChartType; title: string }) {
    const sourceRange = inferSourceRange(activeSheet, selection);
    const sourceRows = selectedRangeToDisplayRows(activeSheet, sourceRange, formulaEngine);
    const chartData = createChartData(sourceRows, "Chart");
    if (chartData.data.length === 0) {
      setStatus("Select labels and numeric values for a chart");
      return;
    }

    const normalized = normalizeRange(sourceRange);
    const nextWorkbook = addSheetChart(workbook, activeSheet.id, normalized, {
      anchor: {
        row: normalized.start.row,
        column: Math.min(activeSheet.columnCount - 1, normalized.end.column + 1)
      },
      title: config.title || chartData.title,
      type: config.type
    });
    commitWorkbook(nextWorkbook, "Created chart");
    setChartPanelOpen(false);
  }

  function handleDeleteChart(chartId: string) {
    if (!ensureSheetStructureEditable()) {
      return;
    }
    commitWorkbook(deleteSheetChart(workbook, activeSheet.id, chartId), "Deleted chart");
  }

  function handleColumnResize(column: number, width: number) {
    if (!ensureSheetStructureEditable()) {
      return;
    }
    commitWorkbook(
      setColumnWidth(workbook, activeSheet.id, column, width),
      `Set column ${columnIndexToName(column)} width`
    );
  }

  function handleRowResize(row: number, height: number) {
    if (!ensureSheetStructureEditable()) {
      return;
    }
    commitWorkbook(setRowHeight(workbook, activeSheet.id, row, height), `Set row ${row + 1} height`);
  }

  function handleAutoFitColumns() {
    if (!ensureSheetStructureEditable()) {
      return;
    }

    const plan = createAutoFitColumnPlan(activeSheet, selection, (address) => getVisibleCellText(address));
    if (plan.length === 0) {
      setStatus("No columns to auto-fit");
      return;
    }

    let nextWorkbook = workbook;
    for (const item of plan) {
      nextWorkbook = setColumnWidth(nextWorkbook, activeSheet.id, item.column, item.width);
    }

    if (nextWorkbook === workbook) {
      setStatus(`${pluralize(plan.length, "column")} already fit`);
      return;
    }

    commitWorkbook(nextWorkbook, `Auto-fit ${pluralize(plan.length, "column")}`);
  }

  function handleAutoFitRows() {
    if (!ensureSheetStructureEditable()) {
      return;
    }

    const plan = createAutoFitRowPlan(activeSheet, selection, (address) => getVisibleCellText(address));
    if (plan.length === 0) {
      setStatus("No rows to auto-fit");
      return;
    }

    let nextWorkbook = workbook;
    for (const item of plan) {
      nextWorkbook = setRowHeight(nextWorkbook, activeSheet.id, item.row, item.height);
    }

    if (nextWorkbook === workbook) {
      setStatus(`${pluralize(plan.length, "row")} already fit`);
      return;
    }

    commitWorkbook(nextWorkbook, `Auto-fit ${pluralize(plan.length, "row")}`);
  }

  function handleSort(direction: "asc" | "desc") {
    if (!ensureEditableRange(selection)) {
      return;
    }
    commitWorkbook(
      sortRange(workbook, activeSheet.id, selection, direction),
      direction === "asc" ? "Sorted A to Z" : "Sorted Z to A"
    );
  }

  function handleRemoveDuplicates() {
    if (!ensureEditableRange(selection)) {
      return;
    }

    const result = removeDuplicateRows(workbook, activeSheet.id, selection);
    if (result.removedCount === 0) {
      setStatus("No duplicate rows found");
      return;
    }

    commitWorkbook(result.workbook, `Removed ${result.removedCount} ${result.removedCount === 1 ? "duplicate row" : "duplicate rows"}`);
  }

  function handleInsertRows() {
    if (!ensureSheetStructureEditable()) {
      return;
    }
    const normalized = normalizeRange(selection);
    const count = normalized.end.row - normalized.start.row + 1;
    commitWorkbook(insertRows(workbook, activeSheet.id, normalized.start.row, count), `Inserted ${pluralize(count, "row")}`);
    setSelection({
      start: normalized.start,
      end: { row: normalized.start.row + count - 1, column: normalized.end.column }
    });
  }

  function handleDeleteRows() {
    if (!ensureSheetStructureEditable()) {
      return;
    }
    const normalized = normalizeRange(selection);
    const count = normalized.end.row - normalized.start.row + 1;
    commitWorkbook(deleteRows(workbook, activeSheet.id, normalized.start.row, count), `Deleted ${pluralize(count, "row")}`);
    const nextRow = clamp(normalized.start.row, 0, Math.max(activeSheet.rowCount - count - 1, 0));
    setSelection({
      start: { row: nextRow, column: normalized.start.column },
      end: { row: nextRow, column: normalized.end.column }
    });
  }

  function handleInsertColumns() {
    if (!ensureSheetStructureEditable()) {
      return;
    }
    const normalized = normalizeRange(selection);
    const count = normalized.end.column - normalized.start.column + 1;
    commitWorkbook(insertColumns(workbook, activeSheet.id, normalized.start.column, count), `Inserted ${pluralize(count, "column")}`);
    setSelection({
      start: normalized.start,
      end: { row: normalized.end.row, column: normalized.start.column + count - 1 }
    });
  }

  function handleDeleteColumns() {
    if (!ensureSheetStructureEditable()) {
      return;
    }
    const normalized = normalizeRange(selection);
    const count = normalized.end.column - normalized.start.column + 1;
    commitWorkbook(deleteColumns(workbook, activeSheet.id, normalized.start.column, count), `Deleted ${pluralize(count, "column")}`);
    const nextColumn = clamp(normalized.start.column, 0, Math.max(activeSheet.columnCount - count - 1, 0));
    setSelection({
      start: { row: normalized.start.row, column: nextColumn },
      end: { row: normalized.end.row, column: nextColumn }
    });
  }

  function handleHideRows() {
    if (!ensureSheetStructureEditable()) {
      return;
    }
    const normalized = normalizeRange(selection);
    if (!hasVisibleIndexAfterHiding(activeSheet.rowCount, activeSheet.hiddenRows ?? {}, normalized.start.row, normalized.end.row)) {
      setStatus("Cannot hide all rows");
      return;
    }

    const count = normalized.end.row - normalized.start.row + 1;
    const nextWorkbook = setRowsHidden(workbook, activeSheet.id, normalized.start.row, normalized.end.row, true);
    commitWorkbook(nextWorkbook, `Hid ${pluralize(count, "row")}`);
    const nextSheet = getActiveSheet(nextWorkbook);
    const nextRow = findNearestVisibleIndex(nextSheet.rowCount, nextSheet.hiddenRows ?? {}, normalized.end.row + 1);
    const nextColumn = findNearestVisibleIndex(nextSheet.columnCount, nextSheet.hiddenColumns ?? {}, normalized.start.column);
    setSelection({ start: { row: nextRow, column: nextColumn }, end: { row: nextRow, column: nextColumn } });
  }

  function handleHideColumns() {
    if (!ensureSheetStructureEditable()) {
      return;
    }
    const normalized = normalizeRange(selection);
    if (!hasVisibleIndexAfterHiding(activeSheet.columnCount, activeSheet.hiddenColumns ?? {}, normalized.start.column, normalized.end.column)) {
      setStatus("Cannot hide all columns");
      return;
    }

    const count = normalized.end.column - normalized.start.column + 1;
    const nextWorkbook = setColumnsHidden(workbook, activeSheet.id, normalized.start.column, normalized.end.column, true);
    commitWorkbook(nextWorkbook, `Hid ${pluralize(count, "column")}`);
    const nextSheet = getActiveSheet(nextWorkbook);
    const nextRow = findNearestVisibleIndex(nextSheet.rowCount, nextSheet.hiddenRows ?? {}, normalized.start.row);
    const nextColumn = findNearestVisibleIndex(nextSheet.columnCount, nextSheet.hiddenColumns ?? {}, normalized.end.column + 1);
    setSelection({ start: { row: nextRow, column: nextColumn }, end: { row: nextRow, column: nextColumn } });
  }

  function handleUnhideAll() {
    if (!ensureSheetStructureEditable()) {
      return;
    }

    const nextWorkbook = clearHiddenRowsAndColumns(workbook, activeSheet.id);
    if (nextWorkbook === workbook) {
      setStatus("No hidden rows or columns");
      return;
    }

    commitWorkbook(nextWorkbook, "Unhid rows and columns");
  }

  function handleClearSelection() {
    if (!ensureEditableRange(selection)) {
      return;
    }
    commitWorkbook(clearRange(workbook, activeSheet.id, selection), "Cleared selection");
  }

  function handleClearAll() {
    if (!ensureEditableRange(selection)) {
      return;
    }
    commitWorkbook(clearRangeAll(workbook, activeSheet.id, selection), "Cleared all");
  }

  function handleClearComments() {
    if (!ensureEditableRange(selection)) {
      return;
    }
    commitWorkbook(clearCellComments(workbook, activeSheet.id, selection), "Cleared comments");
  }

  function handleClearFormats() {
    if (!ensureEditableRange(selection)) {
      return;
    }
    commitWorkbook(clearCellFormats(workbook, activeSheet.id, selection), "Cleared formats");
  }

  function handleClearHyperlinks() {
    if (!ensureEditableRange(selection)) {
      return;
    }
    commitWorkbook(clearCellHyperlinks(workbook, activeSheet.id, selection), "Cleared hyperlinks");
  }

  function openCellContextMenu(event: { address: string; row: number; column: number; x: number; y: number }) {
    setSelection({ start: { row: event.row, column: event.column }, end: { row: event.row, column: event.column } });
    setEditingCell(null);
    setCellContextMenu({ address: event.address, x: event.x, y: event.y });
  }

  function handleFillDown() {
    if (!ensureEditableRange(selection)) {
      return;
    }
    commitWorkbook(fillDown(workbook, activeSheet.id, selection), "Filled down");
  }

  function handleFillRight() {
    if (!ensureEditableRange(selection)) {
      return;
    }
    commitWorkbook(fillRight(workbook, activeSheet.id, selection), "Filled right");
  }

  function handleAutoFill(sourceRange: CellRange, targetRange: CellRange) {
    const writeRange = getAutoFillWriteRange(sourceRange, targetRange);
    if (!writeRange) {
      setStatus("Drag AutoFill down or right");
      return;
    }
    if (!ensureEditableRange(writeRange)) {
      return;
    }

    const normalizedTarget = normalizeRange(targetRange);
    const targetLabel = formatSelectionAddress(normalizedTarget);
    const nextWorkbook = autoFillRange(workbook, activeSheet.id, sourceRange, normalizedTarget);
    setSelection(normalizedTarget);
    commitWorkbook(nextWorkbook, `AutoFilled ${targetLabel}`);
    if (nextWorkbook === workbook) {
      setStatus(`AutoFilled ${targetLabel}`);
    }
  }

  function handleLockCells() {
    const nextWorkbook = setRangeReadOnly(workbook, activeSheet.id, selection, true);
    commitWorkbook(nextWorkbook, `Locked ${formatSelectionAddress(selection)}`);
  }

  function handleUnlockCells() {
    const nextWorkbook = setRangeReadOnly(workbook, activeSheet.id, selection, false);
    commitWorkbook(nextWorkbook, `Unlocked ${formatSelectionAddress(selection)}`);
  }

  function handleToggleProtection() {
    const isProtected = Boolean(activeSheet.protection?.isProtected);
    commitWorkbook(
      setSheetProtection(workbook, activeSheet.id, !isProtected),
      isProtected ? "Unprotected sheet" : "Protected sheet"
    );
  }

  function handleCreatePivotTable(config: PivotConfig) {
    try {
      const sourceRows = selectedRangeToDisplayRows(activeSheet, selection, formulaEngine);
      const pivot = createPivotTableWithDrilldowns(sourceRows, config);
      const pivotName = nextPivotSheetName(workbook);
      let nextWorkbook = addSheet(workbook, pivotName);
      const pivotSheetId = nextWorkbook.activeSheetId;

      nextWorkbook = replaceGeneratedPivotSheetRows(nextWorkbook, pivotSheetId, pivot.metadata);

      commitWorkbook(nextWorkbook, `Created ${pivotName}`);
      setSelection(INITIAL_SELECTION);
      setPivotPanelOpen(false);
    } catch (error) {
      setStatus(error instanceof Error ? error.message : "Could not create pivot table");
    }
  }

  function handleTogglePivotDrilldown(entryId: string) {
    const pivot = activeSheet.pivot;
    if (!pivot || !pivot.drilldowns[entryId]) {
      return;
    }

    const wasExpanded = Boolean(pivot.expanded[entryId]);
    const nextPivot = togglePivotDrilldown(pivot, entryId);
    const nextWorkbook = replaceGeneratedPivotSheetRows(workbook, activeSheet.id, nextPivot);
    commitWorkbook(nextWorkbook, wasExpanded ? "Collapsed pivot drilldown" : "Expanded pivot drilldown");
  }

  function validateRichPaste(mode: RichPasteMode, options: PasteRichRangeOptions = {}): boolean {
    if (!richClipboard) {
      setStatus("Copy cells before using paste special");
      return false;
    }

    const preview = previewRichPaste(richClipboard.range, activeAddress, { ...options, mode });
    for (const item of preview) {
      if (!ensureEditableAddress(item.address)) {
        return false;
      }
      if (mode === "formats") {
        continue;
      }
      const rule = mode === "values" ? getCellValidation(workbook, activeSheet.id, item.address) : item.validation;
      const result = validateCellValue(item.content === null ? "" : String(item.content), rule);
      if (!result.valid) {
        const coord = parseCellAddress(item.address);
        setSelection({ start: coord, end: coord });
        setStatus(result.message);
        return false;
      }
    }

    return true;
  }

  function ensureMoveSourceEditable(clipboard: RichClipboardState, targetAddress: string): boolean {
    const preservedSourceAddresses =
      clipboard.sourceSheetId === activeSheet.id
        ? new Set(previewRichPaste(clipboard.range, targetAddress, { translateFormulas: false }).map((item) => item.address))
        : new Set<string>();

    const readOnlyAddress = getRangeAddresses(clipboard.range.range).find(
      (address) => !preservedSourceAddresses.has(address) && isAddressReadOnlyOnSheet(clipboard.sourceSheetId, address)
    );
    return readOnlyAddress ? ensureEditableAddressOnSheet(clipboard.sourceSheetId, readOnlyAddress) : true;
  }

  function pasteRichClipboardMode(mode: RichPasteMode, nextStatus: string): boolean {
    const isMovePaste = mode === "all" && richClipboard?.operation === "cut";
    const options: PasteRichRangeOptions = isMovePaste ? { mode, translateFormulas: false } : { mode };

    if (!validateRichPaste(mode, options)) {
      return false;
    }
    if (isMovePaste && !ensureMoveSourceEditable(richClipboard!, activeAddress)) {
      return false;
    }

    if (isMovePaste) {
      commitWorkbook(
        moveRichRange(workbook, richClipboard!.sourceSheetId, activeSheet.id, activeAddress, richClipboard!.range),
        "Moved selection"
      );
      setRichClipboard(null);
      return true;
    }

    commitWorkbook(pasteRichRange(workbook, activeSheet.id, activeAddress, richClipboard!.range, { mode }), nextStatus);
    return true;
  }

  function pasteRichClipboard(text: string): boolean {
    if (!richClipboard || text !== richClipboard.text) {
      return false;
    }

    pasteRichClipboardMode("all", "Pasted cells with formatting");
    return true;
  }

  function handlePasteAll() {
    pasteRichClipboardMode("all", "Pasted cells with formatting");
  }

  function handlePasteValues() {
    pasteRichClipboardMode("values", "Pasted values");
  }

  function handlePasteFormats() {
    pasteRichClipboardMode("formats", "Pasted formats");
  }

  function handleTransposePaste() {
    pasteRichClipboardMode("transpose", "Transposed paste");
  }

  function pasteText(text: string) {
    const normalizedText = normalizeClipboardText(text);
    if (pasteRichClipboard(normalizedText)) {
      return;
    }
    setRichClipboard(null);

    const matrix = normalizedText
      .split("\n")
      .filter((line, index, lines) => line.length > 0 || index < lines.length - 1)
      .map((line) => line.split("\t"));

    if (matrix.length === 0) {
      return;
    }

    const start = parseCellAddress(activeAddress);
    for (let rowOffset = 0; rowOffset < matrix.length; rowOffset += 1) {
      for (let columnOffset = 0; columnOffset < matrix[rowOffset].length; columnOffset += 1) {
        const address = formatCellAddress({ row: start.row + rowOffset, column: start.column + columnOffset });
        if (!ensureEditableAddress(address)) {
          return;
        }
        if (!validateCellCommit(address, matrix[rowOffset][columnOffset])) {
          return;
        }
      }
    }

    commitWorkbook(pasteMatrix(workbook, activeSheet.id, activeAddress, matrix), "Pasted cells");
  }

  function createRichClipboardPayload(operation: RichClipboardState["operation"]): RichClipboardState {
    const matrix = copyRange(workbook, activeSheet.id, selection);
    const text = normalizeClipboardText(matrix.map((row) => row.map((cell) => (cell === null ? "" : String(cell))).join("\t")).join("\n"));
    return {
      text,
      range: copyRichRange(workbook, activeSheet.id, selection, { getDisplayValue: getVisibleCellText }),
      operation,
      sourceSheetId: activeSheet.id
    };
  }

  function copySelection() {
    const nextClipboard = createRichClipboardPayload("copy");
    setRichClipboard(nextClipboard);
    void navigator.clipboard?.writeText(nextClipboard.text);
    setStatus("Copied selection");
  }

  function cutSelection() {
    if (!ensureEditableRange(selection)) {
      return;
    }

    const nextClipboard = createRichClipboardPayload("cut");
    setRichClipboard(nextClipboard);
    void navigator.clipboard?.writeText(nextClipboard.text);
    setStatus("Cut selection");
  }

  function handleShellKeyCommand(event: React.KeyboardEvent<HTMLElement>) {
    if (event.defaultPrevented || isEditableEventTarget(event.target)) {
      return;
    }
    handleKeyCommand(event);
  }

  function handleKeyCommand(event: React.KeyboardEvent<HTMLElement>) {
    if (editingCell) {
      return;
    }

    const isCommand = event.metaKey || event.ctrlKey;
    if (event.key === "Escape" && formatPainter) {
      event.preventDefault();
      setFormatPainter(null);
      setStatus("Format painter canceled");
      return;
    }

    if (isCommand && event.key.toLowerCase() === "c") {
      event.preventDefault();
      copySelection();
      return;
    }

    if (isCommand && event.key.toLowerCase() === "x") {
      event.preventDefault();
      cutSelection();
      return;
    }

    if (isCommand && event.key.toLowerCase() === "v") {
      return;
    }

    if (isCommand && !event.shiftKey && event.key.toLowerCase() === "b") {
      event.preventDefault();
      applyFormat({ bold: !activeFormat.bold }, "Applied bold");
      return;
    }

    if (isCommand && !event.shiftKey && event.key.toLowerCase() === "i") {
      event.preventDefault();
      applyFormat({ italic: !activeFormat.italic }, "Applied italic");
      return;
    }

    if (isCommand && !event.shiftKey && event.key.toLowerCase() === "k") {
      event.preventDefault();
      handleLink();
      return;
    }

    if (isCommand && event.shiftKey && event.key.toLowerCase() === "l") {
      event.preventDefault();
      toggleFilterPanel();
      return;
    }

    if (isCommand && event.key.toLowerCase() === "f") {
      event.preventDefault();
      setPivotPanelOpen(false);
      setFilterPanelOpen(false);
      setValidationPanelOpen(false);
      setConditionalPanelOpen(false);
      setChartPanelOpen(false);
      setGoToPanelOpen(false);
      setFormulaAuditOpen(false);
      setFindPanelOpen(true);
      return;
    }

    if (isCommand && event.key.toLowerCase() === "z") {
      event.preventDefault();
      applyHistoryTransition(
        event.shiftKey ? redoHistory(history) : undoHistory(history),
        event.shiftKey ? "Redone" : "Undone"
      );
      return;
    }

    if (isCommand && event.key.toLowerCase() === "y") {
      event.preventDefault();
      applyHistoryTransition(redoHistory(history), "Redone");
      return;
    }

    if (isCommand && event.key.toLowerCase() === "d") {
      event.preventDefault();
      handleFillDown();
      return;
    }

    if (isCommand && event.key.toLowerCase() === "r") {
      event.preventDefault();
      handleFillRight();
      return;
    }

    if (event.key === "Delete" || event.key === "Backspace") {
      event.preventDefault();
      handleClearSelection();
      return;
    }

    if (event.key === "Enter") {
      event.preventDefault();
      if (!ensureEditableAddress(activeAddress)) {
        return;
      }
      setEditingCell({ address: activeAddress, value: formulaDraft });
      return;
    }

    if (event.key === "ArrowDown" || event.key === "ArrowUp" || event.key === "ArrowLeft" || event.key === "ArrowRight") {
      event.preventDefault();
      moveSelection(event.key);
      return;
    }

    if (event.key.length === 1 && !event.altKey && !isCommand) {
      event.preventDefault();
      if (!ensureEditableAddress(activeAddress)) {
        return;
      }
      setEditingCell({ address: activeAddress, value: event.key });
    }
  }

  function moveSelection(key: string) {
    const delta = {
      ArrowDown: { row: 1, column: 0 },
      ArrowUp: { row: -1, column: 0 },
      ArrowLeft: { row: 0, column: -1 },
      ArrowRight: { row: 0, column: 1 }
    }[key];

    if (!delta) {
      return;
    }

    const next = {
      row: clamp(selection.start.row + delta.row, 0, activeSheet.rowCount - 1),
      column: clamp(selection.start.column + delta.column, 0, activeSheet.columnCount - 1)
    };
    setSelection({ start: next, end: next });
  }

  function handleAddSheet() {
    commitWorkbook(addSheet(workbook), "Added sheet");
    setSelection(INITIAL_SELECTION);
  }

  function handleRenameSheet() {
    const nextName = window.prompt("Rename active sheet", activeSheet.name);
    if (nextName) {
      commitWorkbook(renameSheet(workbook, activeSheet.id, nextName), "Renamed sheet");
    }
  }

  function handleDuplicateSheet() {
    commitWorkbook(duplicateSheet(workbook, activeSheet.id), "Duplicated sheet");
    setSelection(INITIAL_SELECTION);
  }

  function handleDeleteSheet() {
    commitWorkbook(deleteSheet(workbook, activeSheet.id), "Deleted sheet");
    setSelection(INITIAL_SELECTION);
  }

  function handleHideSheet() {
    const hiddenSheetName = activeSheet.name;
    const nextWorkbook = setSheetHidden(workbook, activeSheet.id, true);
    if (nextWorkbook === workbook) {
      setStatus("Cannot hide the only visible sheet");
      return;
    }

    commitWorkbook(nextWorkbook, `Hid ${hiddenSheetName}`);
    setSelection(INITIAL_SELECTION);
  }

  function handleUnhideSheets() {
    const nextWorkbook = unhideAllSheets(workbook);
    if (nextWorkbook === workbook) {
      setStatus("No hidden sheets");
      return;
    }

    commitWorkbook(nextWorkbook, "Unhid sheets");
  }

  function handleMoveSheet(direction: "left" | "right") {
    const visibleSheets = workbook.sheets.filter((sheet) => sheet.isHidden !== true);
    const visibleIndex = visibleSheets.findIndex((sheet) => sheet.id === activeSheet.id);
    const nextVisibleIndex = direction === "left" ? visibleIndex - 1 : visibleIndex + 1;
    const boundary = direction === "left" ? "first" : "last";

    if (visibleIndex === -1 || nextVisibleIndex < 0 || nextVisibleIndex >= visibleSheets.length) {
      setStatus(`${activeSheet.name} is already ${boundary}`);
      return;
    }

    const targetSheet = visibleSheets[nextVisibleIndex];
    const targetIndex = workbook.sheets.findIndex((sheet) => sheet.id === targetSheet.id);
    const nextWorkbook = moveSheet(workbook, activeSheet.id, targetIndex);
    if (nextWorkbook === workbook) {
      setStatus(`${activeSheet.name} is already ${boundary}`);
      return;
    }

    commitWorkbook(nextWorkbook, `Moved ${activeSheet.name} ${direction}`);
  }

  function handleSheetTabColor(color: string) {
    commitWorkbook(setSheetTabColor(workbook, activeSheet.id, color), `Changed ${activeSheet.name} tab color`);
  }

  function handleNewWorkbook() {
    const next = createBlankWorkbook();
    setHistory(createHistory(next));
    setSelection(INITIAL_SELECTION);
    setFormatPainter(null);
    setStatus("New workbook");
  }

  function handleImportCsv(file: File | undefined) {
    if (!file) {
      return;
    }

    const reader = new FileReader();
    reader.onload = () => {
      try {
        const rows = parseCsv(String(reader.result ?? ""));
        setHistory((current) => commitHistory(current, replaceActiveSheetWithRows(current.present, rows)));
        setSelection(INITIAL_SELECTION);
        setFormatPainter(null);
        setStatus(`Imported ${file.name}`);
      } catch (error) {
        setStatus(error instanceof Error ? error.message : "CSV import failed");
      }
    };
    reader.readAsText(file);
  }

  function handleExportCsv() {
    const csv = serializeCsv(activeSheetToRows(activeSheet));
    const blob = new Blob([csv], { type: "text/csv;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = `${activeSheet.name}.csv`;
    anchor.click();
    URL.revokeObjectURL(url);
    setStatus("Exported CSV");
  }

  function handleImportXlsx(file: File | undefined) {
    if (!file) {
      return;
    }

    file
      .arrayBuffer()
      .then(async (buffer) => {
        const nextWorkbook = await importWorkbookFromXlsx(buffer);
        setHistory((current) => commitHistory(current, nextWorkbook));
        setSelection(INITIAL_SELECTION);
        setFormatPainter(null);
        setStatus(`Imported ${file.name}`);
      })
      .catch((error: unknown) => {
        setStatus(error instanceof Error ? error.message : "XLSX import failed");
      });
  }

  function handleExportXlsx() {
    exportWorkbookToXlsx(workbook)
      .then((bytes) => {
        const blob = new Blob([bytes], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
        const url = URL.createObjectURL(blob);
        const anchor = document.createElement("a");
        anchor.href = url;
        anchor.download = "spreadsheet.xlsx";
        anchor.click();
        URL.revokeObjectURL(url);
        setStatus("Exported XLSX");
      })
      .catch((error: unknown) => {
        setStatus(error instanceof Error ? error.message : "XLSX export failed");
      });
  }

  function handlePrintWorkbook() {
    setStatus("Opened print dialog");
    window.print();
  }

  function handleResetView() {
    const grid = gridScrollRef.current;
    if (grid) {
      grid.scrollLeft = 0;
      grid.scrollTop = 0;
      setStatus("View reset");
    }
  }

  function handleToggleGridlines() {
    const nextShowGridlines = !showGridlines;
    setShowGridlines(nextShowGridlines);
    setStatus(nextShowGridlines ? "Gridlines shown" : "Gridlines hidden");
  }

  function handleToggleHeaders() {
    const nextShowHeaders = !showHeaders;
    setShowHeaders(nextShowHeaders);
    setStatus(nextShowHeaders ? "Headers shown" : "Headers hidden");
  }

  function handleToggleFormulaBar() {
    const nextShowFormulaBar = !showFormulaBar;
    setShowFormulaBar(nextShowFormulaBar);
    setStatus(nextShowFormulaBar ? "Formula bar shown" : "Formula bar hidden");
  }

  function handleToggleShowFormulas() {
    const nextShowFormulas = !showFormulas;
    setShowFormulas(nextShowFormulas);
    setStatus(nextShowFormulas ? "Formulas shown" : "Formula results shown");
  }

  function handleToggleSheetTabs() {
    const nextShowSheetTabs = !showSheetTabs;
    setShowSheetTabs(nextShowSheetTabs);
    setStatus(nextShowSheetTabs ? "Sheet tabs shown" : "Sheet tabs hidden");
  }

  function handleZoom(delta: number) {
    const nextZoomLevel = clamp(zoomLevel + delta, MIN_ZOOM, MAX_ZOOM);
    setZoomLevel(nextZoomLevel);
    setStatus(`Zoom ${nextZoomLevel}%`);
  }

  function handleResetZoom() {
    setZoomLevel(100);
    setStatus("Zoom reset");
  }

  function handleFindNext() {
    const matches = findMatches(activeSheet, findDraft);
    if (findDraft.trim() === "") {
      setStatus("Enter text to find");
      return;
    }
    if (matches.length === 0) {
      setStatus(`No matches for ${findDraft}`);
      return;
    }

    const current = selection.start;
    const nextMatch =
      matches.find((match) => match.row > current.row || (match.row === current.row && match.column > current.column)) ??
      matches[0];
    setSelection({ start: { row: nextMatch.row, column: nextMatch.column }, end: { row: nextMatch.row, column: nextMatch.column } });
    setStatus(`Found ${nextMatch.address}`);
  }

  function handleReplaceCurrent() {
    if (findDraft.trim() === "") {
      setStatus("Enter text to find");
      return;
    }

    const address = activeAddress;
    const currentContent = getCellContent(workbook, activeSheet.id, address);
    const nextContent =
      currentContent === null ? null : replaceFirstMatch(String(currentContent), findDraft, replaceDraft);
    if (nextContent === currentContent || nextContent === null) {
      handleFindNext();
      return;
    }

    if (commitCell(address, nextContent)) {
      setStatus(`Replaced ${address}`);
    }
  }

  function handleReplaceAll() {
    if (findDraft.trim() === "") {
      setStatus("Enter text to find");
      return;
    }

    const matches = findMatches(activeSheet, findDraft);
    const readOnlyMatch = matches.find(
      (match) => replaceEveryMatch(String(match.content), findDraft, replaceDraft) !== match.content && isAddressReadOnly(match.address)
    );
    if (readOnlyMatch) {
      ensureEditableAddress(readOnlyMatch.address);
      return;
    }

    let nextWorkbook = workbook;
    let changedCells = 0;
    for (const match of matches) {
      const nextContent = replaceEveryMatch(String(match.content), findDraft, replaceDraft);
      if (nextContent !== match.content) {
        nextWorkbook = setCellContent(nextWorkbook, activeSheet.id, match.address, nextContent);
        changedCells += 1;
      }
    }

    if (changedCells === 0) {
      setStatus(`No matches for ${findDraft}`);
      return;
    }

    commitWorkbook(nextWorkbook, `Replaced ${changedCells} ${changedCells === 1 ? "cell" : "cells"}`);
  }

  return (
    <main className="app-shell" onKeyDown={handleShellKeyCommand}>
      <section className="spreadsheet-surface" aria-label="JavaScript spreadsheet">
        <Toolbar
          canUndo={history.past.length > 0}
          canRedo={history.future.length > 0}
          onNew={handleNewWorkbook}
          onImport={() => fileInputRef.current?.click()}
          onExport={handleExportCsv}
          onImportXlsx={() => xlsxInputRef.current?.click()}
          onExportXlsx={handleExportXlsx}
          onPrint={handlePrintWorkbook}
          onUndo={() => {
            applyHistoryTransition(undoHistory(history), "Undone");
          }}
          onRedo={() => {
            applyHistoryTransition(redoHistory(history), "Redone");
          }}
          onClear={handleClearSelection}
          canPasteSpecial={Boolean(richClipboard)}
          onPasteValues={handlePasteValues}
          onPasteFormats={handlePasteFormats}
          formatPainterActive={Boolean(formatPainter)}
          onFormatPainter={handleFormatPainter}
          onTransposePaste={handleTransposePaste}
          onAutoSum={handleAutoSum}
          onFunctionLibrary={handleFunctionLibrary}
          formulaAuditOpen={isFormulaAuditOpen}
          onFormulaAudit={handleFormulaAudit}
          onNamedRanges={handleNamedRanges}
          onGoTo={() => {
            const nextOpen = !isGoToPanelOpen;
            closeFloatingPanels();
            setGoToPanelOpen(nextOpen);
            if (nextOpen) {
              setGoToDraft("");
              setStatus("Go To opened");
            }
          }}
          onComment={handleComment}
          onLink={handleLink}
          onUnlink={handleUnlink}
          onMergeCells={handleMergeCells}
          onUnmergeCells={handleUnmergeCells}
          onFillDown={handleFillDown}
          onFillRight={handleFillRight}
          sheetProtected={Boolean(activeSheet.protection?.isProtected)}
          onLockCells={handleLockCells}
          onUnlockCells={handleUnlockCells}
          onToggleProtectSheet={handleToggleProtection}
          onSortAsc={() => handleSort("asc")}
          onSortDesc={() => handleSort("desc")}
          onRemoveDuplicates={handleRemoveDuplicates}
          onInsertRows={handleInsertRows}
          onDeleteRows={handleDeleteRows}
          onInsertColumns={handleInsertColumns}
          onDeleteColumns={handleDeleteColumns}
          onAutoFitRows={handleAutoFitRows}
          onAutoFitColumns={handleAutoFitColumns}
          onHideRows={handleHideRows}
          onHideColumns={handleHideColumns}
          onUnhideAll={handleUnhideAll}
          freezeTopRow={freezeTopRow}
          freezeFirstColumn={freezeFirstColumn}
          onToggleFreezeTopRow={() => {
            const nextFreezeTopRow = !freezeTopRow;
            commitWorkbook(
              setSheetFreezePanes(workbook, activeSheet.id, {
                freezeTopRow: nextFreezeTopRow,
                freezeFirstColumn
              }),
              nextFreezeTopRow ? "Froze top row" : "Unfroze top row"
            );
          }}
          onToggleFreezeFirstColumn={() => {
            const nextFreezeFirstColumn = !freezeFirstColumn;
            commitWorkbook(
              setSheetFreezePanes(workbook, activeSheet.id, {
                freezeTopRow,
                freezeFirstColumn: nextFreezeFirstColumn
              }),
              nextFreezeFirstColumn ? "Froze first column" : "Unfroze first column"
            );
          }}
          onFindReplace={() => {
            setFindPanelOpen((isOpen) => !isOpen);
            setPivotPanelOpen(false);
            setFilterPanelOpen(false);
            setValidationPanelOpen(false);
            setConditionalPanelOpen(false);
            setChartPanelOpen(false);
            setFunctionLibraryOpen(false);
            setNamedRangesOpen(false);
            setGoToPanelOpen(false);
            setFormulaAuditOpen(false);
          }}
          onFilter={() => {
            toggleFilterPanel();
          }}
          onDataValidation={() => {
            setValidationPanelOpen((isOpen) => !isOpen);
            setFindPanelOpen(false);
            setPivotPanelOpen(false);
            setFilterPanelOpen(false);
            setConditionalPanelOpen(false);
            setChartPanelOpen(false);
            setFunctionLibraryOpen(false);
            setNamedRangesOpen(false);
            setGoToPanelOpen(false);
            setFormulaAuditOpen(false);
          }}
          onConditionalFormatting={() => {
            setConditionalPanelOpen((isOpen) => !isOpen);
            setFindPanelOpen(false);
            setPivotPanelOpen(false);
            setFilterPanelOpen(false);
            setValidationPanelOpen(false);
            setChartPanelOpen(false);
            setFunctionLibraryOpen(false);
            setNamedRangesOpen(false);
            setGoToPanelOpen(false);
            setFormulaAuditOpen(false);
          }}
          onAddSheet={handleAddSheet}
          onRenameSheet={handleRenameSheet}
          onDuplicateSheet={handleDuplicateSheet}
          onDeleteSheet={handleDeleteSheet}
          onHideSheet={handleHideSheet}
          onUnhideSheets={handleUnhideSheets}
          onMoveSheetLeft={() => handleMoveSheet("left")}
          onMoveSheetRight={() => handleMoveSheet("right")}
          activeSheetTabColor={activeSheet.tabColor ?? DEFAULT_SHEET_TAB_COLOR}
          onSheetTabColor={handleSheetTabColor}
          showGridlines={showGridlines}
          onToggleGridlines={handleToggleGridlines}
          showHeaders={showHeaders}
          onToggleHeaders={handleToggleHeaders}
          showFormulaBar={showFormulaBar}
          onToggleFormulaBar={handleToggleFormulaBar}
          showFormulas={showFormulas}
          onToggleShowFormulas={handleToggleShowFormulas}
          showSheetTabs={showSheetTabs}
          onToggleSheetTabs={handleToggleSheetTabs}
          onResetView={handleResetView}
          activeFormat={activeFormat}
          onPivot={() => {
            setPivotPanelOpen(true);
            setFindPanelOpen(false);
            setFilterPanelOpen(false);
            setValidationPanelOpen(false);
            setConditionalPanelOpen(false);
            setChartPanelOpen(false);
            setFunctionLibraryOpen(false);
            setNamedRangesOpen(false);
            setGoToPanelOpen(false);
            setFormulaAuditOpen(false);
          }}
          onChart={() => {
            setChartPanelOpen((isOpen) => !isOpen);
            setPivotPanelOpen(false);
            setFindPanelOpen(false);
            setFilterPanelOpen(false);
            setValidationPanelOpen(false);
            setConditionalPanelOpen(false);
            setFunctionLibraryOpen(false);
            setNamedRangesOpen(false);
            setGoToPanelOpen(false);
            setFormulaAuditOpen(false);
          }}
          onBold={() => applyFormat({ bold: !activeFormat.bold }, "Applied bold")}
          onItalic={() => applyFormat({ italic: !activeFormat.italic }, "Applied italic")}
          onWrapText={() => applyFormat({ wrapText: !activeFormat.wrapText }, activeFormat.wrapText ? "Unwrapped text" : "Wrapped text")}
          onNumberFormat={(numberFormat) => applyFormat({ numberFormat }, `Applied ${numberFormat} format`)}
          onHorizontalAlign={(horizontalAlign) => applyFormat({ horizontalAlign }, `Aligned ${horizontalAlign}`)}
          onVerticalAlign={(verticalAlign) => applyFormat({ verticalAlign }, `Aligned ${verticalAlign}`)}
          onBorders={applyBorders}
          onTextColor={(color) => applyFormat({ textColor: color }, "Changed text color")}
          onFillColor={(color) => applyFormat({ backgroundColor: color }, "Changed fill color")}
        />
        <input
          ref={fileInputRef}
          className="hidden-file-input"
          type="file"
          accept=".csv,text/csv"
          aria-label="CSV file"
          onChange={(event) => {
            const file = event.currentTarget.files?.[0];
            handleImportCsv(file);
            event.currentTarget.value = "";
          }}
        />
        <input
          ref={xlsxInputRef}
          className="hidden-file-input"
          type="file"
          accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
          aria-label="XLSX file"
          onChange={(event) => {
            const file = event.currentTarget.files?.[0];
            handleImportXlsx(file);
            event.currentTarget.value = "";
          }}
        />
        {showFormulaBar ? (
          <FormulaBar
            nameBoxValue={nameBoxDraft}
            onNameBoxChange={setNameBoxDraft}
            onNameBoxCommit={commitNameBox}
            onNameBoxCancel={() => setNameBoxDraft(selectionName)}
            value={formulaDraft}
            onChange={setFormulaDraft}
            onCommit={commitFormulaBar}
            suggestions={formulaSuggestions}
            onSelectSuggestion={(name) => setFormulaDraft((current) => insertFormulaSuggestion(current, name))}
          />
        ) : null}
        <FunctionLibraryPanel
          isOpen={isFunctionLibraryOpen}
          onClose={() => setFunctionLibraryOpen(false)}
          onInsert={handleInsertFunction}
        />
        <NamedRangesPanel
          isOpen={isNamedRangesOpen}
          namedRanges={workbook.namedRanges ?? []}
          sheets={workbook.sheets}
          onClose={() => setNamedRangesOpen(false)}
          onSelect={handleSelectNamedRange}
          onDelete={handleDeleteNamedRange}
        />
        {isFormulaAuditOpen ? (
          <FormulaAuditPanel
            selectionLabel={selectionName}
            activeAddress={activeAddress}
            formula={activeFormula}
            precedents={formulaAuditPrecedents}
            dependents={formulaAuditDependents}
            onSelectReference={selectRangeReference}
            onClose={() => setFormulaAuditOpen(false)}
          />
        ) : null}
        {isGoToPanelOpen ? (
          <GoToPanel
            referenceValue={goToDraft}
            namedRanges={workbook.namedRanges ?? []}
            sheets={workbook.sheets}
            onReferenceChange={setGoToDraft}
            onGoToReference={handleGoToReference}
            onSelectNamedRange={(namedRange) => {
              selectNamedRange(namedRange);
              setGoToDraft("");
              setGoToPanelOpen(false);
            }}
            onClose={() => setGoToPanelOpen(false)}
          />
        ) : null}
        <PivotPanel
          headers={pivotHeaders}
          isOpen={isPivotPanelOpen}
          onClose={() => setPivotPanelOpen(false)}
          onCreate={handleCreatePivotTable}
        />
        <ChartPanel
          isOpen={isChartPanelOpen}
          onClose={() => setChartPanelOpen(false)}
          onCreate={handleCreateChart}
        />
        <DataValidationPanel
          isOpen={isValidationPanelOpen}
          onClose={() => setValidationPanelOpen(false)}
          onApply={applyValidation}
          onClear={clearValidation}
          validationRules={dataValidationRules}
          onDeleteRule={deleteDataValidationRule}
        />
        <ConditionalFormattingPanel
          isOpen={isConditionalPanelOpen}
          onClose={() => setConditionalPanelOpen(false)}
          onApply={applyConditionalFormatting}
          onClear={clearConditionalFormatting}
          rules={activeSheet.conditionalFormats ?? []}
          onDeleteRule={deleteConditionalFormatRule}
        />
        <FilterPanel
          isOpen={isFilterPanelOpen}
          onClose={() => setFilterPanelOpen(false)}
          onApply={applyFilter}
          onClear={clearFilters}
        />
        {isFindPanelOpen ? (
          <FindReplacePanel
            findValue={findDraft}
            replaceValue={replaceDraft}
            onFindChange={setFindDraft}
            onReplaceChange={setReplaceDraft}
            onFindNext={handleFindNext}
            onReplace={handleReplaceCurrent}
            onReplaceAll={handleReplaceAll}
            onClose={() => setFindPanelOpen(false)}
          />
        ) : null}
        <Grid
          sheet={activeSheet}
          formulaEngine={formulaEngine}
          selection={selection}
          editingCell={editingCell}
          zoomLevel={zoomLevel}
          showGridlines={showGridlines}
          showHeaders={showHeaders}
          showFormulas={showFormulas}
          freezeTopRow={freezeTopRow}
          freezeFirstColumn={freezeFirstColumn}
          getCellFormat={(address) => activeSheet.formats[address]}
          getCellComment={(address) => activeSheet.comments?.[address] ?? null}
          getCellHyperlink={(address) => activeSheet.hyperlinks?.[address] ?? null}
          getCellReadOnly={(address) => getCellReadOnly(workbook, activeSheet.id, address)}
          getCellValidation={(address) => activeSheet.validations[address]}
          getCellConditionalFormatRules={(address) => getCellConditionalFormatRules(workbook, activeSheet.id, address)}
          scrollRef={gridScrollRef}
          onSelectionChange={handleSelectionChange}
          onStartEdit={(address) => {
            if (!ensureEditableAddress(address)) {
              return;
            }
            setEditingCell({ address, value: String(getCellContent(workbook, activeSheet.id, address) ?? "") });
          }}
          onEditValueChange={(value) => setEditingCell((current) => (current ? { ...current, value } : current))}
          onCommitEdit={(address, value) => {
            if (commitCell(address, value)) {
              setEditingCell(null);
            }
          }}
          onCancelEdit={() => setEditingCell(null)}
          onPasteText={pasteText}
          onKeyCommand={handleKeyCommand}
          onAutoFill={handleAutoFill}
          onCellContextMenu={openCellContextMenu}
          onAutoFilterColumn={applyAutoFilterColumn}
          onClearAutoFilterColumn={clearAutoFilterColumn}
          onSortAutoFilterColumn={sortAutoFilterColumn}
          onTogglePivotDrilldown={handleTogglePivotDrilldown}
          onColumnResize={handleColumnResize}
          onRowResize={handleRowResize}
        />
        {cellContextMenu ? (
          <CellContextMenu
            address={cellContextMenu.address}
            x={cellContextMenu.x}
            y={cellContextMenu.y}
            canPasteSpecial={Boolean(richClipboard)}
            isWrapped={Boolean(activeFormat.wrapText)}
            onClose={() => setCellContextMenu(null)}
            onCut={cutSelection}
            onCopy={copySelection}
            onPaste={handlePasteAll}
            onPasteValues={handlePasteValues}
            onPasteFormats={handlePasteFormats}
            onClearAll={handleClearAll}
            onClearContents={handleClearSelection}
            onClearFormats={handleClearFormats}
            onClearConditionalFormats={clearConditionalFormatting}
            onClearHyperlinks={handleClearHyperlinks}
            onClearValidation={clearValidation}
            onClearComments={handleClearComments}
            onToggleWrapText={() =>
              applyFormat({ wrapText: !activeFormat.wrapText }, activeFormat.wrapText ? "Unwrapped text" : "Wrapped text")
            }
            onInsertRow={handleInsertRows}
            onDeleteRow={handleDeleteRows}
            onInsertColumn={handleInsertColumns}
            onDeleteColumn={handleDeleteColumns}
            onComment={handleComment}
            onLink={handleLink}
          />
        ) : null}
        <SheetCharts
          charts={activeSheet.charts ?? []}
          sheet={activeSheet}
          formulaEngine={formulaEngine}
          onDelete={handleDeleteChart}
        />
        {showSheetTabs ? (
          <SheetTabs
            sheets={workbook.sheets}
            activeSheetId={activeSheet.id}
            onSelect={(sheetId) => {
              const nextWorkbook = setActiveSheet(workbook, sheetId);
              if (nextWorkbook === workbook) {
                return;
              }
              setHistory({ ...history, present: nextWorkbook });
              setSelection(INITIAL_SELECTION);
            }}
            onAdd={handleAddSheet}
          />
        ) : null}
        <StatusBar
          status={status}
          activeAddress={activeAddress}
          selectedCount={getRangeAddresses(selection).length}
          selectionSummary={selectionSummary}
          formulaFunctions="HyperFormula 418+ functions"
          zoomLevel={zoomLevel}
          onZoomOut={() => handleZoom(-ZOOM_STEP)}
          onZoomIn={() => handleZoom(ZOOM_STEP)}
          onResetZoom={handleResetZoom}
        />
      </section>
    </main>
  );
}

function readInitialWorkbook(): WorkbookModel {
  try {
    return loadWorkbook(window.localStorage);
  } catch {
    return createBlankWorkbook();
  }
}

function replaceActiveSheetWithRows(workbook: WorkbookModel, rows: string[][]): WorkbookModel {
  const activeSheet = getActiveSheet(workbook);
  const rowCount = Math.max(100, rows.length);
  const columnCount = Math.max(26, ...rows.map((row) => row.length));
  const cells: Record<string, CellContent> = {};

  rows.forEach((row, rowIndex) => {
    row.forEach((value, columnIndex) => {
      if (value !== "") {
        cells[formatCellAddress({ row: rowIndex, column: columnIndex })] = value;
      }
    });
  });

  return {
    ...workbook,
    sheets: workbook.sheets.map((sheet) =>
      sheet.id === activeSheet.id
        ? {
            ...sheet,
            rowCount,
            columnCount,
            cells,
            formats: {},
            columnWidths: {},
            rowHeights: {},
            hiddenColumns: {},
            hiddenRows: {},
            freezeTopRow: false,
            freezeFirstColumn: false,
            comments: {},
            hyperlinks: {},
            validations: {},
            conditionalFormats: [],
            autoFilterRange: undefined,
            filters: [],
            charts: [],
            merges: [],
            protection: { isProtected: false, lockedCells: {}, unlockedCells: {} }
          }
        : sheet
    ),
    namedRanges: (workbook.namedRanges ?? []).filter((namedRange) => namedRange.sheetId !== activeSheet.id)
  };
}

function activeSheetToRows(sheet: WorkbookModel["sheets"][number]): string[][] {
  const populatedAddresses = Object.keys(sheet.cells);
  const bounds = populatedAddresses.reduce(
    (current, address) => {
      const coord = parseCellAddress(address);
      return {
        row: Math.max(current.row, coord.row),
        column: Math.max(current.column, coord.column)
      };
    },
    { row: 0, column: 0 }
  );

  return Array.from({ length: bounds.row + 1 }, (_, row) =>
    Array.from({ length: bounds.column + 1 }, (_, column) => {
      const content = sheet.cells[formatCellAddress({ row, column })];
      return content === undefined || content === null ? "" : String(content);
    })
  );
}

function normalizeHyperlinkUrl(value: string): string | null {
  const trimmed = value.trim();
  if (!trimmed) {
    return "";
  }

  if (/\s/.test(trimmed)) {
    return null;
  }

  if (/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(trimmed)) {
    return `mailto:${trimmed}`;
  }

  if (/^(https?:\/\/|mailto:)/i.test(trimmed)) {
    return trimmed;
  }

  if (/^[a-z][a-z0-9+.-]*:/i.test(trimmed)) {
    return null;
  }

  return `https://${trimmed}`;
}

function borderPresetLabel(preset: Exclude<BorderPreset, "none">): string {
  switch (preset) {
    case "all":
      return "all";
    case "outer":
      return "outer";
    case "top":
      return "top";
    case "right":
      return "right";
    case "bottom":
      return "bottom";
    case "left":
      return "left";
  }
}

function formatSelectionAddress(selection: CellRange): string {
  const normalized = normalizeRange(selection);
  const start = formatCellAddress(normalized.start);
  const end = formatCellAddress(normalized.end);
  return start === end ? start : `${start}:${end}`;
}

function autoFunctionLabel(functionName: AutoFunctionName): string {
  const labels: Record<AutoFunctionName, string> = {
    SUM: "AutoSum",
    AVERAGE: "Average",
    COUNT: "Count",
    MAX: "Max",
    MIN: "Min"
  };

  return labels[functionName];
}

function getAutoFillWriteRange(sourceRange: CellRange, targetRange: CellRange): CellRange | null {
  const source = normalizeRange(sourceRange);
  const target = normalizeRange(targetRange);

  if (
    target.start.row === source.start.row &&
    target.start.column === source.start.column &&
    target.end.column === source.end.column &&
    target.end.row > source.end.row
  ) {
    return {
      start: { row: source.end.row + 1, column: source.start.column },
      end: { row: target.end.row, column: source.end.column }
    };
  }

  if (
    target.start.row === source.start.row &&
    target.start.column === source.start.column &&
    target.end.row === source.end.row &&
    target.end.column > source.end.column
  ) {
    return {
      start: { row: source.start.row, column: source.end.column + 1 },
      end: { row: source.end.row, column: target.end.column }
    };
  }

  return null;
}

function getFormulaDependents(sheet: SheetModel, selection: CellRange): FormulaAuditReference[] {
  const targetRange = normalizeRange(selection);

  return Object.entries(sheet.cells).flatMap(([address, content]) => {
    if (typeof content !== "string" || !content.startsWith("=")) {
      return [];
    }

    const references = extractFormulaReferences(content);
    if (!references.some((reference) => rangesIntersect(reference.range, targetRange))) {
      return [];
    }

    const coord = parseCellAddress(address);
    return [
      {
        label: address,
        range: { start: coord, end: coord },
        formula: content
      }
    ];
  });
}

function rangesIntersect(left: CellRange, right: CellRange): boolean {
  const normalizedLeft = normalizeRange(left);
  const normalizedRight = normalizeRange(right);
  return (
    normalizedLeft.start.row <= normalizedRight.end.row &&
    normalizedLeft.end.row >= normalizedRight.start.row &&
    normalizedLeft.start.column <= normalizedRight.end.column &&
    normalizedLeft.end.column >= normalizedRight.start.column
  );
}

function parseNameBoxRange(value: string, sheet: SheetModel): CellRange | null {
  try {
    const range = normalizeRange(parseRangeAddress(value));
    return {
      start: {
        row: clamp(range.start.row, 0, sheet.rowCount - 1),
        column: clamp(range.start.column, 0, sheet.columnCount - 1)
      },
      end: {
        row: clamp(range.end.row, 0, sheet.rowCount - 1),
        column: clamp(range.end.column, 0, sheet.columnCount - 1)
      }
    };
  } catch {
    return null;
  }
}

function findNamedRangeByName(namedRanges: NamedRange[], name: string): NamedRange | null {
  const normalizedName = name.trim().toLowerCase();
  return namedRanges.find((namedRange) => namedRange.name.toLowerCase() === normalizedName) ?? null;
}

function replaceGeneratedPivotSheetRows(
  workbook: WorkbookModel,
  sheetId: string,
  pivot: PivotSheetMetadata
): WorkbookModel {
  const matrix = materializePivotRows(pivot);
  const columnCount = Math.max(26, matrix.reduce((width, row) => Math.max(width, row.length), 0));
  const rowCount = Math.max(100, matrix.length);
  const cells: Record<string, CellContent> = {};

  matrix.forEach((rowValues, row) => {
    rowValues.forEach((content, column) => {
      if (content !== "") {
        cells[formatCellAddress({ row, column })] = content;
      }
    });
  });

  let nextWorkbook: WorkbookModel = {
    ...workbook,
    sheets: workbook.sheets.map((sheet) =>
      sheet.id === sheetId
        ? {
            ...sheet,
            rowCount,
            columnCount,
            cells,
            formats: {},
            comments: {},
            hyperlinks: {},
            validations: {},
            conditionalFormats: [],
            filters: [],
            charts: [],
            merges: [],
            autoFilterRange: undefined,
            pivot
          }
        : sheet
    )
  };

  nextWorkbook = setCellFormat(nextWorkbook, sheetId, rowRange(0, columnCount), {
    bold: true,
    textColor: "#17634a",
    backgroundColor: "#eaf7f2"
  });

  const grandTotalRow = materializedPivotBaseRow(pivot, pivot.baseRows.length - 1);
  matrix.forEach((_, row) => {
    const rowKind = getPivotMaterializedRowKind(pivot, row);
    if (rowKind?.kind === "detail-header") {
      nextWorkbook = setCellFormat(nextWorkbook, sheetId, rowRange(row, columnCount), {
        bold: true,
        textColor: "#475569",
        backgroundColor: "#f8fafc"
      });
      return;
    }

    if (rowKind?.kind === "detail-row") {
      nextWorkbook = setCellFormat(nextWorkbook, sheetId, rowRange(row, columnCount), {
        textColor: "#334155",
        backgroundColor: "#fffdf7"
      });
      return;
    }

    if (row > 0 && row === grandTotalRow) {
      nextWorkbook = setCellFormat(nextWorkbook, sheetId, rowRange(row, columnCount), {
        bold: true,
        backgroundColor: "#f1f5f8"
      });
    }
  });

  return nextWorkbook;
}

function materializedPivotBaseRow(pivot: PivotSheetMetadata, targetBaseRow: number): number | null {
  let materializedRow = 0;

  for (let baseRow = 0; baseRow < pivot.baseRows.length; baseRow += 1) {
    if (baseRow === targetBaseRow) {
      return materializedRow;
    }

    materializedRow += 1;
    for (const entry of Object.values(pivot.drilldowns)) {
      if (entry.baseRow === baseRow && pivot.expanded[entry.id]) {
        materializedRow += entry.sourceRows.length + 1;
      }
    }
  }

  return null;
}

function selectedRangeToDisplayRows(sheet: SheetModel, selection: CellRange, formulaEngine: FormulaEngine): string[][] {
  const range = inferSourceRange(sheet, selection);
  const rows: string[][] = [];

  for (let row = range.start.row; row <= range.end.row; row += 1) {
    const values: string[] = [];
    for (let column = range.start.column; column <= range.end.column; column += 1) {
      values.push(formulaEngine.getDisplayValue(sheet.id, formatCellAddress({ row, column })));
    }
    rows.push(values);
  }

  return trimEmptyEdges(rows);
}

function inferSourceRange(sheet: SheetModel, selection: CellRange): CellRange {
  const normalized = normalizeRange(selection);
  const hasExplicitSelection =
    normalized.start.row !== normalized.end.row || normalized.start.column !== normalized.end.column;
  if (hasExplicitSelection) {
    return normalized;
  }

  const populatedCoords = Object.keys(sheet.cells).map(parseCellAddress);
  if (populatedCoords.length === 0) {
    return normalized;
  }

  return {
    start: {
      row: Math.min(...populatedCoords.map((coord) => coord.row)),
      column: Math.min(...populatedCoords.map((coord) => coord.column))
    },
    end: {
      row: Math.max(...populatedCoords.map((coord) => coord.row)),
      column: Math.max(...populatedCoords.map((coord) => coord.column))
    }
  };
}

function trimEmptyEdges(rows: string[][]): string[][] {
  let lastRow = rows.length - 1;
  while (lastRow >= 0 && rows[lastRow].every((value) => value === "")) {
    lastRow -= 1;
  }

  const keptRows = rows.slice(0, lastRow + 1);
  const lastColumn = keptRows.reduce((maxColumn, row) => Math.max(maxColumn, lastNonEmptyIndex(row)), -1);

  if (keptRows.length === 0 || lastColumn < 0) {
    return [];
  }

  return keptRows.map((row) => row.slice(0, lastColumn + 1));
}

function summarizeSelection(sheet: SheetModel, selection: CellRange, formulaEngine: FormulaEngine): string {
  const values = getRangeAddresses(selection).map((address) => formulaEngine.getDisplayValue(sheet.id, address));
  const nonEmptyValues = values.filter((value) => value.trim() !== "");
  if (nonEmptyValues.length === 0) {
    return "Count 0";
  }

  const numericValues = nonEmptyValues.map((value) => Number(value)).filter((value) => Number.isFinite(value));
  if (numericValues.length === 0) {
    return `Count ${nonEmptyValues.length}`;
  }

  const sum = numericValues.reduce((total, value) => total + value, 0);
  const min = Math.min(...numericValues);
  const max = Math.max(...numericValues);
  return `Count ${nonEmptyValues.length}  Sum ${formatSummaryNumber(sum)}  Avg ${formatSummaryNumber(
    sum / numericValues.length
  )}  Min ${formatSummaryNumber(min)}  Max ${formatSummaryNumber(max)}`;
}

type FindMatch = {
  address: string;
  row: number;
  column: number;
  content: Exclude<CellContent, null>;
};

function findMatches(sheet: SheetModel, query: string): FindMatch[] {
  const normalizedQuery = query.trim().toLocaleLowerCase();
  if (!normalizedQuery) {
    return [];
  }

  return Object.entries(sheet.cells)
    .flatMap(([address, content]) => {
      if (content === null || !String(content).toLocaleLowerCase().includes(normalizedQuery)) {
        return [];
      }
      const coord = parseCellAddress(address);
      return [{ address, row: coord.row, column: coord.column, content }];
    })
    .sort((left, right) => left.row - right.row || left.column - right.column);
}

function replaceFirstMatch(value: string, query: string, replacement: string): string {
  const index = value.toLocaleLowerCase().indexOf(query.toLocaleLowerCase());
  if (index < 0) {
    return value;
  }
  return `${value.slice(0, index)}${replacement}${value.slice(index + query.length)}`;
}

function replaceEveryMatch(value: string, query: string, replacement: string): string {
  const escapedQuery = query.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  return value.replace(new RegExp(escapedQuery, "gi"), replacement);
}

function normalizeClipboardText(text: string): string {
  return text.replace(/\r\n/g, "\n").replace(/\r/g, "\n");
}

function isEditableEventTarget(target: EventTarget | null): boolean {
  if (!(target instanceof HTMLElement)) {
    return false;
  }
  return target.isContentEditable || ["INPUT", "SELECT", "TEXTAREA"].includes(target.tagName);
}

function formatSummaryNumber(value: number): string {
  return Number.isInteger(value) ? String(value) : value.toFixed(2).replace(/\.?0+$/, "");
}

function pluralize(count: number, singular: string): string {
  return `${count} ${count === 1 ? singular : `${singular}s`}`;
}

function getPivotHeaders(rows: string[][]): string[] {
  return Array.from(new Set((rows[0] ?? []).map((header) => header.trim()).filter(Boolean)));
}

function lastNonEmptyIndex(row: string[]): number {
  for (let index = row.length - 1; index >= 0; index -= 1) {
    if (row[index] !== "") {
      return index;
    }
  }
  return -1;
}

function nextPivotSheetName(workbook: WorkbookModel): string {
  const existingNames = new Set(workbook.sheets.map((sheet) => sheet.name));
  let index = 1;
  let name = `Pivot ${index}`;
  while (existingNames.has(name)) {
    index += 1;
    name = `Pivot ${index}`;
  }
  return name;
}

function rowRange(row: number, columnCount: number): CellRange {
  return {
    start: { row, column: 0 },
    end: { row, column: Math.max(columnCount - 1, 0) }
  };
}

function clamp(value: number, min: number, max: number): number {
  return Math.min(max, Math.max(min, value));
}

function hasVisibleIndexAfterHiding(
  count: number,
  hiddenIndexes: Record<string, boolean>,
  startIndex: number,
  endIndex: number
): boolean {
  const start = Math.min(startIndex, endIndex);
  const end = Math.max(startIndex, endIndex);
  for (let index = 0; index < count; index += 1) {
    if (!hiddenIndexes[String(index)] && (index < start || index > end)) {
      return true;
    }
  }
  return false;
}

function findNearestVisibleIndex(count: number, hiddenIndexes: Record<string, boolean>, preferredIndex: number): number {
  if (count <= 0) {
    return 0;
  }

  const preferred = clamp(preferredIndex, 0, count - 1);
  for (let offset = 0; offset < count; offset += 1) {
    const forward = preferred + offset;
    if (forward < count && !hiddenIndexes[String(forward)]) {
      return forward;
    }

    const backward = preferred - offset;
    if (backward >= 0 && !hiddenIndexes[String(backward)]) {
      return backward;
    }
  }

  return preferred;
}
