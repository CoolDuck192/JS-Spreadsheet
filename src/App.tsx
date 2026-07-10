import {
  useCallback,
  useEffect,
  useMemo,
  useRef,
  useState,
  useSyncExternalStore,
  type CSSProperties
} from "react";
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
import { Grid, type CommitEditMove, type GridScrollApi } from "./components/Grid";
import { NamedRangesPanel } from "./components/NamedRangesPanel";
import { PivotPanel } from "./components/PivotPanel";
import { SheetCharts } from "./components/SheetCharts";
import { SheetTabs } from "./components/SheetTabs";
import { StatusBar } from "./components/StatusBar";
import { Toolbar } from "./components/Toolbar";
import type {
  BorderPreset,
  CellContent,
  CellCoord,
  CellFormat,
  CellRange,
  ConditionalFormatRule,
  DataValidationRule,
  FilterOperator,
  MixedFormatValue,
  SelectionFormatSummary,
  SheetChartType,
  SheetModel,
  NamedRange,
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
import { createAutoFitColumnPlan, createAutoFitRowPlan, textToAutoFitColumnWidth } from "./lib/autoFit";
import { isRowVisibleForFilter } from "./lib/filters";
import { createAutoSumPlan, formatRangeAddress, type AutoFunctionName } from "./lib/autoSum";
import { createChartData } from "./lib/charts";
import { parseCsv, serializeCsv } from "./lib/csv";
import { formatDisplayValue } from "./lib/displayFormat";
import { summarizeDataValidationRules, type DataValidationSummary } from "./lib/dataValidationSummary";
import { createBrowserTokenProvider, type TokenProvider } from "./lib/googleAuth";
import { importWorkbookFromGoogleSheets } from "./lib/googleSheets";
import { extractFormulaReferences } from "./lib/formulaReferences";
import { getFormulaSuggestions, insertFormulaSuggestion } from "./lib/formulaSuggestions";
import { createPivotTableWithDetails, type PivotConfig, type PivotDrillDownGrid } from "./lib/pivot";
import { exportWorkbookToXlsx, importWorkbookFromXlsx } from "./lib/xlsx";
import {
  isHorizontalAutoFill,
  isVerticalAutoFill,
  copyRange,
  copyRichRange,
  createBlankWorkbook,
  getActiveSheet,
  getCellComment,
  getCellConditionalFormatRules,
  getCellContent,
  getCellFormat,
  getCellHyperlink,
  getCellReadOnly,
  getNamedRangeForSelection,
  isValidNamedRangeName
} from "./lib/workbook";
import type { FormulaEngine } from "./lib/formulaEngine";
import type { RichClipboardRange, RichPasteMode } from "./lib/workbook";
import type {
  WorkbookCommandResult,
  WorkbookDiagnosticEvent,
  WorkbookSession
} from "./core/workbook/WorkbookSession";
import type { WorkbookCommand } from "./core/workbook/commands";
import type {
  SpreadsheetServices,
  WorkbookExportArtifact,
  WorkbookExporter,
  WorkbookImporter
} from "./core/workbook/services";
import { useWorkbookSession } from "./react/useWorkbookSession";

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

export type WorkbookStorage = {
  load(): WorkbookModel | null | Promise<WorkbookModel | null>;
  save(workbook: WorkbookModel): void | Promise<void>;
  clear?(): void | Promise<void>;
};

export type WorkbookFeatureConfiguration = Readonly<{
  toolbar?: boolean;
  formulaBar?: boolean;
  sheetTabs?: boolean;
  import?: boolean;
  export?: boolean;
  charts?: boolean;
  structuredTables?: boolean;
  googleSheets?: boolean;
}>;

export type WorkbookThemeToken =
  | "font-family"
  | "font-size"
  | "surface"
  | "surface-muted"
  | "text"
  | "text-muted"
  | "border"
  | "accent"
  | "accent-contrast"
  | "selection"
  | "danger";

export type WorkbookChangeEvent = Readonly<{
  workbook: WorkbookModel;
  revision: string;
  previousRevision: string;
  origin: "command" | "undo" | "redo" | "import" | "external" | "storage";
  commandId?: string;
}>;

export type SpreadsheetErrorEvent = Readonly<{
  code: string;
  message: string;
  recoverable: boolean;
}>;

export type SpreadsheetCommonProps = Readonly<{
  className?: string;
  style?: CSSProperties;
  features?: WorkbookFeatureConfiguration;
  services?: SpreadsheetServices;
  theme?: Partial<Record<WorkbookThemeToken, string>>;
  onDiagnostic?: (event: WorkbookDiagnosticEvent) => void;
  onCommandResult?: (event: Readonly<{
    command: WorkbookCommand;
    result: WorkbookCommandResult;
  }>) => void;
  onWorkbookChangeEvent?: (event: WorkbookChangeEvent) => void;
  onError?: (event: SpreadsheetErrorEvent) => void;
}>;

export type SpreadsheetProps = SpreadsheetCommonProps & (
  | {
      session: WorkbookSession;
      workbook?: never;
      defaultWorkbook?: never;
      onWorkbookChange?: never;
      storage?: never;
    }
  | {
      workbook: WorkbookModel;
      onWorkbookChange(workbook: WorkbookModel): void;
      session?: never;
      defaultWorkbook?: never;
      storage?: WorkbookStorage | false;
    }
  | {
      defaultWorkbook?: WorkbookModel;
      session?: never;
      workbook?: never;
      onWorkbookChange?: never;
      storage?: WorkbookStorage | false;
    }
);

type SpreadsheetWorkbookProps = SpreadsheetCommonProps & {
  session: WorkbookSession;
  suppliedSession: boolean;
};

export function Spreadsheet(props: SpreadsheetProps) {
  if (props.session) {
    return <SuppliedSpreadsheet {...props} session={props.session} />;
  }
  return <OwnedSpreadsheet {...props} />;
}

function SuppliedSpreadsheet(
  props: Extract<SpreadsheetProps, { session: WorkbookSession }>
) {
  const onCommandResultRef = useRef(props.onCommandResult);
  onCommandResultRef.current = props.onCommandResult;
  const session = useMemo<WorkbookSession>(() => ({
    getSnapshot: props.session.getSnapshot,
    getCellEvaluation: props.session.getCellEvaluation,
    subscribe: props.session.subscribe,
    subscribeDiagnostics: props.session.subscribeDiagnostics,
    dispatch(commandOrEnvelope) {
      const result = props.session.dispatch(commandOrEnvelope);
      invokeHostCallback(onCommandResultRef.current, {
        command: "intent" in commandOrEnvelope ? commandOrEnvelope.intent : commandOrEnvelope,
        result
      });
      return result;
    },
    replaceWorkbook(workbook, options) {
      const command: WorkbookCommand = {
        type: "workbook.replace",
        workbook,
        history: options?.history === "preserve" ? "commit" : "reset"
      };
      const result = props.session.replaceWorkbook(workbook, options);
      invokeHostCallback(onCommandResultRef.current, { command, result });
      return result;
    },
    // This facade is view-local. Ownership always remains with the host.
    destroy() {}
  }), [props.session]);
  return <SpreadsheetWorkbook {...props} session={session} suppliedSession />;
}

function OwnedSpreadsheet(
  props: Extract<SpreadsheetProps, { session?: never }>
) {
  const session = useWorkbookSession(props.workbook !== undefined
    ? {
        workbook: props.workbook,
        onWorkbookChange: props.onWorkbookChange,
        storage: props.storage,
        services: props.services,
        onDiagnostic: props.onDiagnostic,
        onCommandResult: props.onCommandResult,
        onWorkbookChangeEvent: props.onWorkbookChangeEvent,
        onError: props.onError
      }
    : {
        defaultWorkbook: props.defaultWorkbook,
        storage: props.storage,
        services: props.services,
        onDiagnostic: props.onDiagnostic,
        onCommandResult: props.onCommandResult,
        onWorkbookChangeEvent: props.onWorkbookChangeEvent,
        onError: props.onError
      });
  return <SpreadsheetWorkbook {...props} session={session} suppliedSession={false} />;
}

function SpreadsheetWorkbook({
  session,
  suppliedSession,
  className,
  style,
  features,
  services,
  theme,
  onDiagnostic,
  onWorkbookChangeEvent,
  onError
}: SpreadsheetWorkbookProps) {
  const sessionSnapshot = useSyncExternalStore(
    session.subscribe,
    session.getSnapshot,
    session.getSnapshot
  );
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
  const [isDropTargetActive, setDropTargetActive] = useState(false);
  // Drill-down metadata for pivot sheets created this session, keyed by sheet id.
  // Double-clicking a pivot value cell opens the contributing source rows, like Excel.
  // Sheet ids are reused after deletion, so entries are pruned on sheet delete,
  // cleared on workbook replacement, and cross-checked against the stored table
  // before use (covers undo/redo resurrecting an id for an unrelated sheet).
  const [pivotDrillDowns, setPivotDrillDowns] = useState<
    Record<string, { sourceRows: string[][]; drillDown: PivotDrillDownGrid; table: string[][] }>
  >({});
  const [showGridlines, setShowGridlines] = useState(true);
  const [showHeaders, setShowHeaders] = useState(true);
  const [showFormulaBar, setShowFormulaBar] = useState(true);
  const [showFormulas, setShowFormulas] = useState(false);
  const [showSheetTabs, setShowSheetTabs] = useState(true);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const xlsxInputRef = useRef<HTMLInputElement>(null);
  const gridScrollRef = useRef<HTMLDivElement>(null);
  const gridApiRef = useRef<GridScrollApi | null>(null);
  const gridEditCommitInProgressRef = useRef(false);
  const googleTokenProviderRef = useRef<{
    factory: (clientId: string) => TokenProvider;
    clientId: string;
    provider: TokenProvider;
  } | null>(null);
  const registerGridScrollApi = useCallback((api: GridScrollApi) => {
    gridApiRef.current = api;
  }, []);

  function getComponentGoogleTokenProvider(clientId: string): TokenProvider {
    const factory = services?.googleTokenProviderFactory ?? createBrowserTokenProvider;
    const cached = googleTokenProviderRef.current;
    if (cached?.factory === factory && cached.clientId === clientId) {
      return cached.provider;
    }

    const provider = factory(clientId);
    googleTokenProviderRef.current = { factory, clientId, provider };
    return provider;
  }

  const lastSuppliedEventSnapshot = useRef(sessionSnapshot);
  useEffect(() => {
    if (!suppliedSession) {
      return;
    }
    return session.subscribeDiagnostics((event) => {
      invokeHostCallback(onDiagnostic, event);
      const previous = lastSuppliedEventSnapshot.current;
      const next = session.getSnapshot();
      lastSuppliedEventSnapshot.current = next;
      if (previous.revision === next.revision || previous.workbook === next.workbook) {
        return;
      }
      invokeHostCallback(onWorkbookChangeEvent, {
        workbook: next.workbook,
        revision: next.revision,
        previousRevision: previous.revision,
        origin: diagnosticOrigin(event),
        ...(event.commandId ? { commandId: event.commandId } : {})
      });
    });
  }, [onDiagnostic, onWorkbookChangeEvent, session, suppliedSession]);

  const validatedServices = useRef(false);
  useEffect(() => {
    if (validatedServices.current) {
      return;
    }
    validatedServices.current = true;
    if (hasInvalidServiceRegistry(services?.importers) || hasInvalidServiceRegistry(services?.exporters)) {
      invokeHostCallback(onError, {
        code: "service.registry.invalid",
        message: "A workbook service registry contains an invalid key",
        recoverable: true
      });
    }
  }, [onError, services]);

  const workbook = sessionSnapshot.workbook;
  const selection = sessionSnapshot.selection;
  const activeSheet = getActiveSheet(workbook);
  const freezeTopRow = Boolean(activeSheet.freezeTopRow);
  const freezeFirstColumn = Boolean(activeSheet.freezeFirstColumn);
  const activeAddress = formatCellAddress(selection.start);
  const selectionName = useMemo(
    () => getNamedRangeForSelection(workbook, activeSheet.id, selection)?.name ?? formatSelectionAddress(selection),
    [activeSheet.id, selection, workbook]
  );
  const workbookRef = useRef(workbook);
  workbookRef.current = workbook;
  const formulaEngine = useMemo<FormulaEngine>(() => ({
    getComputedValue: session.getCellEvaluation,
    getDisplayValue(sheetId, address) {
      return displaySessionValue(session.getCellEvaluation(sheetId, address));
    },
    getRawContent(sheetId, address) {
      return getCellContent(workbookRef.current, sheetId, address);
    },
    update() {},
    rebuild() {},
    destroy() {}
  }), [session]);
  const activeFormat = getCellFormat(workbook, activeSheet.id, activeAddress);
  // Ribbon state reflects the WHOLE selection, Excel-style: a control shows a
  // concrete value when every selected cell agrees and "mixed" otherwise.
  const selectionFormat = useMemo(() => summarizeSelectionFormats(activeSheet, selection), [activeSheet, selection]);
  const formulaSuggestions = useMemo(() => getFormulaSuggestions(formulaDraft), [formulaDraft]);
  const activeFormulaContent = getCellContent(workbook, activeSheet.id, activeAddress);
  const activeFormula = typeof activeFormulaContent === "string" && activeFormulaContent.startsWith("=") ? activeFormulaContent : "";
  const formulaAuditPrecedents = useMemo(() => extractFormulaReferences(activeFormula), [activeFormula]);
  const formulaAuditDependents = useMemo(
    // Scans every formula in the sheet — only worth it while the audit panel is open.
    () => (isFormulaAuditOpen ? getFormulaDependents(activeSheet, selection) : []),
    [activeSheet, isFormulaAuditOpen, selection]
  );
  const pivotSourceRows = useMemo(
    () => (isPivotPanelOpen ? selectedRangeToDisplayRows(activeSheet, selection, formulaEngine) : []),
    [activeSheet, formulaEngine, isPivotPanelOpen, selection]
  );
  const pivotHeaders = useMemo(() => getPivotHeaders(pivotSourceRows), [pivotSourceRows]);
  const pivotSourceLabel = useMemo(
    () => (isPivotPanelOpen ? formatSelectionAddress(inferSourceRange(activeSheet, selection)) : undefined),
    [activeSheet, isPivotPanelOpen, selection]
  );
  const dataValidationRules = useMemo(
    () => summarizeDataValidationRules(activeSheet.validations ?? {}),
    [activeSheet.validations]
  );
  const selectionSummary = useMemo(
    () => summarizeSelection(activeSheet, selection, formulaEngine),
    [activeSheet, formulaEngine, selection]
  );

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

  function dispatchCommand(command: WorkbookCommand, nextStatus?: string): WorkbookCommandResult {
    const result = session.dispatch(command);
    if (result.status === "committed") {
      if (result.changed && nextStatus) {
        setStatus(nextStatus);
      }
      return result;
    }
    if (result.status === "conflict") {
      setStatus(`Workbook changed; retry from revision ${result.revision}`);
      return result;
    }
    if (result.status === "pending") {
      setStatus(`Workbook operation ${result.operationId} is pending`);
      return result;
    }
    const issue = result.issues?.[0];
    setStatus(
      issue?.code === "permission.readOnly" && issue.address
        ? `${issue.address} is read-only`
        : issue?.message ?? "Workbook change was rejected"
    );
    return result;
  }

  function setSelection(nextSelection: CellRange): WorkbookCommandResult {
    return dispatchCommand({ type: "selection.set", selection: nextSelection });
  }

  // Transitional compatibility helpers removed by the final bypass audit.
  function commitWorkbook(nextWorkbook: WorkbookModel, nextStatus = "Saved") {
    return dispatchCommand({
      type: "workbook.replace",
      workbook: nextWorkbook,
      history: "commit"
    }, nextStatus);
  }

  function applyHistoryTransition(direction: "undo" | "redo", nextStatus: string) {
    return dispatchCommand({ type: `history.${direction}` }, nextStatus);
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

  function commitCell(address: string, raw: string): boolean {
    return dispatchCommand({
      type: "cell.set",
      sheetId: activeSheet.id,
      address,
      input: raw
    }, "Saved").status === "committed";
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

    const rangeLabel = formatSelectionAddress(selection);
    dispatchCommand({
      type: "namedRange.define",
      namedRange: { name: nextName, sheetId: activeSheet.id, range: selection }
    }, `Named ${rangeLabel} as ${nextName}`);
  }

  function selectRangeReference(targetRange: CellRange) {
    const normalizedTarget = normalizeRange(targetRange);
    const targetLabel = formatSelectionAddress(normalizedTarget);
    setSelection(normalizedTarget);
    gridApiRef.current?.ensureCellVisible(normalizedTarget.start.row, normalizedTarget.start.column);
    setStatus(`Selected ${targetLabel}`);
  }

  function selectNamedRange(namedRange: NamedRange) {
    const normalizedTarget = normalizeRange(namedRange.range);
    dispatchCommand({
      type: "transaction",
      commands: [
        { type: "sheet.activate", sheetId: namedRange.sheetId },
        { type: "selection.set", selection: normalizedTarget }
      ]
    });
    gridApiRef.current?.ensureCellVisible(normalizedTarget.start.row, normalizedTarget.start.column);
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
    dispatchCommand({ type: "range.format", sheetId: activeSheet.id, range: selection, format }, nextStatus);
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
    const targetLabel = formatSelectionAddress(normalizedTarget);
    const result = dispatchCommand({
      type: "transaction",
      commands: [
        {
          type: "range.format.replace",
          sheetId: activeSheet.id,
          range: normalizedTarget,
          format: painter.format
        },
        { type: "selection.set", selection: normalizedTarget }
      ]
    }, `Painted format to ${targetLabel}`);
    if (result.status === "committed") {
      setFormatPainter(null);
    }
  }

  function applyBorders(preset: BorderPreset) {
    const selectionLabel = formatSelectionAddress(selection);
    const nextStatus =
      preset === "none"
        ? `Cleared borders from ${selectionLabel}`
        : `Applied ${borderPresetLabel(preset)} borders to ${selectionLabel}`;
    dispatchCommand({
      type: "range.borders",
      sheetId: activeSheet.id,
      range: selection,
      preset
    }, nextStatus);
  }

  function applyValidation(rule: DataValidationRule) {
    const result = dispatchCommand({
      type: "range.validation.set",
      sheetId: activeSheet.id,
      range: selection,
      rule
    }, "Applied data validation");
    if (result.status === "committed") setValidationPanelOpen(false);
  }

  function clearValidation() {
    const result = dispatchCommand({
      type: "range.validation.clear",
      sheetId: activeSheet.id,
      range: selection
    }, "Cleared data validation");
    if (result.status === "committed") setValidationPanelOpen(false);
  }

  function deleteDataValidationRule(summary: DataValidationSummary) {
    dispatchCommand({
      type: "range.validation.clear",
      sheetId: activeSheet.id,
      range: summary.range
    }, "Deleted data validation rule");
  }

  function applyConditionalFormatting(rule: {
    condition: ConditionalFormatRule["condition"];
    format: CellFormat;
  }) {
    const result = dispatchCommand({
      type: "range.conditionalFormat.add",
      sheetId: activeSheet.id,
      range: selection,
      rule: {
        id: nextConditionalFormatCommandId(activeSheet.conditionalFormats ?? []),
        range: selection,
        condition: rule.condition,
        format: rule.format
      }
    }, "Applied conditional formatting");
    if (result.status === "committed") setConditionalPanelOpen(false);
  }

  function clearConditionalFormatting() {
    const result = dispatchCommand({
      type: "range.conditionalFormat.clear",
      sheetId: activeSheet.id,
      range: selection
    }, "Cleared conditional formatting");
    if (result.status === "committed") setConditionalPanelOpen(false);
  }

  function deleteConditionalFormatRule(ruleId: string) {
    const rule = (activeSheet.conditionalFormats ?? []).find((candidate) => candidate.id === ruleId);
    if (!rule) {
      return;
    }
    dispatchCommand({
      type: "range.conditionalFormat.remove",
      sheetId: activeSheet.id,
      ruleId
    }, "Deleted conditional format rule");
  }

  function applyFilter(filter: { operator: FilterOperator; value: string }) {
    const normalized = normalizeRange(selection);
    dispatchCommand({
      type: "sheet.filter.set",
      sheetId: activeSheet.id,
      filter: {
        id: nextSheetFilterCommandId(activeSheet.filters ?? []),
        range: normalized,
        ...filter,
        column: normalized.start.column,
        hasHeader: normalized.start.row !== normalized.end.row
      }
    }, "Filter applied");
    setFilterPanelOpen(false);
  }

  function clearFilters() {
    dispatchCommand({ type: "sheet.filter.clear", sheetId: activeSheet.id }, "Filters cleared");
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
      dispatchCommand({
        type: "sheet.filter.clear",
        sheetId: activeSheet.id,
        column
      }, `Cleared filter from ${columnIndexToName(column)}`);
      return;
    }

    dispatchCommand({
      type: "sheet.filter.set",
      sheetId: activeSheet.id,
      filter: {
        id: nextSheetFilterCommandId(activeSheet.filters ?? []),
        range,
        column,
        operator: "equals",
        value: values[0] ?? "",
        values,
        hasHeader: true
      }
    }, `Filtered ${columnIndexToName(column)} by ${values.join(", ")}`);
  }

  function clearAutoFilterColumn(column: number) {
    const range = activeSheet.autoFilterRange ? normalizeRange(activeSheet.autoFilterRange) : null;
    if (!range) {
      return;
    }

    dispatchCommand({
      type: "sheet.filter.clear",
      sheetId: activeSheet.id,
      column
    }, `Cleared filter from ${columnIndexToName(column)}`);
  }

  function sortAutoFilterColumn(column: number, direction: "asc" | "desc") {
    const range = activeSheet.autoFilterRange ? normalizeRange(activeSheet.autoFilterRange) : null;
    if (!range) {
      return;
    }

    dispatchCommand({
      type: "range.sort",
      sheetId: activeSheet.id,
      range,
      direction,
      sortColumn: column
    }, direction === "asc" ? `Sorted ${columnIndexToName(column)} A to Z` : `Sorted ${columnIndexToName(column)} Z to A`);
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
    dispatchCommand({
      type: "transaction",
      commands: [
        { type: "cell.set", sheetId: activeSheet.id, address: targetAddress, input: plan.formula },
        { type: "selection.set", selection: { start: plan.target, end: plan.target } }
      ]
    },
      functionName === "SUM"
        ? `Inserted AutoSum for ${formatRangeAddress(plan.source)}`
        : `Inserted ${autoFunctionLabel(functionName)} for ${formatRangeAddress(plan.source)}`
    );
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
    dispatchCommand({ type: "namedRange.remove", name }, `Deleted named range ${name}`);
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
    const action = hasComment ? (hadComment ? "Updated" : "Added") : "Removed";
    dispatchCommand({
      type: "cell.comment.set",
      sheetId: activeSheet.id,
      address: activeAddress,
      comment: nextComment
    }, `${action} comment ${hasComment ? "to" : "from"} ${activeAddress}`);
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
    const action = hasLink ? (hadLink ? "Updated" : "Added") : "Removed";
    dispatchCommand({
      type: "cell.hyperlink.set",
      sheetId: activeSheet.id,
      address: activeAddress,
      hyperlink: normalizedUrl
    }, `${action} link ${hasLink ? "to" : "from"} ${activeAddress}`);
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

    dispatchCommand({
      type: "cell.hyperlink.set",
      sheetId: activeSheet.id,
      address: activeAddress,
      hyperlink: null
    }, `Removed link from ${activeAddress}`);
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

    dispatchCommand({
      type: "range.merge",
      sheetId: activeSheet.id,
      range: normalized
    }, `Merged ${formatSelectionAddress(normalized)}`);
  }

  function handleUnmergeCells() {
    const normalized = normalizeRange(selection);
    const result = dispatchCommand({
      type: "range.unmerge",
      sheetId: activeSheet.id,
      range: normalized
    }, `Unmerged ${formatSelectionAddress(normalized)}`);
    if (result.status === "committed" && !result.changed) {
      setStatus("No merged cells in selection");
    }
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
    const result = dispatchCommand({
      type: "sheet.chart.add",
      sheetId: activeSheet.id,
      chart: {
        id: nextSheetChartCommandId(activeSheet.charts ?? []),
        range: normalized,
        anchor: {
        row: normalized.start.row,
        column: Math.min(activeSheet.columnCount - 1, normalized.end.column + 1)
        },
        title: config.title || chartData.title,
        type: config.type
      }
    }, "Created chart");
    if (result.status === "committed") setChartPanelOpen(false);
  }

  function handleDeleteChart(chartId: string) {
    if (!ensureSheetStructureEditable()) {
      return;
    }
    dispatchCommand({ type: "sheet.chart.delete", sheetId: activeSheet.id, chartId }, "Deleted chart");
  }

  function handleColumnResize(column: number, width: number) {
    if (!ensureSheetStructureEditable()) {
      return;
    }
    dispatchCommand({
      type: "columns.resize",
      sheetId: activeSheet.id,
      columns: [column],
      width
    }, `Set column ${columnIndexToName(column)} width`);
  }

  function handleRowResize(row: number, height: number) {
    if (!ensureSheetStructureEditable()) {
      return;
    }
    dispatchCommand({
      type: "rows.resize",
      sheetId: activeSheet.id,
      rows: [row],
      height
    }, `Set row ${row + 1} height`);
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

    const result = dispatchCommand({
      type: "transaction",
      commands: plan.map((item) => ({
        type: "columns.resize" as const,
        sheetId: activeSheet.id,
        columns: [item.column],
        width: item.width
      }))
    }, `Auto-fit ${pluralize(plan.length, "column")}`);
    if (result.status === "committed" && !result.changed) {
      setStatus(`${pluralize(plan.length, "column")} already fit`);
    }
  }

  function handleAutoFitColumn(column: number) {
    if (!ensureSheetStructureEditable()) {
      return;
    }

    const columnRange = { start: { row: 0, column }, end: { row: activeSheet.rowCount - 1, column } };
    const plan = createAutoFitColumnPlan(activeSheet, columnRange, (address) => getVisibleCellText(address));
    const result = dispatchCommand({
      type: "transaction",
      commands: plan.map((item) => ({
        type: "columns.resize" as const,
        sheetId: activeSheet.id,
        columns: [item.column],
        width: item.width
      }))
    }, `Auto-fit column ${columnIndexToName(column)}`);
    if (result.status === "committed" && !result.changed) {
      setStatus(`Column ${columnIndexToName(column)} already fits`);
    }
  }

  function handleAutoFitRow(row: number) {
    if (!ensureSheetStructureEditable()) {
      return;
    }

    const rowRangeTarget = { start: { row, column: 0 }, end: { row, column: activeSheet.columnCount - 1 } };
    const plan = createAutoFitRowPlan(activeSheet, rowRangeTarget, (address) => getVisibleCellText(address));
    const result = dispatchCommand({
      type: "transaction",
      commands: plan.map((item) => ({
        type: "rows.resize" as const,
        sheetId: activeSheet.id,
        rows: [item.row],
        height: item.height
      }))
    }, `Auto-fit row ${row + 1}`);
    if (result.status === "committed" && !result.changed) {
      setStatus(`Row ${row + 1} already fits`);
    }
  }

  function commitMoveFrom(address: string, move: CommitEditMove) {
    const origin = parseCellAddress(address);
    const delta =
      move === "down"
        ? { row: 1, column: 0 }
        : move === "up"
        ? { row: -1, column: 0 }
        : move === "right"
        ? { row: 0, column: 1 }
        : { row: 0, column: -1 };
    const target = stepPastHidden(activeSheet, origin, delta, isRowHiddenAt);
    setSelection({ start: target, end: target });
    gridApiRef.current?.ensureCellVisible(target.row, target.column);
    gridScrollRef.current?.focus({ preventScroll: true });
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

    const result = dispatchCommand({
      type: "transaction",
      commands: plan.map((item) => ({
        type: "rows.resize" as const,
        sheetId: activeSheet.id,
        rows: [item.row],
        height: item.height
      }))
    }, `Auto-fit ${pluralize(plan.length, "row")}`);
    if (result.status === "committed" && !result.changed) {
      setStatus(`${pluralize(plan.length, "row")} already fit`);
    }
  }

  function handleSort(direction: "asc" | "desc") {
    dispatchCommand({
      type: "range.sort",
      sheetId: activeSheet.id,
      range: selection,
      direction
    }, direction === "asc" ? "Sorted A to Z" : "Sorted Z to A");
  }

  function handleRemoveDuplicates() {
    const removedCount = countDuplicateRows(activeSheet, selection);
    if (removedCount === 0) {
      setStatus("No duplicate rows found");
      return;
    }
    dispatchCommand({
      type: "range.removeDuplicates",
      sheetId: activeSheet.id,
      range: selection
    }, `Removed ${removedCount} ${removedCount === 1 ? "duplicate row" : "duplicate rows"}`);
  }

  function handleInsertRows() {
    if (!ensureSheetStructureEditable()) {
      return;
    }
    const normalized = normalizeRange(selection);
    const count = normalized.end.row - normalized.start.row + 1;
    dispatchCommand({
      type: "transaction",
      commands: [
        { type: "rows.insert", sheetId: activeSheet.id, index: normalized.start.row, count },
        { type: "selection.set", selection: {
      start: normalized.start,
      end: { row: normalized.start.row + count - 1, column: normalized.end.column }
        } }
      ]
    }, `Inserted ${pluralize(count, "row")}`);
  }

  function handleDeleteRows() {
    if (!ensureSheetStructureEditable()) {
      return;
    }
    const normalized = normalizeRange(selection);
    const count = normalized.end.row - normalized.start.row + 1;
    const nextRow = clamp(normalized.start.row, 0, Math.max(activeSheet.rowCount - count - 1, 0));
    dispatchCommand({
      type: "transaction",
      commands: [
        { type: "rows.delete", sheetId: activeSheet.id, index: normalized.start.row, count },
        { type: "selection.set", selection: {
          start: { row: nextRow, column: normalized.start.column },
          end: { row: nextRow, column: normalized.end.column }
        } }
      ]
    }, `Deleted ${pluralize(count, "row")}`);
  }

  function handleInsertColumns() {
    if (!ensureSheetStructureEditable()) {
      return;
    }
    const normalized = normalizeRange(selection);
    const count = normalized.end.column - normalized.start.column + 1;
    dispatchCommand({
      type: "transaction",
      commands: [
        { type: "columns.insert", sheetId: activeSheet.id, index: normalized.start.column, count },
        { type: "selection.set", selection: {
          start: normalized.start,
          end: { row: normalized.end.row, column: normalized.start.column + count - 1 }
        } }
      ]
    }, `Inserted ${pluralize(count, "column")}`);
  }

  function handleDeleteColumns() {
    if (!ensureSheetStructureEditable()) {
      return;
    }
    const normalized = normalizeRange(selection);
    const count = normalized.end.column - normalized.start.column + 1;
    const nextColumn = clamp(normalized.start.column, 0, Math.max(activeSheet.columnCount - count - 1, 0));
    dispatchCommand({
      type: "transaction",
      commands: [
        { type: "columns.delete", sheetId: activeSheet.id, index: normalized.start.column, count },
        { type: "selection.set", selection: {
          start: { row: normalized.start.row, column: nextColumn },
          end: { row: normalized.end.row, column: nextColumn }
        } }
      ]
    }, `Deleted ${pluralize(count, "column")}`);
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
    const nextHiddenRows = { ...(activeSheet.hiddenRows ?? {}) };
    for (let row = normalized.start.row; row <= normalized.end.row; row += 1) nextHiddenRows[String(row)] = true;
    const nextRow = findNearestVisibleIndex(activeSheet.rowCount, nextHiddenRows, normalized.end.row + 1);
    const nextColumn = findNearestVisibleIndex(activeSheet.columnCount, activeSheet.hiddenColumns ?? {}, normalized.start.column);
    dispatchCommand({
      type: "transaction",
      commands: [
        {
          type: "rows.hidden.set",
          sheetId: activeSheet.id,
          rows: inclusiveIndexes(normalized.start.row, normalized.end.row),
          hidden: true
        },
        { type: "selection.set", selection: {
          start: { row: nextRow, column: nextColumn },
          end: { row: nextRow, column: nextColumn }
        } }
      ]
    }, `Hid ${pluralize(count, "row")}`);
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
    const nextHiddenColumns = { ...(activeSheet.hiddenColumns ?? {}) };
    for (let column = normalized.start.column; column <= normalized.end.column; column += 1) nextHiddenColumns[String(column)] = true;
    const nextRow = findNearestVisibleIndex(activeSheet.rowCount, activeSheet.hiddenRows ?? {}, normalized.start.row);
    const nextColumn = findNearestVisibleIndex(activeSheet.columnCount, nextHiddenColumns, normalized.end.column + 1);
    dispatchCommand({
      type: "transaction",
      commands: [
        {
          type: "columns.hidden.set",
          sheetId: activeSheet.id,
          columns: inclusiveIndexes(normalized.start.column, normalized.end.column),
          hidden: true
        },
        { type: "selection.set", selection: {
          start: { row: nextRow, column: nextColumn },
          end: { row: nextRow, column: nextColumn }
        } }
      ]
    }, `Hid ${pluralize(count, "column")}`);
  }

  function handleUnhideAll() {
    if (!ensureSheetStructureEditable()) {
      return;
    }

    const rows = flaggedIndexes(activeSheet.hiddenRows ?? {});
    const columns = flaggedIndexes(activeSheet.hiddenColumns ?? {});
    if (rows.length === 0 && columns.length === 0) {
      setStatus("No hidden rows or columns");
      return;
    }
    dispatchCommand({
      type: "transaction",
      commands: [
        ...(rows.length > 0 ? [{
          type: "rows.hidden.set" as const,
          sheetId: activeSheet.id,
          rows,
          hidden: false
        }] : []),
        ...(columns.length > 0 ? [{
          type: "columns.hidden.set" as const,
          sheetId: activeSheet.id,
          columns,
          hidden: false
        }] : [])
      ]
    }, "Unhid rows and columns");
  }

  function handleClearSelection() {
    dispatchCommand({
      type: "range.clear",
      sheetId: activeSheet.id,
      range: selection,
      mode: "contents"
    }, "Cleared selection");
  }

  function handleClearAll() {
    dispatchCommand({
      type: "range.clear",
      sheetId: activeSheet.id,
      range: selection,
      mode: "all"
    }, "Cleared all");
  }

  function handleClearComments() {
    dispatchCommand({
      type: "range.clear",
      sheetId: activeSheet.id,
      range: selection,
      mode: "comments"
    }, "Cleared comments");
  }

  function handleClearFormats() {
    dispatchCommand({
      type: "range.clear",
      sheetId: activeSheet.id,
      range: selection,
      mode: "formats"
    }, "Cleared formats");
  }

  function handleClearHyperlinks() {
    dispatchCommand({
      type: "range.clear",
      sheetId: activeSheet.id,
      range: selection,
      mode: "hyperlinks"
    }, "Cleared hyperlinks");
  }

  function openCellContextMenu(event: { address: string; row: number; column: number; x: number; y: number }) {
    // Excel keeps a multi-cell selection when right-clicking inside it, so the
    // menu can act on the whole range; clicking outside collapses to that cell.
    const normalized = normalizeRange(selection);
    const isInsideSelection =
      event.row >= normalized.start.row &&
      event.row <= normalized.end.row &&
      event.column >= normalized.start.column &&
      event.column <= normalized.end.column;
    if (!isInsideSelection) {
      setSelection({ start: { row: event.row, column: event.column }, end: { row: event.row, column: event.column } });
    }
    setEditingCell(null);
    setCellContextMenu({ address: event.address, x: event.x, y: event.y });
  }

  function handleFillDown() {
    dispatchCommand({
      type: "range.fill",
      sheetId: activeSheet.id,
      range: selection,
      direction: "down"
    }, "Filled down");
  }

  function handleFillRight() {
    dispatchCommand({
      type: "range.fill",
      sheetId: activeSheet.id,
      range: selection,
      direction: "right"
    }, "Filled right");
  }

  function handleAutoFill(sourceRange: CellRange, targetRange: CellRange) {
    const source = normalizeRange(sourceRange);
    const target = normalizeRange(targetRange);

    // Dragging the fill handle back inside the source shrinks the range: the
    // cells left behind are cleared, matching Excel.
    const isShrink =
      target.start.row === source.start.row &&
      target.start.column === source.start.column &&
      target.end.row <= source.end.row &&
      target.end.column <= source.end.column &&
      (target.end.row < source.end.row || target.end.column < source.end.column);
    if (isShrink) {
      const clearTarget =
        target.end.row < source.end.row
          ? { start: { row: target.end.row + 1, column: source.start.column }, end: source.end }
          : { start: { row: source.start.row, column: target.end.column + 1 }, end: source.end };
      dispatchCommand({
        type: "transaction",
        commands: [
          { type: "range.clear", sheetId: activeSheet.id, range: clearTarget, mode: "contents" },
          { type: "selection.set", selection: target }
        ]
      }, `Cleared ${formatSelectionAddress(clearTarget)}`);
      return;
    }

    const writeRange = getAutoFillWriteRange(source, target);
    if (!writeRange) {
      setStatus("Drag the fill handle along one direction");
      return;
    }
    const targetLabel = formatSelectionAddress(target);
    const result = dispatchCommand({
      type: "transaction",
      commands: [
        { type: "range.autoFill", sheetId: activeSheet.id, source, target },
        { type: "selection.set", selection: target }
      ]
    }, `AutoFilled ${targetLabel}`);
    if (result.status === "committed" && !result.changed) {
      setStatus(`AutoFilled ${targetLabel}`);
    }
  }

  function handleAutoFillDoubleClick() {
    // Excel: double-clicking the fill handle fills down as far as the adjacent
    // columns' contiguous data extends.
    const source = normalizeRange(selection);
    let extent = source.end.row;
    for (const column of [source.start.column - 1, source.end.column + 1]) {
      if (column < 0 || column >= activeSheet.columnCount) {
        continue;
      }
      let row = source.end.row + 1;
      while (row < activeSheet.rowCount && !isBlankCell(workbook, activeSheet.id, row, column)) {
        row += 1;
      }
      extent = Math.max(extent, row - 1);
    }

    if (extent <= source.end.row) {
      setStatus("No adjacent data to fill down to");
      return;
    }

    handleAutoFill(source, { start: source.start, end: { row: extent, column: source.end.column } });
  }

  function handleLockCells() {
    dispatchCommand({
      type: "range.readOnly.set",
      sheetId: activeSheet.id,
      range: selection,
      readOnly: true
    }, `Locked ${formatSelectionAddress(selection)}`);
  }

  function handleUnlockCells() {
    dispatchCommand({
      type: "range.readOnly.set",
      sheetId: activeSheet.id,
      range: selection,
      readOnly: false
    }, `Unlocked ${formatSelectionAddress(selection)}`);
  }

  function handleToggleProtection() {
    const isProtected = Boolean(activeSheet.protection?.isProtected);
    dispatchCommand({
      type: "sheet.protection.set",
      sheetId: activeSheet.id,
      protected: !isProtected
    }, isProtected ? "Unprotected sheet" : "Protected sheet");
  }

  function showPivotDrillDown(address: string): boolean {
    const meta = pivotDrillDowns[activeSheet.id];
    if (!meta) {
      return false;
    }

    const coord = parseCellAddress(address);
    const indexes = meta.drillDown[coord.row]?.[coord.column];
    if (!indexes || indexes.length === 0) {
      return false;
    }

    // Staleness guard: the sheet id may now belong to a different sheet (ids are
    // reused; undo can also rewind past the pivot's creation). Only drill down if
    // the cell still holds the value this pivot produced.
    const expected = meta.table[coord.row]?.[coord.column] ?? "";
    const actual = getCellContent(workbook, activeSheet.id, address);
    if (String(actual ?? "") !== expected) {
      return false;
    }

    const rows = [meta.sourceRows[0], ...indexes.map((index) => meta.sourceRows[index])];
    const detailsName = nextDetailsSheetName(workbook);
    const detailsSheetId = nextGeneratedSheetId(workbook);
    const result = dispatchCommand({
      type: "transaction",
      commands: [
        {
          type: "sheet.createFromMatrix",
          sheetId: detailsSheetId,
          name: detailsName,
          rows,
          formats: [{
            range: rowRange(0, rows[0].length),
            format: { bold: true, backgroundColor: "#eaf7f2" }
          }]
        },
        { type: "selection.set", selection: INITIAL_SELECTION }
      ]
    }, `Showing ${indexes.length} source ${indexes.length === 1 ? "row" : "rows"} in ${detailsName}`);
    if (result.status !== "committed") {
      return false;
    }
    setRichClipboard(null);
    setFormatPainter(null);
    return true;
  }

  function handleCreatePivotTable(config: PivotConfig) {
    try {
      const sourceRows = selectedRangeToDisplayRows(activeSheet, selection, formulaEngine);
      const { table: pivotRows, drillDown } = createPivotTableWithDetails(sourceRows, config);
      const pivotName = nextPivotSheetName(workbook);
      const pivotSheetId = nextGeneratedSheetId(workbook);
      const formats: Array<{ range: CellRange; format: Partial<CellFormat> }> = [
        {
          range: rowRange(0, pivotRows[0].length),
          format: {
            bold: true,
            textColor: "#17634a",
            backgroundColor: "#eaf7f2",
            borders: { bottom: { style: "thin", color: "#17634a" } }
          }
        },
        {
          range: rowRange(pivotRows.length - 1, pivotRows[0].length),
          format: {
            bold: true,
            backgroundColor: "#f1f5f8",
            borders: { top: { style: "thin", color: "#94a3b8" } }
          }
        }
      ];
      // Value columns read as numbers: right-align them like Excel's pivot output.
      const valueColumnStart = config.rowFields.length;
      const valueColumnEnd = pivotRows[0].length - 1;
      if (valueColumnEnd >= valueColumnStart && pivotRows.length > 1) {
        formats.push({
          range: {
            start: { row: 1, column: valueColumnStart },
            end: { row: pivotRows.length - 1, column: valueColumnEnd }
          },
          format: { horizontalAlign: "right" }
        });
      }
      // Size each pivot column to its widest cell so nothing arrives truncated.
      // Incremental max: spreading all rows into Math.max overflows the call
      // stack on very large pivot outputs.
      const columnWidths: number[] = [];
      for (let column = 0; column < pivotRows[0].length; column += 1) {
        let width = 0;
        for (const tableRow of pivotRows) {
          width = Math.max(width, textToAutoFitColumnWidth(String(tableRow[column] ?? "")));
        }
        columnWidths.push(width);
      }
      const result = dispatchCommand({
        type: "transaction",
        commands: [
          {
            type: "sheet.createFromMatrix",
            sheetId: pivotSheetId,
            name: pivotName,
            rows: pivotRows,
            formats,
            columnWidths,
            freeze: { rows: 1, columns: 0 }
          },
          { type: "selection.set", selection: INITIAL_SELECTION }
        ]
      }, `Created ${pivotName}`);
      if (result.status === "committed") {
        setPivotDrillDowns((current) => ({
          ...current,
          [pivotSheetId]: { sourceRows, drillDown, table: pivotRows }
        }));
        setRichClipboard(null);
        setFormatPainter(null);
        setPivotPanelOpen(false);
      }
    } catch (error) {
      setStatus(error instanceof Error ? error.message : "Could not create pivot table");
    }
  }

  function pasteRichClipboardMode(mode: RichPasteMode, nextStatus: string): boolean {
    if (!richClipboard) {
      setStatus("Copy cells before using paste special");
      return false;
    }
    const isMovePaste = mode === "all" && richClipboard.operation === "cut";
    const result = isMovePaste
      ? dispatchCommand({
          type: "clipboard.move",
          sourceSheetId: richClipboard.sourceSheetId,
          source: richClipboard.range.range,
          targetSheetId: activeSheet.id,
          target: parseCellAddress(activeAddress)
        }, "Moved selection")
      : dispatchCommand({
          type: "clipboard.paste",
          sheetId: activeSheet.id,
          target: parseCellAddress(activeAddress),
          payload: richClipboard.range,
          mode
        }, nextStatus);

    if (isMovePaste && result.status === "committed") {
      setRichClipboard(null);
    }
    return result.status === "committed";
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
    const matrix = normalizedText
      .split("\n")
      .filter((line, index, lines) => line.length > 0 || index < lines.length - 1)
      .map((line) => line.split("\t"));

    if (matrix.length === 0) {
      return;
    }

    const result = dispatchCommand({
      type: "clipboard.pasteMatrix",
      sheetId: activeSheet.id,
      target: parseCellAddress(activeAddress),
      matrix
    }, "Pasted cells");
    if (result.status === "committed") {
      setRichClipboard(null);
    }
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
    // Keep the toolbar keyboard-operable: Enter/Space on a focused control must
    // activate it, not start a cell edit. Everything else (arrows, shortcuts)
    // stays global so grid navigation works right after clicking a button.
    if (
      (event.key === "Enter" || event.key === " ") &&
      event.target instanceof Element &&
      event.target.closest("button, select, a, [role='tab']")
    ) {
      return;
    }
    handleKeyCommand(event);
  }

  function handleKeyCommand(event: React.KeyboardEvent<HTMLElement>) {
    // Keys already consumed by grid-internal widgets (cell editor, validation
    // dropdowns, AutoFilter menus) must not double-trigger grid commands.
    if (editingCell || event.defaultPrevented || isEditableEventTarget(event.target)) {
      return;
    }

    const isCommand = event.metaKey || event.ctrlKey;
    const isGridEvent = Boolean(gridScrollRef.current?.contains(event.target as Node));
    if (event.key === "Escape" && formatPainter) {
      event.preventDefault();
      setFormatPainter(null);
      setStatus("Format painter canceled");
      return;
    }

    if (event.key === "Escape" && richClipboard) {
      event.preventDefault();
      setRichClipboard(null);
      setStatus("Copy canceled");
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
      // Same mixed-selection semantics as the toolbar button: bold-all unless
      // every selected cell is already bold.
      applyFormat({ bold: selectionFormat.bold !== true }, "Applied bold");
      return;
    }

    if (isCommand && !event.shiftKey && event.key.toLowerCase() === "i") {
      event.preventDefault();
      applyFormat({ italic: selectionFormat.italic !== true }, "Applied italic");
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
      dispatchCommand(
        { type: event.shiftKey ? "history.redo" : "history.undo" },
        event.shiftKey ? "Redone" : "Undone"
      );
      return;
    }

    if (isCommand && event.key.toLowerCase() === "y") {
      event.preventDefault();
      dispatchCommand({ type: "history.redo" }, "Redone");
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

    if (isCommand && event.key.toLowerCase() === "a" && isGridEvent) {
      event.preventDefault();
      selectCurrentRegionThenSheet();
      return;
    }

    if (event.key === "Delete" || event.key === "Backspace") {
      event.preventDefault();
      handleClearSelection();
      return;
    }

    if (event.key === "Enter" || event.key === "F2") {
      event.preventDefault();
      if (!ensureEditableAddress(activeAddress)) {
        return;
      }
      setEditingCell({ address: activeAddress, value: formulaDraft });
      return;
    }

    if (event.key === "Tab" && !isCommand && isGridEvent) {
      event.preventDefault();
      const target = stepPastHidden(activeSheet, selection.start, { row: 0, column: event.shiftKey ? -1 : 1 });
      moveActiveCellTo(target);
      return;
    }

    if (event.key === "Home" && isGridEvent) {
      event.preventDefault();
      const baseRow = isCommand ? 0 : selection.start.row;
      // Home must not land on hidden rows/columns any more than arrows may.
      const hiddenColumns = activeSheet.hiddenColumns ?? {};
      const column = hiddenColumns["0"]
        ? stepPastHidden(activeSheet, { row: baseRow, column: 0 }, { row: 0, column: 1 }).column
        : 0;
      moveActiveCellTo(snapRowVisible({ row: baseRow, column }, 1));
      return;
    }

    if (event.key === "End" && isCommand && isGridEvent) {
      event.preventDefault();
      moveActiveCellTo(snapRowVisible(getUsedBounds(activeSheet), -1));
      return;
    }

    if ((event.key === "PageDown" || event.key === "PageUp") && isGridEvent) {
      event.preventDefault();
      const zoomFactor = zoomLevel / 100;
      const viewportHeight = (gridScrollRef.current?.clientHeight ?? 560) / zoomFactor;
      const pageRows = Math.max(1, Math.floor((viewportHeight - 28) / 28));
      const rowDelta = event.key === "PageDown" ? pageRows : -pageRows;
      if (event.shiftKey) {
        const focus = snapRowVisible(
          { row: clamp(selection.end.row + rowDelta, 0, activeSheet.rowCount - 1), column: selection.end.column },
          rowDelta > 0 ? 1 : -1
        );
        setSelection({ start: selection.start, end: focus });
        gridApiRef.current?.ensureCellVisible(focus.row, focus.column);
        return;
      }
      moveActiveCellTo(
        snapRowVisible(
          { row: clamp(selection.start.row + rowDelta, 0, activeSheet.rowCount - 1), column: selection.start.column },
          rowDelta > 0 ? 1 : -1
        )
      );
      return;
    }

    if (isCommand && event.key === " " && isGridEvent) {
      event.preventDefault();
      const normalized = normalizeRange(selection);
      setSelection({
        start: { row: 0, column: normalized.start.column },
        end: { row: activeSheet.rowCount - 1, column: normalized.end.column }
      });
      return;
    }

    if (!isCommand && event.shiftKey && event.key === " " && isGridEvent) {
      event.preventDefault();
      const normalized = normalizeRange(selection);
      setSelection({
        start: { row: normalized.start.row, column: 0 },
        end: { row: normalized.end.row, column: activeSheet.columnCount - 1 }
      });
      return;
    }

    if (event.key === "ArrowDown" || event.key === "ArrowUp" || event.key === "ArrowLeft" || event.key === "ArrowRight") {
      event.preventDefault();
      const delta = {
        ArrowDown: { row: 1, column: 0 },
        ArrowUp: { row: -1, column: 0 },
        ArrowLeft: { row: 0, column: -1 },
        ArrowRight: { row: 0, column: 1 }
      }[event.key]!;

      if (isCommand) {
        // Ctrl+Arrow jumps to the data-region edge; with Shift the selection
        // extends from the anchor instead of collapsing.
        const origin = event.shiftKey ? selection.end : selection.start;
        const jumped = jumpToDataEdge(workbook, activeSheet, origin, delta);
        const target = delta.row !== 0 ? snapRowVisible(jumped, delta.row > 0 ? 1 : -1) : jumped;
        if (event.shiftKey) {
          setSelection({ start: selection.start, end: target });
        } else {
          setSelection({ start: target, end: target });
        }
        gridApiRef.current?.ensureCellVisible(target.row, target.column);
        return;
      }

      if (event.shiftKey) {
        const target = stepPastHidden(activeSheet, selection.end, delta, isRowHiddenAt);
        setSelection({ start: selection.start, end: target });
        gridApiRef.current?.ensureCellVisible(target.row, target.column);
        return;
      }

      moveActiveCellTo(stepPastHidden(activeSheet, selection.start, delta, isRowHiddenAt));
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

  function moveActiveCellTo(target: CellCoord) {
    setSelection({ start: target, end: target });
    gridApiRef.current?.ensureCellVisible(target.row, target.column);
  }

  // A row is hidden when explicitly hidden on the sheet OR excluded by an
  // active filter — keyboard navigation must skip both, or the active cell
  // lands somewhere invisible.
  function isRowHiddenAt(row: number): boolean {
    if ((activeSheet.hiddenRows ?? {})[String(row)]) {
      return true;
    }
    const filters = activeSheet.filters ?? [];
    if (filters.length === 0) {
      return false;
    }
    return !filters.every((filter) =>
      isRowVisibleForFilter(row, filter, (filterRow, filterColumn) =>
        formulaEngine.getComputedValue(activeSheet.id, formatCellAddress({ row: filterRow, column: filterColumn }))
      )
    );
  }

  // Walks a landing target off hidden rows in the movement direction, falling
  // back the other way at the sheet edge.
  function snapRowVisible(target: CellCoord, direction: 1 | -1): CellCoord {
    let row = target.row;
    while (row >= 0 && row < activeSheet.rowCount && isRowHiddenAt(row)) {
      row += direction;
    }
    if (row < 0 || row >= activeSheet.rowCount) {
      row = target.row;
      while (row >= 0 && row < activeSheet.rowCount && isRowHiddenAt(row)) {
        row -= direction;
      }
      if (row < 0 || row >= activeSheet.rowCount) {
        return target;
      }
    }
    return { row, column: target.column };
  }

  function selectCurrentRegionThenSheet() {
    const wholeSheet = {
      start: { row: 0, column: 0 },
      end: { row: activeSheet.rowCount - 1, column: activeSheet.columnCount - 1 }
    };
    const region = normalizeRange(expandDataRegion(workbook, activeSheet, selection.start));
    const normalized = normalizeRange(selection);
    const regionIsSingleCell = region.start.row === region.end.row && region.start.column === region.end.column;
    const alreadyRegionSelected =
      normalized.start.row === region.start.row &&
      normalized.start.column === region.start.column &&
      normalized.end.row === region.end.row &&
      normalized.end.column === region.end.column;
    setSelection(regionIsSingleCell || alreadyRegionSelected ? wholeSheet : region);
  }

  function handleAddSheet() {
    const result = dispatchCommand({
      type: "transaction",
      commands: [
        { type: "sheet.add" },
        { type: "selection.set", selection: INITIAL_SELECTION }
      ]
    }, "Added sheet");
    if (result.status === "committed") {
      setRichClipboard(null);
      setFormatPainter(null);
      setPivotPanelOpen(false);
    }
  }

  function handleRenameSheet() {
    const nextName = window.prompt("Rename active sheet", activeSheet.name);
    if (nextName) {
      dispatchCommand({
        type: "sheet.rename",
        sheetId: activeSheet.id,
        name: nextName
      }, "Renamed sheet");
    }
  }

  function handleDuplicateSheet() {
    const result = dispatchCommand({
      type: "transaction",
      commands: [
        { type: "sheet.duplicate", sheetId: activeSheet.id },
        { type: "selection.set", selection: INITIAL_SELECTION }
      ]
    }, "Duplicated sheet");
    if (result.status === "committed") {
      setRichClipboard(null);
      setFormatPainter(null);
      setPivotPanelOpen(false);
    }
  }

  function handleDeleteSheet() {
    const deletedSheetId = activeSheet.id;
    const result = dispatchCommand({
      type: "transaction",
      commands: [
        { type: "sheet.delete", sheetId: deletedSheetId },
        { type: "selection.set", selection: INITIAL_SELECTION }
      ]
    }, "Deleted sheet");
    if (result.status === "committed" && result.changed) {
      setPivotDrillDowns((current) => {
        if (!(deletedSheetId in current)) {
          return current;
        }
        const { [deletedSheetId]: removed, ...rest } = current;
        void removed;
        return rest;
      });
      setRichClipboard(null);
      setFormatPainter(null);
      setPivotPanelOpen(false);
    }
  }

  function handleHideSheet() {
    const hiddenSheetName = activeSheet.name;
    const result = dispatchCommand({
      type: "transaction",
      commands: [
        { type: "sheet.hidden.set", sheetId: activeSheet.id, hidden: true },
        { type: "selection.set", selection: INITIAL_SELECTION }
      ]
    }, `Hid ${hiddenSheetName}`);
    if (result.status === "committed" && !result.changed) {
      setStatus("Cannot hide the only visible sheet");
      return;
    }
    if (result.status === "committed") {
      setRichClipboard(null);
      setFormatPainter(null);
      setPivotPanelOpen(false);
    }
  }

  function handleUnhideSheets() {
    const hiddenSheetIds = session.getSnapshot().workbook.sheets
      .filter((sheet) => sheet.isHidden === true)
      .map((sheet) => sheet.id);
    if (hiddenSheetIds.length === 0) {
      setStatus("No hidden sheets");
      return;
    }
    dispatchCommand({
      type: "transaction",
      commands: hiddenSheetIds.map((sheetId) => ({
        type: "sheet.hidden.set" as const,
        sheetId,
        hidden: false
      }))
    }, "Unhid sheets");
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
    dispatchCommand({
      type: "sheet.move",
      sheetId: activeSheet.id,
      targetIndex
    }, `Moved ${activeSheet.name} ${direction}`);
  }

  function handleSheetTabColor(color: string) {
    dispatchCommand({
      type: "sheet.tabColor.set",
      sheetId: activeSheet.id,
      color
    }, `Changed ${activeSheet.name} tab color`);
  }

  function handleNewWorkbook() {
    const next = createBlankWorkbook();
    const result = session.replaceWorkbook(next, { history: "reset", origin: "external" });
    if (result.status === "committed") {
      resetAfterWorkbookReplacement("New workbook");
    }
  }

  function handleImportCsv(file: File | undefined) {
    if (!file) {
      return;
    }

    const reader = new FileReader();
    reader.onload = () => {
      const text = String(reader.result ?? "");
      const importer = services?.importers?.csv;
      Promise.resolve().then(async () => {
        if (importer) {
          const nextWorkbook = await importer.import({ kind: "text", text, fileName: file.name });
          const result = session.replaceWorkbook(nextWorkbook, { history: "preserve", origin: "import" });
          if (result.status === "committed") {
            resetAfterWorkbookReplacement(`Imported ${file.name}`);
          }
          return;
        }

        const snapshot = session.getSnapshot();
        const result = dispatchCommand({
          type: "transaction",
          commands: [
            {
              type: "sheet.replaceWithRows",
              sheetId: snapshot.workbook.activeSheetId,
              rows: parseCsv(text)
            },
            { type: "selection.set", selection: INITIAL_SELECTION }
          ]
        }, `Imported ${file.name}`);
        if (result.status === "committed") {
          clearWorkbookReplacementUiState();
        }
      })
        .catch(() => reportServiceFailure("service.import.csv.failed", "CSV import failed"));
    };
    reader.readAsText(file);
  }

  function handleExportCsv() {
    const exporter = services?.exporters?.csv;
    if (exporter) {
      Promise.resolve().then(() => exporter.export(workbook)).then((artifact) => {
        downloadWorkbookArtifact(artifact);
        setStatus("Exported CSV");
      }).catch(() => reportServiceFailure("service.export.csv.failed", "CSV export failed"));
      return;
    }
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
        const importer = services?.importers?.xlsx;
        const nextWorkbook = importer
          ? await importer.import({
              kind: "bytes",
              bytes: new Uint8Array(buffer),
              fileName: file.name
            })
          : await importWorkbookFromXlsx(buffer);
        const result = session.replaceWorkbook(nextWorkbook, { history: "preserve", origin: "import" });
        if (result.status === "committed") {
          resetAfterWorkbookReplacement(`Imported ${file.name}`);
        }
      })
      .catch(() => reportServiceFailure("service.import.xlsx.failed", "XLSX import failed"));
  }

  function handleImportGoogleSheet() {
    const clientId = import.meta.env.VITE_GOOGLE_CLIENT_ID as string | undefined;
    if (!clientId) {
      setStatus("Set VITE_GOOGLE_CLIENT_ID (a Google OAuth client id) to link Google Sheets");
      return;
    }

    const input = window.prompt("Paste a Google Sheets URL (or spreadsheet id):");
    if (!input) {
      return;
    }

    let tokenProvider: TokenProvider;
    try {
      tokenProvider = getComponentGoogleTokenProvider(clientId);
    } catch {
      reportServiceFailure("service.google.auth.failed", "Google Sheets connection failed");
      return;
    }
    setStatus("Connecting to Google Sheets…");
    importWorkbookFromGoogleSheets(input, tokenProvider)
      .then(({ workbook: nextWorkbook, spreadsheetTitle }) => {
        const result = session.replaceWorkbook(nextWorkbook, { history: "preserve", origin: "import" });
        if (result.status === "committed") {
          resetAfterWorkbookReplacement(`Linked ${spreadsheetTitle}`);
        }
      })
      .catch(() => reportServiceFailure("service.google.import.failed", "Google Sheets import failed"));
  }

  function handleExportXlsx() {
    const exporter = services?.exporters?.xlsx;
    Promise.resolve().then(async () => exporter
      ? exporter.export(workbook)
      : {
          bytes: new Uint8Array(await exportWorkbookToXlsx(workbook)),
          mediaType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
          fileName: "spreadsheet.xlsx"
        })
      .then((artifact) => {
        downloadWorkbookArtifact(artifact);
        setStatus("Exported XLSX");
      })
      .catch(() => reportServiceFailure("service.export.xlsx.failed", "XLSX export failed"));
  }

  function reportServiceFailure(code: string, message: string) {
    setStatus(message);
    invokeHostCallback(onError, { code, message, recoverable: true });
  }

  function clearWorkbookReplacementUiState() {
    setPivotDrillDowns({});
    setRichClipboard(null);
    setFormatPainter(null);
    setPivotPanelOpen(false);
  }

  function resetAfterWorkbookReplacement(nextStatus: string) {
    dispatchCommand({ type: "selection.set", selection: INITIAL_SELECTION });
    clearWorkbookReplacementUiState();
    setStatus(nextStatus);
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
    gridApiRef.current?.ensureCellVisible(nextMatch.row, nextMatch.column);
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

    const currentWorkbook = session.getSnapshot().workbook;
    const currentSheet = getActiveSheet(currentWorkbook);
    const currentMatches = findMatches(currentSheet, findDraft);
    const commands: WorkbookCommand[] = [];
    let changedCells = 0;
    for (const match of currentMatches) {
      const nextContent = replaceEveryMatch(String(match.content), findDraft, replaceDraft);
      if (nextContent !== match.content) {
        commands.push({
          type: "cell.set",
          sheetId: currentSheet.id,
          address: match.address,
          input: nextContent
        });
        changedCells += 1;
      }
    }

    if (changedCells === 0) {
      setStatus(`No matches for ${findDraft}`);
      return;
    }

    dispatchCommand({ type: "transaction", commands }, `Replaced ${changedCells} ${changedCells === 1 ? "cell" : "cells"}`);
  }

  function handleFileDrop(event: React.DragEvent) {
    event.preventDefault();
    setDropTargetActive(false);
    if (features?.import === false) {
      return;
    }
    const file = event.dataTransfer.files?.[0];
    if (!file) {
      return;
    }

    const name = file.name.toLowerCase();
    if (name.endsWith(".xlsx")) {
      handleImportXlsx(file);
    } else if (name.endsWith(".csv")) {
      handleImportCsv(file);
    } else {
      setStatus("Drop an .xlsx or .csv file to import it");
    }
  }

  const themedStyle = {
    ...themeToRootStyle(theme),
    ...style
  } as CSSProperties;

  return (
    <div
      className={["js-spreadsheet-root", "js-spreadsheet-workbook", className]
        .filter(Boolean)
        .join(" ")}
      data-js-spreadsheet-root="workbook"
      style={themedStyle}
    >
      <main
        className="app-shell"
        onKeyDown={handleShellKeyCommand}
        onDragOver={(event) => {
          if (event.dataTransfer.types.includes("Files")) {
            event.preventDefault();
            setDropTargetActive(true);
          }
        }}
        onDragLeave={(event) => {
          if (!event.currentTarget.contains(event.relatedTarget as Node | null)) {
            setDropTargetActive(false);
          }
        }}
        onDrop={handleFileDrop}
      >
        {isDropTargetActive ? (
          <div className="file-drop-overlay" aria-hidden="true">
            Drop to import workbook
          </div>
        ) : null}
        <section className="spreadsheet-surface" aria-label="JavaScript spreadsheet">
        {features?.toolbar !== false ? (
        <Toolbar
          features={features}
          canUndo={sessionSnapshot.canUndo}
          canRedo={sessionSnapshot.canRedo}
          onNew={handleNewWorkbook}
          onImport={() => fileInputRef.current?.click()}
          onExport={handleExportCsv}
          onImportXlsx={() => xlsxInputRef.current?.click()}
          onImportGoogleSheet={handleImportGoogleSheet}
          onExportXlsx={handleExportXlsx}
          onPrint={handlePrintWorkbook}
          onUndo={() => {
            dispatchCommand({ type: "history.undo" }, "Undone");
          }}
          onRedo={() => {
            dispatchCommand({ type: "history.redo" }, "Redone");
          }}
          onClear={handleClearSelection}
          canPasteSpecial={Boolean(richClipboard)}
          onPaste={handlePasteAll}
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
            dispatchCommand({
              type: "sheet.freeze.set",
              sheetId: activeSheet.id,
              rows: nextFreezeTopRow ? 1 : 0,
              columns: freezeFirstColumn ? 1 : 0
            }, nextFreezeTopRow ? "Froze top row" : "Unfroze top row");
          }}
          onToggleFreezeFirstColumn={() => {
            const nextFreezeFirstColumn = !freezeFirstColumn;
            dispatchCommand({
              type: "sheet.freeze.set",
              sheetId: activeSheet.id,
              rows: freezeTopRow ? 1 : 0,
              columns: nextFreezeFirstColumn ? 1 : 0
            }, nextFreezeFirstColumn ? "Froze first column" : "Unfroze first column");
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
          selectionFormat={selectionFormat}
          findPanelOpen={isFindPanelOpen}
          filterPanelOpen={isFilterPanelOpen}
          validationPanelOpen={isValidationPanelOpen}
          conditionalPanelOpen={isConditionalPanelOpen}
          pivotPanelOpen={isPivotPanelOpen}
          chartPanelOpen={isChartPanelOpen}
          functionLibraryOpen={isFunctionLibraryOpen}
          namedRangesOpen={isNamedRangesOpen}
          goToPanelOpen={isGoToPanelOpen}
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
          onBold={() => applyFormat({ bold: selectionFormat.bold !== true }, "Applied bold")}
          onItalic={() => applyFormat({ italic: selectionFormat.italic !== true }, "Applied italic")}
          onWrapText={() =>
            applyFormat(
              { wrapText: selectionFormat.wrapText !== true },
              selectionFormat.wrapText === true ? "Unwrapped text" : "Wrapped text"
            )
          }
          onFontFamily={(fontFamily) =>
            applyFormat({ fontFamily: fontFamily || undefined }, fontFamily ? `Applied ${fontFamily}` : "Reset font")
          }
          onFontSize={(fontSize) =>
            applyFormat({ fontSize: fontSize ?? undefined }, fontSize ? `Applied ${fontSize}pt size` : "Reset font size")
          }
          onNumberFormat={(numberFormat) => applyFormat({ numberFormat }, `Applied ${numberFormat} format`)}
          onHorizontalAlign={(horizontalAlign) => applyFormat({ horizontalAlign }, `Aligned ${horizontalAlign}`)}
          onVerticalAlign={(verticalAlign) => applyFormat({ verticalAlign }, `Aligned ${verticalAlign}`)}
          onBorders={applyBorders}
          onTextColor={(color) => applyFormat({ textColor: color }, "Changed text color")}
          onFillColor={(color) => applyFormat({ backgroundColor: color }, "Changed fill color")}
        />
        ) : null}
        {features?.import !== false ? (
        <>
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
        </>
        ) : null}
        {showFormulaBar && features?.formulaBar !== false ? (
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
          sourceLabel={pivotSourceLabel}
          sourceRowCount={Math.max(0, pivotSourceRows.length - 1)}
          isOpen={isPivotPanelOpen}
          onClose={() => setPivotPanelOpen(false)}
          onCreate={handleCreatePivotTable}
        />
        {features?.charts !== false ? <ChartPanel
          isOpen={isChartPanelOpen}
          onClose={() => setChartPanelOpen(false)}
          onCreate={handleCreateChart}
        /> : null}
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
          copiedRange={richClipboard && richClipboard.sourceSheetId === activeSheet.id ? richClipboard.range.range : null}
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
            if (showPivotDrillDown(address)) {
              return;
            }
            if (!ensureEditableAddress(address)) {
              return;
            }
            setEditingCell({ address, value: String(getCellContent(workbook, activeSheet.id, address) ?? "") });
          }}
          onEditValueChange={(value) => setEditingCell((current) => (current ? { ...current, value } : current))}
          onCommitEdit={(address, value, move) => {
            if (gridEditCommitInProgressRef.current) {
              return;
            }
            gridEditCommitInProgressRef.current = true;
            try {
              if (commitCell(address, value)) {
                setEditingCell(null);
                if (move) {
                  commitMoveFrom(address, move);
                }
              }
            } finally {
              gridEditCommitInProgressRef.current = false;
            }
          }}
          onCancelEdit={() => {
            setEditingCell(null);
            // Escape came from the keyboard: hand focus back to the grid so
            // arrows/typing keep working instead of falling to document.body.
            gridScrollRef.current?.focus({ preventScroll: true });
          }}
          onPasteText={pasteText}
          onKeyCommand={handleKeyCommand}
          onAutoFill={handleAutoFill}
          onAutoFillDoubleClick={handleAutoFillDoubleClick}
          onCellContextMenu={openCellContextMenu}
          onAutoFilterColumn={applyAutoFilterColumn}
          onClearAutoFilterColumn={clearAutoFilterColumn}
          onSortAutoFilterColumn={sortAutoFilterColumn}
          onColumnResize={handleColumnResize}
          onRowResize={handleRowResize}
          onColumnAutoFit={handleAutoFitColumn}
          onRowAutoFit={handleAutoFitRow}
          onRegisterScrollApi={registerGridScrollApi}
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
        {features?.charts !== false ? <SheetCharts
          charts={activeSheet.charts ?? []}
          sheet={activeSheet}
          formulaEngine={formulaEngine}
          onDelete={handleDeleteChart}
        /> : null}
        {showSheetTabs && features?.sheetTabs !== false ? (
          <SheetTabs
            sheets={workbook.sheets}
            activeSheetId={activeSheet.id}
            onSelect={(sheetId) => {
              dispatchCommand({
                type: "transaction",
                commands: [
                  { type: "sheet.activate", sheetId },
                  { type: "selection.set", selection: INITIAL_SELECTION }
                ]
              });
            }}
            onAdd={handleAddSheet}
          />
        ) : null}
        <StatusBar
          status={status}
          activeAddress={activeAddress}
          selectedCount={countSelectedCells(selection)}
          selectionSummary={selectionSummary}
          formulaFunctions="HyperFormula 418+ functions"
          zoomLevel={zoomLevel}
          onZoomOut={() => handleZoom(-ZOOM_STEP)}
          onZoomIn={() => handleZoom(ZOOM_STEP)}
          onResetZoom={handleResetZoom}
        />
      </section>
      </main>
    </div>
  );
}

export default function App() {
  return <Spreadsheet />;
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

// Direction validity is delegated to the same predicates autoFillRange uses,
// so "what handleAutoFill validates" and "what the engine fills" cannot diverge.
function getAutoFillWriteRange(sourceRange: CellRange, targetRange: CellRange): CellRange | null {
  const source = normalizeRange(sourceRange);
  const target = normalizeRange(targetRange);

  if (isVerticalAutoFill(source, target)) {
    return target.end.row > source.end.row
      ? {
          start: { row: source.end.row + 1, column: source.start.column },
          end: { row: target.end.row, column: source.end.column }
        }
      : {
          start: { row: target.start.row, column: source.start.column },
          end: { row: source.start.row - 1, column: source.end.column }
        };
  }

  if (isHorizontalAutoFill(source, target)) {
    return target.end.column > source.end.column
      ? {
          start: { row: source.start.row, column: source.end.column + 1 },
          end: { row: source.end.row, column: target.end.column }
        }
      : {
          start: { row: source.start.row, column: target.start.column },
          end: { row: source.end.row, column: source.start.column - 1 }
        };
  }

  return null;
}

// Above this, the sparse-format scan per selection change is no longer cheap;
// fall back to anchor-only state (no mixed detection) instead of stalling.
const SELECTION_FORMAT_SCAN_LIMIT = 20_000;

// "general" is the absence of a number format: cells that store it explicitly
// (after picking General in the ribbon) are identical in effect to cells with
// no entry at all, and must not read as a mixed selection next to them.
function effectiveNumberFormat(format: CellFormat): CellFormat["numberFormat"] {
  return format.numberFormat === "general" ? undefined : format.numberFormat;
}

// Summarizes formats over the selection by scanning the sheet's SPARSE format
// map (O(populated formats)), never the selection area itself — whole-column
// selections cover 100k+ cells and must stay cheap.
function summarizeSelectionFormats(sheet: SheetModel, selection: CellRange): SelectionFormatSummary {
  const normalized = normalizeRange(selection);
  const cellCount =
    (normalized.end.row - normalized.start.row + 1) * (normalized.end.column - normalized.start.column + 1);
  const formats = sheet.formats ?? {};
  const entries = Object.entries(formats);

  // Single cells (the common case while arrow-key navigating) resolve with a
  // direct lookup instead of a scan.
  if (cellCount === 1 || entries.length > SELECTION_FORMAT_SCAN_LIMIT) {
    const anchor = formats[formatCellAddress(normalized.start)] ?? {};
    return {
      bold: Boolean(anchor.bold),
      italic: Boolean(anchor.italic),
      wrapText: Boolean(anchor.wrapText),
      fontFamily: anchor.fontFamily,
      fontSize: anchor.fontSize,
      numberFormat: effectiveNumberFormat(anchor),
      horizontalAlign: anchor.horizontalAlign,
      verticalAlign: anchor.verticalAlign
    };
  }

  const createTracker = <T,>() => ({
    value: undefined as T | undefined,
    defined: 0,
    conflicting: false,
    add(candidate: T | undefined) {
      if (candidate === undefined) {
        return;
      }
      if (this.defined === 0) {
        this.value = candidate;
      } else if (this.value !== candidate) {
        this.conflicting = true;
      }
      this.defined += 1;
    }
  });

  let boldCount = 0;
  let italicCount = 0;
  let wrapCount = 0;
  const fontFamily = createTracker<string>();
  const fontSize = createTracker<number>();
  const numberFormat = createTracker<NonNullable<CellFormat["numberFormat"]>>();
  const horizontalAlign = createTracker<NonNullable<CellFormat["horizontalAlign"]>>();
  const verticalAlign = createTracker<NonNullable<CellFormat["verticalAlign"]>>();

  for (const [address, format] of entries) {
    const coord = parseCellAddress(address);
    if (
      coord.row < normalized.start.row ||
      coord.row > normalized.end.row ||
      coord.column < normalized.start.column ||
      coord.column > normalized.end.column
    ) {
      continue;
    }
    if (format.bold) {
      boldCount += 1;
    }
    if (format.italic) {
      italicCount += 1;
    }
    if (format.wrapText) {
      wrapCount += 1;
    }
    fontFamily.add(format.fontFamily);
    fontSize.add(format.fontSize);
    numberFormat.add(effectiveNumberFormat(format));
    horizontalAlign.add(format.horizontalAlign);
    verticalAlign.add(format.verticalAlign);
  }

  const toggleState = (count: number): boolean | "mixed" =>
    count === cellCount ? true : count === 0 ? false : "mixed";
  const resolve = <T,>(tracker: { value: T | undefined; defined: number; conflicting: boolean }): MixedFormatValue<T> =>
    tracker.defined === 0 ? undefined : tracker.conflicting || tracker.defined < cellCount ? "mixed" : tracker.value;

  return {
    bold: toggleState(boldCount),
    italic: toggleState(italicCount),
    wrapText: toggleState(wrapCount),
    fontFamily: resolve(fontFamily),
    fontSize: resolve(fontSize),
    numberFormat: resolve(numberFormat),
    horizontalAlign: resolve(horizontalAlign),
    verticalAlign: resolve(verticalAlign)
  };
}

function isBlankCell(workbook: WorkbookModel, sheetId: string, row: number, column: number): boolean {
  const content = getCellContent(workbook, sheetId, formatCellAddress({ row, column }));
  return content === null || String(content).trim() === "";
}

// Excel Ctrl+Arrow: from inside a data run, jump to the run's edge; from an
// empty cell (or a run edge), jump to the next populated cell, else the sheet edge.
function jumpToDataEdge(
  workbook: WorkbookModel,
  sheet: SheetModel,
  coord: CellCoord,
  delta: { row: number; column: number }
): CellCoord {
  const maxRow = sheet.rowCount - 1;
  const maxColumn = sheet.columnCount - 1;
  const step = (from: CellCoord): CellCoord => ({
    row: Math.min(Math.max(from.row + delta.row, 0), maxRow),
    column: Math.min(Math.max(from.column + delta.column, 0), maxColumn)
  });
  const samePosition = (left: CellCoord, right: CellCoord) => left.row === right.row && left.column === right.column;
  const isFilled = (position: CellCoord) => !isBlankCell(workbook, sheet.id, position.row, position.column);

  let current = coord;
  const next = step(current);
  if (samePosition(next, current)) {
    return current;
  }

  if (isFilled(current) && isFilled(next)) {
    current = next;
    while (true) {
      const following = step(current);
      if (samePosition(following, current) || !isFilled(following)) {
        return current;
      }
      current = following;
    }
  }

  current = next;
  while (!isFilled(current)) {
    const following = step(current);
    if (samePosition(following, current)) {
      return current;
    }
    current = following;
  }
  return current;
}

// One arrow-key step that skips rows/columns hidden via the sheet model (and,
// through the optional predicate, filter-hidden rows), so the active cell
// never lands somewhere invisible. Stays put at the sheet edge.
function stepPastHidden(
  sheet: SheetModel,
  coord: CellCoord,
  delta: { row: number; column: number },
  isRowHidden?: (row: number) => boolean
): CellCoord {
  if (delta.row !== 0) {
    const hiddenRows = sheet.hiddenRows ?? {};
    const rowHidden = isRowHidden ?? ((row: number) => Boolean(hiddenRows[String(row)]));
    let next = coord.row + delta.row;
    while (next >= 0 && next < sheet.rowCount && rowHidden(next)) {
      next += delta.row;
    }
    if (next < 0 || next >= sheet.rowCount) {
      return coord;
    }
    return { row: next, column: coord.column };
  }

  if (delta.column !== 0) {
    const hiddenColumns = sheet.hiddenColumns ?? {};
    let next = coord.column + delta.column;
    while (next >= 0 && next < sheet.columnCount && hiddenColumns[String(next)]) {
      next += delta.column;
    }
    if (next < 0 || next >= sheet.columnCount) {
      return coord;
    }
    return { row: coord.row, column: next };
  }

  return coord;
}

// Excel's "current region": the contiguous block of data around the coordinate,
// grown until every neighboring row/column ring is empty. Builds occupancy
// indexes over the sparse cell map once so each ring probe is a binary search —
// per-cell content probes would be O(N²) on large regions.
function expandDataRegion(workbook: WorkbookModel, sheet: SheetModel, coord: CellCoord): CellRange {
  void workbook;
  const maxRow = sheet.rowCount - 1;
  const maxColumn = sheet.columnCount - 1;
  const rowToColumns = new Map<number, number[]>();
  const columnToRows = new Map<number, number[]>();
  for (const [address, content] of Object.entries(sheet.cells)) {
    if (content === null || String(content).trim() === "") {
      continue;
    }
    const cellCoord = parseCellAddress(address);
    const columns = rowToColumns.get(cellCoord.row);
    if (columns) {
      columns.push(cellCoord.column);
    } else {
      rowToColumns.set(cellCoord.row, [cellCoord.column]);
    }
    const rowsForColumn = columnToRows.get(cellCoord.column);
    if (rowsForColumn) {
      rowsForColumn.push(cellCoord.row);
    } else {
      columnToRows.set(cellCoord.column, [cellCoord.row]);
    }
  }
  for (const columns of rowToColumns.values()) {
    columns.sort((a, b) => a - b);
  }
  for (const rowsForColumn of columnToRows.values()) {
    rowsForColumn.sort((a, b) => a - b);
  }

  const hasValueInRange = (sorted: number[] | undefined, min: number, max: number): boolean => {
    if (!sorted || sorted.length === 0) {
      return false;
    }
    let low = 0;
    let high = sorted.length - 1;
    let index = sorted.length;
    while (low <= high) {
      const mid = (low + high) >> 1;
      if (sorted[mid] >= min) {
        index = mid;
        high = mid - 1;
      } else {
        low = mid + 1;
      }
    }
    return index < sorted.length && sorted[index] <= max;
  };

  let top = coord.row;
  let bottom = coord.row;
  let left = coord.column;
  let right = coord.column;
  let changed = true;
  while (changed) {
    changed = false;
    const ringLeft = Math.max(0, left - 1);
    const ringRight = Math.min(maxColumn, right + 1);
    const ringTop = Math.max(0, top - 1);
    const ringBottom = Math.min(maxRow, bottom + 1);

    if (top > 0 && hasValueInRange(rowToColumns.get(top - 1), ringLeft, ringRight)) {
      top -= 1;
      changed = true;
    }
    if (bottom < maxRow && hasValueInRange(rowToColumns.get(bottom + 1), ringLeft, ringRight)) {
      bottom += 1;
      changed = true;
    }
    if (left > 0 && hasValueInRange(columnToRows.get(left - 1), ringTop, ringBottom)) {
      left -= 1;
      changed = true;
    }
    if (right < maxColumn && hasValueInRange(columnToRows.get(right + 1), ringTop, ringBottom)) {
      right += 1;
      changed = true;
    }
  }

  return { start: { row: top, column: left }, end: { row: bottom, column: right } };
}

function getUsedBounds(sheet: SheetModel): CellCoord {
  let maxRow = 0;
  let maxColumn = 0;
  for (const address of Object.keys(sheet.cells)) {
    const coord = parseCellAddress(address);
    if (coord.row > maxRow) {
      maxRow = coord.row;
    }
    if (coord.column > maxColumn) {
      maxColumn = coord.column;
    }
  }
  return { row: maxRow, column: maxColumn };
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

  let minRow = Number.POSITIVE_INFINITY;
  let minColumn = Number.POSITIVE_INFINITY;
  let maxRow = Number.NEGATIVE_INFINITY;
  let maxColumn = Number.NEGATIVE_INFINITY;
  for (const address in sheet.cells) {
    const coord = parseCellAddress(address);
    minRow = Math.min(minRow, coord.row);
    minColumn = Math.min(minColumn, coord.column);
    maxRow = Math.max(maxRow, coord.row);
    maxColumn = Math.max(maxColumn, coord.column);
  }
  if (!Number.isFinite(minRow)) {
    return normalized;
  }

  return {
    start: { row: minRow, column: minColumn },
    end: { row: maxRow, column: maxColumn }
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

// Threshold check without materializing a key array — stops as soon as the
// target is reached, so it costs min(target, populated) iterations and no
// allocation even on million-cell sheets.
function hasAtLeastCellCount(cells: SheetModel["cells"], target: number): boolean {
  let count = 0;
  for (const address in cells) {
    void address;
    count += 1;
    if (count >= target) {
      return true;
    }
  }
  return false;
}

function countSelectedCells(selection: CellRange): number {
  const range = normalizeRange(selection);
  return (range.end.row - range.start.row + 1) * (range.end.column - range.start.column + 1);
}

function summarizeSelection(sheet: SheetModel, selection: CellRange, formulaEngine: FormulaEngine): string {
  const range = normalizeRange(selection);
  let nonEmptyCount = 0;
  let numericCount = 0;
  let sum = 0;
  let min = Number.POSITIVE_INFINITY;
  let max = Number.NEGATIVE_INFINITY;

  const accumulate = (address: string) => {
    const value = formulaEngine.getDisplayValue(sheet.id, address);
    if (value.trim() === "") {
      return;
    }
    nonEmptyCount += 1;

    const numeric = Number(value);
    if (Number.isFinite(numeric)) {
      numericCount += 1;
      sum += numeric;
      min = Math.min(min, numeric);
      max = Math.max(max, numeric);
    }
  };

  // Work is bounded by min(selection area, populated cells): small selections walk
  // the rectangle directly; whole-column/sheet selections walk populated cells.
  const selectionArea =
    (range.end.row - range.start.row + 1) * (range.end.column - range.start.column + 1);
  if (hasAtLeastCellCount(sheet.cells, selectionArea)) {
    for (let row = range.start.row; row <= range.end.row; row += 1) {
      for (let column = range.start.column; column <= range.end.column; column += 1) {
        const address = formatCellAddress({ row, column });
        if (address in sheet.cells) {
          accumulate(address);
        }
      }
    }
  } else {
    for (const address in sheet.cells) {
      const coord = parseCellAddress(address);
      if (
        coord.row >= range.start.row &&
        coord.row <= range.end.row &&
        coord.column >= range.start.column &&
        coord.column <= range.end.column
      ) {
        accumulate(address);
      }
    }
  }

  if (nonEmptyCount === 0) {
    return "Count 0";
  }

  if (numericCount === 0) {
    return `Count ${nonEmptyCount}`;
  }

  return `Count ${nonEmptyCount}  Sum ${formatSummaryNumber(sum)}  Avg ${formatSummaryNumber(
    sum / numericCount
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
  return nextGeneratedSheetName(workbook, "Pivot");
}

function nextDetailsSheetName(workbook: WorkbookModel): string {
  return nextGeneratedSheetName(workbook, "Details");
}

function nextGeneratedSheetName(workbook: WorkbookModel, base: string): string {
  const existingNames = new Set(workbook.sheets.map((sheet) => sheet.name));
  let index = 1;
  let name = `${base} ${index}`;
  while (existingNames.has(name)) {
    index += 1;
    name = `${base} ${index}`;
  }
  return name;
}

function nextGeneratedSheetId(workbook: WorkbookModel): string {
  let index = workbook.sheets.length + 1;
  let id = `sheet-${index}`;
  while (workbook.sheets.some((sheet) => sheet.id === id)) {
    index += 1;
    id = `sheet-${index}`;
  }
  return id;
}

function nextConditionalFormatCommandId(items: readonly { id: string }[]): string {
  return nextCommandId(items, "conditional-format-");
}

function nextSheetFilterCommandId(items: readonly { id: string }[]): string {
  return nextCommandId(items, "filter-");
}

function nextSheetChartCommandId(items: readonly { id: string }[]): string {
  return nextCommandId(items, "chart-");
}

function nextCommandId(items: readonly { id: string }[], prefix: string): string {
  const existing = new Set(items.map((item) => item.id));
  let index = items.length + 1;
  let id = `${prefix}${index}`;
  while (existing.has(id)) {
    index += 1;
    id = `${prefix}${index}`;
  }
  return id;
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

function inclusiveIndexes(start: number, end: number): number[] {
  const first = Math.min(start, end);
  const last = Math.max(start, end);
  return Array.from({ length: last - first + 1 }, (_, offset) => first + offset);
}

function flaggedIndexes(flags: Record<string, boolean>): number[] {
  return Object.entries(flags)
    .filter(([, flagged]) => flagged)
    .map(([index]) => Number(index))
    .filter((index) => Number.isInteger(index) && index >= 0)
    .sort((left, right) => left - right);
}

function countDuplicateRows(sheet: SheetModel, selection: CellRange): number {
  const range = normalizeRange(selection);
  const seen = new Set<string>();
  let duplicates = 0;
  for (let row = range.start.row; row <= range.end.row; row += 1) {
    const cells: CellContent[] = [];
    for (let column = range.start.column; column <= range.end.column; column += 1) {
      cells.push(sheet.cells[formatCellAddress({ row, column })] ?? null);
    }
    const key = JSON.stringify(cells.map((cell) => [cell === null ? "blank" : typeof cell, cell]));
    if (seen.has(key)) duplicates += 1;
    else seen.add(key);
  }
  return duplicates;
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

const RESERVED_SERVICE_KEYS = new Set(["__proto__", "prototype", "constructor"]);

function hasInvalidServiceRegistry(
  registry: Readonly<Record<string, WorkbookImporter | WorkbookExporter>> | undefined
): boolean {
  if (!registry) {
    return false;
  }
  return Object.keys(registry).some((key) => !key.trim() || RESERVED_SERVICE_KEYS.has(key));
}

function themeToRootStyle(
  theme: Partial<Record<WorkbookThemeToken, string>> | undefined
): CSSProperties {
  if (!theme) {
    return {};
  }
  const properties: Record<string, string> = {};
  for (const [token, value] of Object.entries(theme)) {
    if (value !== undefined) {
      properties[`--js-spreadsheet-${token}`] = value;
    }
  }
  return properties as CSSProperties;
}

function displaySessionValue(value: ReturnType<WorkbookSession["getCellEvaluation"]>): string {
  if (value === null) {
    return "";
  }
  if (typeof value === "object") {
    return value.code;
  }
  if (typeof value === "boolean") {
    return value ? "TRUE" : "FALSE";
  }
  return String(value);
}

function diagnosticOrigin(event: WorkbookDiagnosticEvent): WorkbookChangeEvent["origin"] {
  const commandType = event.metadata.commandType;
  if (commandType === "history.undo") {
    return "undo";
  }
  if (commandType === "history.redo") {
    return "redo";
  }
  return "command";
}

function invokeHostCallback<T>(callback: ((value: T) => void) | undefined, value: T): void {
  if (!callback) {
    return;
  }
  try {
    callback(value);
  } catch {
    // Host callbacks are isolated from rendering and other subscribers.
  }
}

function downloadWorkbookArtifact(artifact: WorkbookExportArtifact): void {
  const bytes = new Uint8Array(artifact.bytes);
  const blob = new Blob([bytes], { type: artifact.mediaType });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = artifact.fileName;
  anchor.click();
  URL.revokeObjectURL(url);
}
