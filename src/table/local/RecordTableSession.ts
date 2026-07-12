import ExcelJS from "exceljs";
import type { CommandResult } from "../../core/commands/types";
import { excelSerialToDate, parseExcelTemporalInput } from "../../core/values/excelDate";
import { parseCellInput } from "../../core/values/parseCellInput";
import { createLocalTableCapabilities, resolveTableOperationStates, type TableFeatureConfiguration } from "../core/capabilities";
import { createCommandIdFactory, type CommandIdFactory } from "../core/commandId";
import { normalizeColumns } from "../core/columnHelper";
import type { PaginationRequest, QueryRequest, QueryRow, TableAggregateRequest } from "../core/query";
import { safeInvokeTableExtension } from "../core/safeInvoke";
import type {
  ChangeContext,
  ColumnDef,
  ExportArtifact,
  ExportOptions,
  RowUpdater,
  TableCellIssue,
  TableCellMetadata,
  TableDiagnosticEvent,
  TableIntent,
  TableMetadataDocument,
  TableMetadataUpdater,
  TableSession,
  TableStateUpdater,
  TableViewSnapshot,
  TableViewState
} from "../core/types";
import { buildLocalRowModel, type LocalEvaluatedValue, type LocalRowModel } from "./localRowModel";
import {
  createLocalHistory,
  pushLocalHistory,
  redoLocalHistory,
  undoLocalHistory,
  type LocalHistory
} from "./localHistory";
import { createTableMetadataKey, parseTableMetadataKey } from "./tableMetadata";

export type LocalRecordSource<TRow> = {
  kind: "local";
  rows: readonly TRow[];
  getRowId(row: TRow): string;
  getSubRows?(row: TRow): readonly TRow[] | undefined;
  onRowsChange?(updater: RowUpdater<TRow>, context: ChangeContext): void;
  resetKey?: string | number;
};

export type RecordFormulaService<TRow> = {
  evaluate(context: {
    expression: string;
    row: TRow;
    rowId: string;
    columnId: string;
    getValue(columnId: string): unknown;
  }): { value: unknown; displayValue: string } | { issues: readonly TableCellIssue[] };
};

export type LocalRecordTableSessionOptions<
  TRow,
  TColumn extends ColumnDef<TRow, any> = ColumnDef<TRow, any>
> = {
  source: LocalRecordSource<TRow>;
  columns: readonly TColumn[];
  document?: TableMetadataDocument;
  defaultDocument?: TableMetadataDocument;
  onDocumentChange?(updater: TableMetadataUpdater, context: ChangeContext): void;
  state?: Partial<TableViewState>;
  defaultState?: Partial<TableViewState>;
  onStateChange?(updater: TableStateUpdater, context: ChangeContext): void;
  formulaService?: RecordFormulaService<TRow>;
  historyLimit?: number;
  features?: TableFeatureConfiguration;
  commandIdFactory?: CommandIdFactory;
  onDiagnostic?(event: TableDiagnosticEvent): void;
};

export class RecordTableSessionError extends Error {
  constructor(readonly code: string, message: string = code) {
    super(message);
    this.name = "RecordTableSessionError";
  }
}

const EMPTY_TABLE_DOCUMENT: TableMetadataDocument = {
  version: 1,
  cells: {},
  calculatedColumns: [],
  namedStyles: []
};
const DEFAULT_TABLE_VIEW_STATE: TableViewState = {
  sorting: [],
  filter: null,
  grouping: [],
  aggregates: [],
  pagination: { kind: "none" },
  selection: null,
  selectedRowIds: [],
  expandedRowIds: [],
  columnOrder: [],
  columnVisibility: {},
  columnWidths: {},
  columnPinning: { left: [], right: [] }
};
const GROUPING_PAGINATION_ISSUE: TableCellIssue = {
  code: "TABLE_GROUPING_PAGINATION_CONFLICT",
  message: "Grouping requires pagination kind none"
};

type AnyColumn<TRow> = ColumnDef<TRow, any>;
type PositionedRow<TRow> = { rowId: string; row: TRow; index: number };
type RowHistoryOperation<TRow> =
  | { kind: "none" }
  | { kind: "replace"; records: readonly { rowId: string; previous: TRow; row: TRow }[] }
  | { kind: "insert"; records: readonly PositionedRow<TRow>[] }
  | { kind: "delete"; rowIds: readonly string[] };
type MetadataHistoryOperation = {
  key: string;
  value: TableCellMetadata | null;
};
type LocalSessionOperation<TRow> = {
  rows: RowHistoryOperation<TRow>;
  metadata: readonly MetadataHistoryOperation[];
};
type Evaluation = {
  storedValue: unknown;
  evaluatedValue: unknown;
  displayValue?: string;
  formula?: string;
  issues: readonly TableCellIssue[];
};

export class RecordTableSession<
  TRow,
  TColumn extends ColumnDef<TRow, any> = ColumnDef<TRow, any>
> implements TableSession<TRow, TColumn> {
  private options: LocalRecordTableSessionOptions<TRow, TColumn>;
  private columns: readonly TColumn[];
  private columnsById: ReadonlyMap<string, TColumn>;
  private rows: readonly TRow[];
  private document: TableMetadataDocument;
  private state: TableViewState;
  private controlledRows: boolean;
  private controlledDocument: boolean;
  private controlledStateKeys: Set<keyof TableViewState>;
  private sourceRowsReference: readonly TRow[];
  private resetKey: string | number | undefined;
  private revision = 0;
  private snapshot: TableViewSnapshot<TRow, TColumn> | null = null;
  private listeners = new Set<() => void>();
  private destroyed = false;
  private commandIdFactory: CommandIdFactory;
  private history: LocalHistory<LocalSessionOperation<TRow>>;
  private invalidControlledState = false;
  private sessionIssues: TableCellIssue[] = [];

  constructor(options: LocalRecordTableSessionOptions<TRow, TColumn>) {
    validateOptions(options);
    this.options = options;
    this.columns = normalizeColumns(options.columns) as readonly TColumn[];
    this.columnsById = new Map(this.columns.map((column) => [column.id, column]));
    this.rows = [...options.source.rows];
    this.document = cloneDocument(options.document ?? options.defaultDocument ?? EMPTY_TABLE_DOCUMENT);
    this.state = mergeState(defaultStateForColumns(this.columns), options.defaultState, options.state);
    this.controlledRows = Boolean(options.source.onRowsChange);
    this.controlledDocument = options.document !== undefined;
    this.controlledStateKeys = new Set(Object.keys(options.state ?? {}) as (keyof TableViewState)[]);
    this.sourceRowsReference = options.source.rows;
    this.resetKey = options.source.resetKey;
    this.commandIdFactory = options.commandIdFactory ?? createCommandIdFactory();
    this.validateHistoryLimit(options.historyLimit);
    this.history = createLocalHistory(options.historyLimit ?? 100);
    this.validateRows(this.rows);
    this.handleInvalidControlledState();
  }

  updateOptions(options: LocalRecordTableSessionOptions<TRow, TColumn>): void {
    validateOptions(options);
    this.validateHistoryLimit(options.historyLimit);
    const nextHistoryLimit = options.historyLimit ?? 100;
    if (nextHistoryLimit !== this.history.limit) {
      this.history = {
        limit: nextHistoryLimit,
        past: this.history.past.slice(-nextHistoryLimit),
        future: this.history.future.slice(-nextHistoryLimit)
      };
    }
    const previousControlledRows = this.controlledRows;
    const nextControlledRows = Boolean(options.source.onRowsChange);
    const modeChanged = previousControlledRows !== nextControlledRows;
    let externalChanged = options.formulaService !== this.options.formulaService
      || !stateSliceEqual(options.features ?? {}, this.options.features ?? {})
      || options.source.getSubRows !== this.options.source.getSubRows;

    if (modeChanged) {
      this.rows = [...options.source.rows];
      this.clearHistory(nextHistoryLimit);
      this.dropInvalidSelection(options.source, this.rows);
      externalChanged = true;
    } else if (nextControlledRows && options.source.rows !== this.sourceRowsReference) {
      const sameIdentityOrder = rowIdList(this.rows, options.source.getRowId).join("\u0000")
        === rowIdList(options.source.rows, options.source.getRowId).join("\u0000");
      this.rows = [...options.source.rows];
      if (!sameIdentityOrder) {
        this.clearHistory(nextHistoryLimit);
        this.dropInvalidSelection(options.source, this.rows);
      }
      externalChanged = true;
    } else if (!nextControlledRows && options.source.resetKey !== this.resetKey) {
      this.rows = [...options.source.rows];
      this.clearHistory(nextHistoryLimit);
      this.dropInvalidSelection(options.source, this.rows);
      externalChanged = true;
    }

    const nextColumns = normalizeColumns(options.columns) as readonly TColumn[];
    if (!sameColumnReferences(this.columns, nextColumns)) {
      this.columns = nextColumns;
      this.columnsById = new Map(this.columns.map((column) => [column.id, column]));
      externalChanged = true;
    }

    const nextControlledDocument = options.document !== undefined;
    if (nextControlledDocument !== this.controlledDocument) {
      this.document = cloneDocument(options.document ?? options.defaultDocument ?? EMPTY_TABLE_DOCUMENT);
      this.clearHistory(nextHistoryLimit);
      externalChanged = true;
    } else if (nextControlledDocument && options.document !== this.options.document) {
      this.document = cloneDocument(options.document!);
      externalChanged = true;
    }

    const nextStateKeys = new Set(Object.keys(options.state ?? {}) as (keyof TableViewState)[]);
    const controlledCandidate = mergeControlledState(this.state, options.state);
    const safeCandidate = sanitizeLocalViewState(
      controlledCandidate,
      nextColumns,
      options.source.getSubRows !== undefined
    );
    this.options = options;
    this.controlledRows = nextControlledRows;
    this.controlledDocument = nextControlledDocument;
    this.controlledStateKeys = nextStateKeys;
    this.sourceRowsReference = options.source.rows;
    this.resetKey = options.source.resetKey;
    this.validateRows(this.rows);

    if (hasGroupingPaginationConflict(controlledCandidate)) {
      this.invalidControlledState = true;
      this.sessionIssues = [GROUPING_PAGINATION_ISSUE];
      this.requestControlledStateCorrection(controlledCandidate);
      externalChanged = true;
    } else {
      this.invalidControlledState = false;
      this.sessionIssues = [];
      if (!stateEqual(this.state, safeCandidate)) {
        this.state = safeCandidate;
        externalChanged = true;
      }
      if ([...nextStateKeys].some((key) => !stateSliceEqual(controlledCandidate[key], safeCandidate[key]))) {
        this.requestControlledStateReplacement(safeCandidate);
      }
    }

    if (externalChanged) {
      this.revision += 1;
      this.invalidate();
    }
  }

  getSnapshot(): TableViewSnapshot<TRow, TColumn> {
    if (this.snapshot) return this.snapshot;
    const evaluationCache = new Map<string, Evaluation>();
    const evaluationStack = new Set<string>();
    const evaluate = (row: TRow, rowId: string, columnId: string): Evaluation => {
      const key = createTableMetadataKey(rowId, columnId);
      const cached = evaluationCache.get(key);
      if (cached) return cached;
      const column = this.columnsById.get(columnId);
      if (!column) throw new Error(`Unknown column id: ${columnId}`);
      if (evaluationStack.has(key)) {
        const cycle: Evaluation = {
          storedValue: null,
          evaluatedValue: { kind: "error", code: "#ERROR!" },
          displayValue: "#ERROR!",
          issues: [{
            code: "calculated-column-cycle",
            message: "Calculated column cycle",
            rowId,
            columnId
          }]
        };
        return cycle;
      }
      evaluationStack.add(key);
      const metadata = this.document.cells[key] ?? {};
      const calculated = this.document.calculatedColumns.find((item) => item.columnId === columnId);
      let result: Evaluation;
      if (metadata.formula || calculated) {
        result = this.evaluateFormula(metadata.formula ?? calculated!.expression, row, rowId, columnId, evaluate);
      } else if (column.kind === "computed") {
        let dependencyIssue: readonly TableCellIssue[] = [];
        const calculatedResult = safeInvokeTableExtension("calculate", () => column.calculate({
          row,
          rowId,
          columnId,
          getValue: (dependency) => {
            const value = evaluate(row, rowId, dependency);
            if (value.issues.length > 0) dependencyIssue = value.issues;
            return value.evaluatedValue;
          }
        }));
        result = calculatedResult.ok && dependencyIssue.length === 0
          ? { storedValue: null, evaluatedValue: calculatedResult.value, issues: [] }
          : {
              storedValue: null,
              evaluatedValue: { kind: "error", code: "#ERROR!" },
              displayValue: "#ERROR!",
              issues: calculatedResult.ok ? dependencyIssue : [{ ...calculatedResult.issue, rowId, columnId }]
            };
      } else if ("accessor" in column && typeof column.accessor === "function") {
        const accessed = safeInvokeTableExtension("accessor", () => column.accessor(row));
        result = accessed.ok
          ? { storedValue: accessed.value, evaluatedValue: accessed.value, issues: [] }
          : {
              storedValue: null,
              evaluatedValue: { kind: "error", code: "#ERROR!" },
              displayValue: "#ERROR!",
              issues: [{ ...accessed.issue, rowId, columnId }]
            };
      } else {
        result = { storedValue: null, evaluatedValue: null, issues: [] };
      }
      evaluationStack.delete(key);
      evaluationCache.set(key, result);
      return result;
    };

    const model = this.buildModel(evaluate);
    const capabilities = createLocalTableCapabilities({
      formula: this.options.formulaService ? "fullLocalDataset" : "none",
      undo: true
    });
    const operationStates = {
      ...resolveTableOperationStates(capabilities, this.options.features ?? {})
    };
    if (this.invalidControlledState) {
      operationStates.pagination = {
        enabled: false,
        scopeLabel: "Complete dataset",
        reason: GROUPING_PAGINATION_ISSUE.message
      };
    }
    const allIssues = dedupeIssues([...this.sessionIssues, ...model.issues]);
    const snapshotRows = model.items;

    const snapshot: TableViewSnapshot<TRow, TColumn> = {
      revision: String(this.revision),
      rows: snapshotRows,
      columns: this.columns,
      rowCount: snapshotRows.length,
      totalRowCount: { kind: "known", value: model.totalDataRowCount },
      completeness: "completeDataset",
      state: this.state,
      selection: this.state.selection,
      status: this.invalidControlledState
        ? { phase: "error", message: GROUPING_PAGINATION_ISSUE.message }
        : { phase: "ready" },
      issues: allIssues,
      capabilities,
      operationStates,
      pendingOperations: [],
      conflicts: [],
      canUndo: this.history.past.length > 0,
      canRedo: this.history.future.length > 0,
      pageInfo: model.pageInfo,
      getCell: (rowId, columnId) => {
        const column = this.columnsById.get(columnId);
        if (!column) throw new Error(`Unknown column id: ${columnId}`);
        const row = model.dataRowsById.get(rowId);
        if (!row) return groupOrAggregateCell(snapshotRows, rowId, columnId, this.state.aggregates);
        const evaluation = evaluate(row, rowId, columnId);
        const metadata = this.document.cells[createTableMetadataKey(rowId, columnId)] ?? {};
        const context = this.columnContext(row, rowId, columnId, evaluate);
        const formatted = evaluation.displayValue === undefined && column.format && !isErrorValue(evaluation.evaluatedValue)
          ? safeInvokeTableExtension("format", () => column.format!(evaluation.evaluatedValue, context))
          : null;
        const formatIssue = formatted && !formatted.ok ? [{ ...formatted.issue, rowId, columnId }] : [];
        const editability = this.isEditable(column, context, metadata);
        return {
          rowId,
          columnId,
          storedValue: evaluation.storedValue,
          evaluatedValue: evaluation.evaluatedValue,
          displayValue: evaluation.displayValue
            ?? (formatted?.ok ? formatted.value : formatDefault(evaluation.evaluatedValue)),
          ...(metadata.formula ? { formula: metadata.formula } : {}),
          metadata,
          editable: editability.editable,
          issues: dedupeIssues([
            ...evaluation.issues,
            ...formatIssue,
            ...(editability.issue ? [editability.issue] : [])
          ])
        };
      },
      getRowIndex: (rowId) => snapshotRows.findIndex((row) => row.id === rowId),
      getColumnIndex: (columnId) => this.columns.findIndex((column) => column.id === columnId)
    };
    this.snapshot = snapshot;
    return snapshot;
  }

  subscribe(listener: () => void): () => void {
    if (this.destroyed) return () => {};
    this.listeners.add(listener);
    return () => this.listeners.delete(listener);
  }

  async dispatch(intent: TableIntent<TRow>): Promise<CommandResult> {
    const commandId = this.commandIdFactory();
    const startedAt = now();
    if (this.destroyed) return this.rejectUnsupported(commandId, intent, startedAt, "Session is destroyed");
    if (intent.type === "reload-authoritative" || intent.type === "retry-with-revision") {
      return this.rejectUnsupported(commandId, intent, startedAt, "Conflict resolution is not supported by local tables");
    }

    const feature = intentFeature(intent);
    if (feature && !this.getSnapshot().operationStates[feature].enabled) {
      return this.rejectUnsupported(commandId, intent, startedAt, this.getSnapshot().operationStates[feature].reason);
    }
    if (intent.type === "undo") return this.replayHistory("undo", commandId, startedAt);
    if (intent.type === "redo") return this.replayHistory("redo", commandId, startedAt);
    if (intent.type === "refresh") return this.finishNoChange(commandId, intent, startedAt);

    if (intent.type === "edit-cells" || intent.type === "clear-cells") {
      const edits = intent.type === "edit-cells"
        ? intent.edits
        : intent.cells.map((cell) => ({ ...cell, rawText: "" }));
      return this.editCells(edits, intent.type, commandId, startedAt);
    }
    if (intent.type === "update-cell-metadata") {
      return this.updateMetadata(intent.updates, commandId, startedAt);
    }
    if (intent.type === "insert-rows") return this.insertRows(intent, commandId, startedAt);
    if (intent.type === "delete-rows") return this.deleteRows(intent.rowIds, commandId, startedAt);

    return this.updateViewState(intent, commandId, startedAt);
  }

  async refresh(): Promise<void> {
    const commandId = this.commandIdFactory();
    this.emitDiagnostic(commandId, "refresh", now(), false, "command", 0, 0);
  }

  async undo(): Promise<CommandResult> {
    const commandId = this.commandIdFactory();
    const startedAt = now();
    if (!this.getSnapshot().operationStates.undo.enabled) {
      return this.rejectUnsupported(commandId, { type: "undo" }, startedAt, this.getSnapshot().operationStates.undo.reason);
    }
    return this.replayHistory("undo", commandId, startedAt);
  }

  async redo(): Promise<CommandResult> {
    const commandId = this.commandIdFactory();
    const startedAt = now();
    if (!this.getSnapshot().operationStates.undo.enabled) {
      return this.rejectUnsupported(commandId, { type: "redo" }, startedAt, this.getSnapshot().operationStates.undo.reason);
    }
    return this.replayHistory("redo", commandId, startedAt);
  }

  async export(options: ExportOptions): Promise<ExportArtifact> {
    const commandId = this.commandIdFactory();
    const startedAt = now();
    if (this.destroyed) throw new Error("Table session is destroyed");
    const operation = this.getSnapshot().operationStates.export;
    if (!operation.enabled) {
      throw new RecordTableSessionError(
        "TABLE_CAPABILITY_UNSUPPORTED",
        operation.reason ?? "Local export is unsupported"
      );
    }
    const { rows, columns } = this.exportRows(options.scope);
    const fileName = exportFileName(options.fileName, options.format);
    let artifact: ExportArtifact;
    if (options.format === "csv") {
      const lines: string[] = [];
      if (options.includeHeaders !== false) lines.push(columns.map((column) => csvEscape(column.header)).join(","));
      for (const row of rows) lines.push(row.map((value) => csvEscape(formatExportText(value))).join(","));
      artifact = {
        bytes: new TextEncoder().encode(`${lines.join("\r\n")}\r\n`),
        mediaType: "text/csv;charset=utf-8",
        fileName
      };
    } else {
      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet("Table");
      if (options.includeHeaders !== false) worksheet.addRow(columns.map((column) => column.header));
      rows.forEach((row) => worksheet.addRow(row.map(toExcelValue)));
      const buffer = await workbook.xlsx.writeBuffer();
      artifact = {
        bytes: new Uint8Array(buffer),
        mediaType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        fileName
      };
    }
    this.emitDiagnostic(commandId, "export", startedAt, false, "command", rows.length, rows.length * columns.length);
    return { ...artifact, bytes: new Uint8Array(artifact.bytes) };
  }

  destroy(): void {
    if (this.destroyed) return;
    this.destroyed = true;
    this.listeners.clear();
    this.snapshot = null;
  }

  private async editCells(
    edits: readonly { rowId: string; columnId: string; rawText: string }[],
    reason: "edit-cells" | "clear-cells",
    commandId: string,
    startedAt: number
  ): Promise<CommandResult> {
    const prepared: Array<{
      edit: { rowId: string; columnId: string; rawText: string };
      metadataKey: string;
      value: unknown;
      formula: string | undefined;
      update: (row: TRow, value: any) => TRow;
    }> = [];
    const formulaOperation = this.getSnapshot().operationStates.formula;

    for (const edit of edits) {
      const rowIndex = findRowIndex(this.rows, edit.rowId, this.options.source.getRowId);
      if (rowIndex < 0) return this.reject("validation", [{ code: "TABLE_ROW_NOT_FOUND", message: "Row not found", rowId: edit.rowId }], commandId, reason, startedAt, edits.length, edits.length);
      const row = this.rows[rowIndex];
      const column = this.columnsById.get(edit.columnId);
      if (!column) return this.reject("validation", [{ code: "TABLE_COLUMN_NOT_FOUND", message: "Column not found", rowId: edit.rowId, columnId: edit.columnId }], commandId, reason, startedAt, edits.length, edits.length);
      const context = this.columnContext(row, edit.rowId, edit.columnId);
      const metadataKey = createTableMetadataKey(edit.rowId, edit.columnId);
      const metadata = this.document.cells[metadataKey] ?? {};
      const permission = this.isEditable(column, context, metadata);
      if (permission.issue) return this.reject("unsupported", [permission.issue], commandId, reason, startedAt, edits.length, edits.length);
      if (!permission.editable) return this.reject("permission", [{ code: "TABLE_CELL_READ_ONLY", message: "Cell is read-only", rowId: edit.rowId, columnId: edit.columnId }], commandId, reason, startedAt, edits.length, edits.length);

      const input = parseCellInput(edit.rawText);
      let value: unknown;
      let formula: string | undefined;
      let evaluated: unknown;
      if (column.parse) {
        const parsed = safeInvokeTableExtension("parse", () => column.parse!(input, context));
        if (!parsed.ok) return this.reject("unsupported", [{ ...parsed.issue, rowId: edit.rowId, columnId: edit.columnId }], commandId, reason, startedAt, edits.length, edits.length);
        if (!parsed.value.ok) return this.reject("validation", parsed.value.issues, commandId, reason, startedAt, edits.length, edits.length);
        value = parsed.value.value;
        formula = parsed.value.formula;
        evaluated = parsed.value.evaluatedValue ?? value;
      } else if (input.formula) {
        formula = input.formula;
        if (!formulaOperation.enabled) {
          return this.rejectUnsupported(commandId, editIntent(reason, edits), startedAt, formulaOperation.reason);
        }
        const formulaResult = this.invokeFormula(formula, row, edit.rowId, edit.columnId);
        if ("issues" in formulaResult) return this.reject("validation", formulaResult.issues, commandId, reason, startedAt, edits.length, edits.length);
        value = formulaResult.value;
        evaluated = formulaResult.value;
      } else {
        value = input.stored;
        evaluated = value;
      }
      const originalValue = value;
      const normalizedValue = normalizeLocalValue(value, column.dataType);
      if (!normalizedValue.ok) {
        return this.reject("validation", [{ code: "TABLE_VALUE_TYPE", message: "Cell value has an incompatible type", rowId: edit.rowId, columnId: edit.columnId }], commandId, reason, startedAt, edits.length, edits.length);
      }
      value = normalizedValue.value;
      if (Object.is(evaluated, originalValue)) evaluated = value;
      if (formula && !formulaOperation.enabled) {
        return this.rejectUnsupported(commandId, editIntent(reason, edits), startedAt, formulaOperation.reason);
      }
      if (!compatibleValue(value, column.dataType)) {
        return this.reject("validation", [{ code: "TABLE_VALUE_TYPE", message: "Cell value has an incompatible type", rowId: edit.rowId, columnId: edit.columnId }], commandId, reason, startedAt, edits.length, edits.length);
      }
      if (column.validate) {
        const validation = safeInvokeTableExtension("validate", () => column.validate!({
          ...context,
          raw: edit.rawText,
          parsed: value,
          evaluated
        }));
        if (!validation.ok) return this.reject("unsupported", [{ ...validation.issue, rowId: edit.rowId, columnId: edit.columnId }], commandId, reason, startedAt, edits.length, edits.length);
        if (validation.value.length > 0) return this.reject("validation", validation.value, commandId, reason, startedAt, edits.length, edits.length);
      }
      const metadataIssues = validateMetadataValue(value, metadata.validation, edit.rowId, edit.columnId);
      if (metadataIssues.length > 0) return this.reject("validation", metadataIssues, commandId, reason, startedAt, edits.length, edits.length);
      const update = "update" in column ? column.update : undefined;
      if (typeof update !== "function") {
        return this.reject("permission", [{ code: "TABLE_CELL_READ_ONLY", message: "Cell is read-only", rowId: edit.rowId, columnId: edit.columnId }], commandId, reason, startedAt, edits.length, edits.length);
      }
      let current: unknown = null;
      if ("accessor" in column && typeof column.accessor === "function") {
        const accessed = safeInvokeTableExtension("accessor", () => column.accessor(row));
        if (!accessed.ok) {
          return this.reject("unsupported", [{ ...accessed.issue, rowId: edit.rowId, columnId: edit.columnId }], commandId, reason, startedAt, edits.length, edits.length);
        }
        current = accessed.value;
      }
      const currentFormula = metadata.formula;
      if (Object.is(current, value) && currentFormula === formula) continue;
      prepared.push({ edit, metadataKey, value, formula, update });
    }

    if (prepared.length === 0) return this.finishNoChange(commandId, { type: reason } as TableIntent<TRow>, startedAt);

    let candidateRows = this.rows;
    let candidateDocument = this.document;
    for (const candidate of prepared) {
      const rowIndex = findRowIndex(candidateRows, candidate.edit.rowId, this.options.source.getRowId);
      const row = candidateRows[rowIndex];
      const updated = safeInvokeTableExtension("update", () => candidate.update(row, candidate.value));
      if (!updated.ok) return this.reject("unsupported", [{ ...updated.issue, rowId: candidate.edit.rowId, columnId: candidate.edit.columnId }], commandId, reason, startedAt, edits.length, edits.length);
      candidateRows = replaceAt(candidateRows, rowIndex, updated.value);
      candidateDocument = patchFormula(candidateDocument, candidate.metadataKey, candidate.formula);
    }
    const candidateIds = this.safeRowIds(candidateRows);
    if (!candidateIds.ok) return this.reject("unsupported", [candidateIds.issue], commandId, reason, startedAt, edits.length, edits.length);
    const originalIds = this.safeRowIds(this.rows);
    if (!originalIds.ok) return this.reject("unsupported", [originalIds.issue], commandId, reason, startedAt, edits.length, edits.length);
    const changedIdIndex = candidateIds.value.findIndex((id, index) => id !== originalIds.value[index]);
    if (changedIdIndex >= 0) {
      return this.reject("validation", [{ code: "TABLE_ROW_ID_CHANGED", message: "Cell edits cannot change stable row IDs", rowId: originalIds.value[changedIdIndex] }], commandId, reason, startedAt, edits.length, edits.length);
    }
    const duplicateId = firstDuplicate(candidateIds.value);
    if (duplicateId) {
      return this.reject("validation", [{ code: "TABLE_ROW_ID_DUPLICATE", message: "Row IDs must remain unique", rowId: duplicateId }], commandId, reason, startedAt, edits.length, edits.length);
    }
    return this.commitDurable(candidateRows, candidateDocument, reason, commandId, startedAt, edits.length, edits.length);
  }

  private async updateMetadata(
    updates: readonly { rowId: string; columnId: string; patch: Partial<TableCellMetadata> }[],
    commandId: string,
    startedAt: number
  ): Promise<CommandResult> {
    let candidate = this.document;
    for (const update of updates) {
      if (findRowIndex(this.rows, update.rowId, this.options.source.getRowId) < 0 || !this.columnsById.has(update.columnId)) {
        return this.reject("validation", [{ code: "TABLE_CELL_NOT_FOUND", message: "Cell not found", rowId: update.rowId, columnId: update.columnId }], commandId, "update-cell-metadata", startedAt, updates.length, updates.length);
      }
      const key = createTableMetadataKey(update.rowId, update.columnId);
      const next = { ...(candidate.cells[key] ?? {}), ...update.patch };
      candidate = { ...candidate, cells: { ...candidate.cells, [key]: next } };
    }
    if (documentEqual(candidate, this.document)) return this.finishNoChange(commandId, { type: "update-cell-metadata", updates } as TableIntent<TRow>, startedAt);
    return this.commitDurable(this.rows, candidate, "update-cell-metadata", commandId, startedAt, updates.length, updates.length);
  }

  private async insertRows(
    intent: Extract<TableIntent<TRow>, { type: "insert-rows" }>,
    commandId: string,
    startedAt: number
  ): Promise<CommandResult> {
    if ("count" in intent) {
      return this.rejectUnsupported(commandId, intent, startedAt, "Local row insertion requires supplied host records");
    }

    const beforeRowId = typeof intent.beforeRowId === "string" ? intent.beforeRowId.trim() : undefined;
    const afterRowId = typeof intent.afterRowId === "string" ? intent.afterRowId.trim() : undefined;
    if ((beforeRowId && afterRowId) || (intent.beforeRowId !== undefined && !beforeRowId) || (intent.afterRowId !== undefined && !afterRowId)) {
      return this.reject("validation", [{ code: "TABLE_ROW_ANCHOR_INVALID", message: "Provide at most one nonblank row anchor" }], commandId, intent.type, startedAt, intent.rows.length, 0);
    }

    const existingIds = this.safeRowIds(this.rows);
    if (!existingIds.ok) return this.reject("unsupported", [existingIds.issue], commandId, intent.type, startedAt, intent.rows.length, 0);
    const insertedIds = this.safeRowIds(intent.rows);
    if (!insertedIds.ok) return this.reject("unsupported", [insertedIds.issue], commandId, intent.type, startedAt, intent.rows.length, 0);
    const allIds = [...existingIds.value, ...insertedIds.value];
    const duplicate = firstDuplicate(allIds);
    if (duplicate) {
      return this.reject("validation", [{ code: "TABLE_ROW_ID_DUPLICATE", message: "Row IDs must be unique", rowId: duplicate }], commandId, intent.type, startedAt, intent.rows.length, 0);
    }

    const anchorId = beforeRowId ?? afterRowId;
    const anchorIndex = anchorId === undefined ? this.rows.length : existingIds.value.indexOf(anchorId);
    if (anchorId !== undefined && anchorIndex < 0) {
      return this.reject("validation", [{ code: "TABLE_ROW_ANCHOR_NOT_FOUND", message: "Row anchor not found", rowId: anchorId }], commandId, intent.type, startedAt, intent.rows.length, 0);
    }
    if (intent.rows.length === 0) return this.finishNoChange(commandId, intent, startedAt);

    const insertionIndex = afterRowId ? anchorIndex + 1 : anchorIndex;
    const rows = [...this.rows];
    rows.splice(insertionIndex, 0, ...intent.rows);
    return this.commitDurable(rows, this.document, intent.type, commandId, startedAt, intent.rows.length, 0);
  }

  private async deleteRows(
    requestedIds: readonly string[],
    commandId: string,
    startedAt: number
  ): Promise<CommandResult> {
    const rowIds = requestedIds.map((id) => id.trim());
    if (rowIds.some((id) => id.length === 0) || firstDuplicate(rowIds)) {
      return this.reject("validation", [{ code: "TABLE_ROW_ID_INVALID", message: "Delete row IDs must be unique and nonblank" }], commandId, "delete-rows", startedAt, rowIds.length, 0);
    }
    if (rowIds.length === 0) return this.finishNoChange(commandId, { type: "delete-rows", rowIds }, startedAt);

    const existingIds = this.safeRowIds(this.rows);
    if (!existingIds.ok) return this.reject("unsupported", [existingIds.issue], commandId, "delete-rows", startedAt, rowIds.length, 0);
    const existing = new Set(existingIds.value);
    const missing = rowIds.find((id) => !existing.has(id));
    if (missing) {
      return this.reject("validation", [{ code: "TABLE_ROW_NOT_FOUND", message: "Row not found", rowId: missing }], commandId, "delete-rows", startedAt, rowIds.length, 0);
    }

    const deleted = new Set(rowIds);
    const rows = this.rows.filter((_, index) => !deleted.has(existingIds.value[index]));
    const cells = Object.fromEntries(Object.entries(this.document.cells).filter(([key]) => {
      const address = safeParseMetadataKey(key);
      return !address || !deleted.has(address.rowId);
    }));
    const document = { ...this.document, cells };
    return this.commitDurable(
      rows,
      document,
      "delete-rows",
      commandId,
      startedAt,
      rowIds.length,
      0,
      pruneDeletedRowsFromState(this.state, deleted)
    );
  }

  private async updateViewState(
    intent: Exclude<TableIntent<TRow>, { type: "edit-cells" | "clear-cells" | "update-cell-metadata" | "insert-rows" | "delete-rows" | "reload-authoritative" | "retry-with-revision" | "undo" | "redo" | "refresh" }>,
    commandId: string,
    startedAt: number
  ): Promise<CommandResult> {
    let candidate = this.state;
    switch (intent.type) {
      case "set-selection": candidate = { ...candidate, selection: intent.selection }; break;
      case "set-row-selection": candidate = { ...candidate, selectedRowIds: [...intent.rowIds] }; break;
      case "set-row-expanded": candidate = {
        ...candidate,
        expandedRowIds: intent.expanded
          ? [...new Set([...candidate.expandedRowIds, intent.rowId])]
          : candidate.expandedRowIds.filter((id) => id !== intent.rowId)
      }; break;
      case "set-sorting": candidate = { ...candidate, sorting: [...intent.sorting] }; break;
      case "set-filter": candidate = { ...candidate, filter: intent.filter }; break;
      case "set-grouping": candidate = {
        ...candidate,
        grouping: [...intent.grouping],
        ...(intent.grouping.length > 0 ? { pagination: { kind: "none" } as const } : {})
      }; break;
      case "set-aggregates": candidate = { ...candidate, aggregates: [...intent.aggregates] }; break;
      case "set-pagination":
        if (candidate.grouping.length > 0 && intent.pagination.kind !== "none") {
          return this.reject("validation", [GROUPING_PAGINATION_ISSUE], commandId, intent.type, startedAt, 0, 0);
        }
        candidate = { ...candidate, pagination: intent.pagination };
        break;
      case "set-column-order": candidate = { ...candidate, columnOrder: [...intent.columnIds] }; break;
      case "resize-column": candidate = { ...candidate, columnWidths: { ...candidate.columnWidths, [intent.columnId]: intent.width } }; break;
      case "set-column-visibility": candidate = { ...candidate, columnVisibility: { ...candidate.columnVisibility, [intent.columnId]: intent.visible } }; break;
      case "set-column-pinning": candidate = pinColumn(candidate, intent.columnId, intent.pin); break;
    }
    if (isQueryStateIntent(intent)) {
      const issue = this.validateViewState(candidate);
      if (issue) return this.reject("validation", [issue], commandId, intent.type, startedAt, 0, 0);
    }
    if (stateEqual(candidate, this.state)) return this.finishNoChange(commandId, intent, startedAt);
    const previous = this.state;
    this.state = candidate;
    this.revision += 1;
    this.invalidate();
    const context = this.context(commandId, intent.type);
    if ([...this.controlledStateKeys].some((key) => !stateSliceEqual(previous[key], candidate[key]))) {
      this.safeHostCallback(() => this.options.onStateChange?.((state) => applyControlledStateDiff(state, previous, candidate, this.controlledStateKeys), context));
    }
    this.publish();
    this.emitDiagnostic(commandId, intent.type, startedAt, true, "command", 0, 0);
    return { status: "committed", revision: String(this.revision), changed: true };
  }

  private async commitDurable(
    rows: readonly TRow[],
    document: TableMetadataDocument,
    reason: TableIntent<TRow>["type"],
    commandId: string,
    startedAt: number,
    rowCount: number,
    cellCount: number,
    nextState: TableViewState = this.state
  ): Promise<CommandResult> {
    const beforeRows = this.rows;
    const beforeDocument = this.document;
    const beforeState = this.state;
    const redo = createHistoryOperation(beforeRows, rows, beforeDocument, document, this.options.source.getRowId);
    const undo = createHistoryOperation(rows, beforeRows, document, beforeDocument, this.options.source.getRowId);
    this.rows = rows;
    this.document = document;
    this.state = nextState;
    this.history = pushLocalHistory(this.history, undo, redo);
    this.revision += 1;
    this.invalidate();
    const context = this.context(commandId, reason);
    if (this.controlledRows && !rowsEqualByReference(beforeRows, rows)) {
      this.safeHostCallback(() => this.options.source.onRowsChange?.(rowsUpdater(beforeRows, rows, this.options.source.getRowId), context));
    }
    if (this.controlledDocument && !documentEqual(beforeDocument, document)) {
      this.safeHostCallback(() => this.options.onDocumentChange?.(documentUpdater(beforeDocument, document), context));
    }
    if ([...this.controlledStateKeys].some((key) => !stateSliceEqual(beforeState[key], nextState[key]))) {
      this.safeHostCallback(() => this.options.onStateChange?.(
        (state) => applyControlledStateDiff(state, beforeState, nextState, this.controlledStateKeys),
        context
      ));
    }
    this.publish();
    this.emitDiagnostic(commandId, reason, startedAt, true, "command", rowCount, cellCount);
    return { status: "committed", revision: String(this.revision), changed: true };
  }

  private async replayHistory(kind: "undo" | "redo", commandId: string, startedAt: number): Promise<CommandResult> {
    if (this.destroyed) return this.rejectUnsupported(commandId, { type: kind } as TableIntent<TRow>, startedAt, "Session is destroyed");
    const replay = kind === "undo" ? undoLocalHistory(this.history) : redoLocalHistory(this.history);
    if (!replay.operation) return this.finishNoChange(commandId, { type: kind } as TableIntent<TRow>, startedAt);
    const beforeRows = this.rows;
    const beforeDocument = this.document;
    const applied = applyHistoryOperation(beforeRows, beforeDocument, replay.operation, this.options.source.getRowId);
    this.history = replay.history;
    this.rows = applied.rows;
    this.document = applied.document;
    this.revision += 1;
    this.invalidate();
    const context = this.context(commandId, kind);
    if (this.controlledRows && !rowsEqualByReference(beforeRows, this.rows)) {
      this.safeHostCallback(() => this.options.source.onRowsChange?.(rowsUpdater(beforeRows, this.rows, this.options.source.getRowId), context));
    }
    if (this.controlledDocument && !documentEqual(beforeDocument, this.document)) {
      this.safeHostCallback(() => this.options.onDocumentChange?.(documentUpdater(beforeDocument, this.document), context));
    }
    this.publish();
    this.emitDiagnostic(commandId, kind, startedAt, true, "command", 0, 0);
    return { status: "committed", revision: String(this.revision), changed: true };
  }

  private buildModel(evaluate: (row: TRow, rowId: string, columnId: string) => Evaluation): LocalRowModel<TRow> {
    return buildLocalRowModel(
      this.rows,
      this.columns,
      queryFromState(this.state),
      this.options.source.getRowId,
      ({ row, rowId, columnId }) => {
        const value = evaluate(row, rowId, columnId);
        return value.issues.length > 0 && isErrorValue(value.evaluatedValue)
          ? { kind: "error", code: value.evaluatedValue.code }
          : { kind: "value", value: value.evaluatedValue };
      },
      this.options.source.getSubRows
    );
  }

  private evaluateFormula(
    expression: string,
    row: TRow,
    rowId: string,
    columnId: string,
    evaluate: (row: TRow, rowId: string, columnId: string) => Evaluation
  ): Evaluation {
    if (!this.options.formulaService) {
      return {
        storedValue: null,
        evaluatedValue: { kind: "error", code: "#ERROR!" },
        displayValue: "#ERROR!",
        formula: expression,
        issues: [{ code: "formula-service-missing", message: "Formula service is not configured", rowId, columnId }]
      };
    }
    const result = safeInvokeTableExtension("formula", () => this.options.formulaService!.evaluate({
      expression,
      row,
      rowId,
      columnId,
      getValue: (dependency) => evaluate(row, rowId, dependency).evaluatedValue
    }));
    if (!result.ok) {
      return {
        storedValue: null,
        evaluatedValue: { kind: "error", code: "#ERROR!" },
        displayValue: "#ERROR!",
        formula: expression,
        issues: [{ ...result.issue, rowId, columnId }]
      };
    }
    if ("issues" in result.value) {
      return {
        storedValue: null,
        evaluatedValue: { kind: "error", code: "#ERROR!" },
        displayValue: "#ERROR!",
        formula: expression,
        issues: result.value.issues
      };
    }
    return {
      storedValue: result.value.value,
      evaluatedValue: result.value.value,
      displayValue: result.value.displayValue,
      formula: expression,
      issues: []
    };
  }

  private invokeFormula(expression: string, row: TRow, rowId: string, columnId: string) {
    const result = safeInvokeTableExtension("formula", () => this.options.formulaService!.evaluate({
      expression,
      row,
      rowId,
      columnId,
      getValue: (dependency) => {
        const column = this.columnsById.get(dependency);
        return column && "accessor" in column && typeof column.accessor === "function" ? column.accessor(row) : null;
      }
    }));
    return result.ok ? result.value : { issues: [{ ...result.issue, rowId, columnId }] };
  }

  private columnContext(
    row: TRow,
    rowId: string,
    columnId: string,
    evaluate?: (row: TRow, rowId: string, columnId: string) => Evaluation
  ) {
    return {
      row,
      rowId,
      columnId,
      getValue: (dependency: string) => {
        if (evaluate) return evaluate(row, rowId, dependency).evaluatedValue;
        const column = this.columnsById.get(dependency);
        return column && "accessor" in column && typeof column.accessor === "function" ? column.accessor(row) : null;
      }
    };
  }

  private isEditable(column: TColumn, context: ReturnType<RecordTableSession<TRow, TColumn>["columnContext"]>, metadata: TableCellMetadata) {
    if (metadata.readOnly || !("update" in column) || typeof column.update !== "function") return { editable: false };
    if (typeof column.editable === "boolean" && !column.editable) return { editable: false };
    if (typeof column.editable === "function") {
      const editable = column.editable;
      const result = safeInvokeTableExtension("permission", () => editable(context));
      if (!result.ok) return { editable: false, issue: { ...result.issue, rowId: context.rowId, columnId: context.columnId } };
      if (!result.value) return { editable: false };
    }
    if (column.permitted) {
      const result = safeInvokeTableExtension("permission", () => column.permitted!(context));
      if (!result.ok) return { editable: false, issue: { ...result.issue, rowId: context.rowId, columnId: context.columnId } };
      if (!result.value) return { editable: false };
    }
    return { editable: true };
  }

  private exportRows(scope: ExportOptions["scope"]): { rows: unknown[][]; columns: { id: string; header: string }[] } {
    const snapshot = this.getSnapshot();
    const orderedColumns = exportColumns<TRow, TColumn>(this.columns, this.state);
    const rowIds = scope === "currentView"
      ? snapshot.rows
          .filter((row): row is Extract<QueryRow<TRow>, { kind: "data" }> => row.kind === "data")
          .map((row) => row.id)
      : this.buildUnpaginatedModel().orderedDataRowIds;
    return {
      columns: orderedColumns.map((column) => ({ id: column.id, header: typeof column.header === "string" ? column.header : column.id })),
      rows: rowIds.map((rowId) => orderedColumns.map((column) => snapshot.getCell(rowId, column.id).evaluatedValue))
    };
  }

  private buildUnpaginatedModel(): LocalRowModel<TRow> {
    const state = { ...this.state, pagination: { kind: "none" } as const };
    const currentSnapshot = this.getSnapshot();
    return buildLocalRowModel(
      this.rows,
      this.columns,
      queryFromState(state),
      this.options.source.getRowId,
      ({ rowId, columnId }) => {
        const value = currentSnapshot.getCell(rowId, columnId).evaluatedValue;
        return isErrorValue(value) ? { kind: "error", code: value.code } : { kind: "value", value };
      },
      this.options.source.getSubRows
    );
  }

  private validateRows(rows: readonly TRow[]): void {
    buildLocalRowModel(rows, this.columns, queryFromState(DEFAULT_TABLE_VIEW_STATE), this.options.source.getRowId, undefined, this.options.source.getSubRows);
  }

  private validateViewState(candidate: TableViewState): TableCellIssue | null {
    try {
      buildLocalRowModel(
        this.rows,
        this.columns,
        queryFromState(candidate),
        this.options.source.getRowId,
        undefined,
        this.options.source.getSubRows
      );
      return null;
    } catch (error) {
      return {
        code: "TABLE_VIEW_STATE_INVALID",
        message: error instanceof Error ? error.message : "Table view state is invalid"
      };
    }
  }

  private safeRowIds(rows: readonly TRow[]):
    | { ok: true; value: string[] }
    | { ok: false; issue: TableCellIssue } {
    const ids: string[] = [];
    for (const row of rows) {
      const result = safeInvokeTableExtension("accessor", () => this.options.source.getRowId(row));
      if (!result.ok) return result;
      if (typeof result.value !== "string" || result.value.trim().length === 0) {
        return {
          ok: false,
          issue: { code: "TABLE_ROW_ID_INVALID", message: "Row IDs must be nonblank strings" }
        };
      }
      ids.push(result.value.trim());
    }
    return { ok: true, value: ids };
  }

  private validateHistoryLimit(limit = 100): void {
    if (!Number.isInteger(limit) || limit < 1 || limit > 1_000) throw new Error("historyLimit must be an integer from 1 through 1000");
  }

  private handleInvalidControlledState(): void {
    if (!hasGroupingPaginationConflict(this.state)) return;
    this.invalidControlledState = true;
    this.sessionIssues = [GROUPING_PAGINATION_ISSUE];
    this.state = { ...this.state, pagination: { kind: "none" } };
    this.requestControlledStateCorrection(this.state);
  }

  private requestControlledStateCorrection(candidate: TableViewState): void {
    const corrected = { ...candidate, pagination: { kind: "none" } as const };
    this.requestControlledStateReplacement(corrected);
  }

  private requestControlledStateReplacement(corrected: TableViewState): void {
    if (!this.options.onStateChange) return;
    this.safeHostCallback(() => this.options.onStateChange?.(() => corrected, this.context(this.commandIdFactory(), "set-pagination")));
  }

  private dropInvalidSelection(source: LocalRecordSource<TRow>, rows: readonly TRow[]): void {
    const ids = new Set(rowIdList(rows, source.getRowId));
    const selection = this.state.selection;
    if (selection && (!ids.has(selection.anchor.rowId) || !ids.has(selection.focus.rowId))) {
      this.state = { ...this.state, selection: null };
    }
  }

  private context(commandId: string, reason: TableIntent<TRow>["type"]): ChangeContext {
    return { commandId, reason, revision: String(this.revision) };
  }

  private invalidate(): void {
    this.snapshot = null;
  }

  private clearHistory(limit = this.options.historyLimit ?? 100): void {
    this.history = createLocalHistory(limit);
  }

  private publish(): void {
    for (const listener of [...this.listeners]) {
      const result = safeInvokeTableExtension("subscriber", listener);
      if (!result.ok) this.emitExtensionDiagnostic("subscriber");
    }
  }

  private safeHostCallback(invoke: () => void): void {
    const result = safeInvokeTableExtension("host-callback", invoke);
    if (!result.ok) this.emitExtensionDiagnostic("host-callback");
  }

  private emitExtensionDiagnostic(kind: "subscriber" | "host-callback"): void {
    const event: TableDiagnosticEvent = {
      category: "extension",
      metadata: { extensionKind: kind, code: "TABLE_EXTENSION_ERROR" }
    };
    safeInvokeTableExtension("host-callback", () => this.options.onDiagnostic?.(event));
  }

  private emitDiagnostic(
    commandId: string,
    commandType: string,
    startedAt: number,
    changed: boolean,
    category: TableDiagnosticEvent["category"],
    rowCount: number,
    cellCount: number,
    outcome: "committed" | "rejected" = "committed"
  ): void {
    const event: TableDiagnosticEvent = {
      category,
      commandId,
      durationMs: Math.max(0, now() - startedAt),
      metadata: { commandType, changed, outcome, rowCount, cellCount }
    };
    safeInvokeTableExtension("host-callback", () => this.options.onDiagnostic?.(event));
  }

  private reject(
    reason: "validation" | "permission" | "unsupported",
    issues: readonly TableCellIssue[],
    commandId: string,
    commandType: string,
    startedAt: number,
    rowCount: number,
    cellCount: number
  ): CommandResult {
    this.emitDiagnostic(commandId, commandType, startedAt, false, reason === "validation" ? "validation" : "command", rowCount, cellCount, "rejected");
    return { status: "rejected", reason, issues };
  }

  private rejectUnsupported(
    commandId: string,
    intent: TableIntent<TRow>,
    startedAt: number,
    message?: string
  ): CommandResult {
    return this.reject("unsupported", [{ code: "TABLE_OPERATION_UNSUPPORTED", message: message ?? "Operation is not supported" }], commandId, intent.type, startedAt, 0, 0);
  }

  private finishNoChange(commandId: string, intent: TableIntent<TRow>, startedAt: number): CommandResult {
    this.emitDiagnostic(commandId, intent.type, startedAt, false, "command", 0, 0);
    return { status: "committed", revision: String(this.revision), changed: false };
  }
}

export function createLocalRecordTableSession<
  TRow,
  TColumn extends ColumnDef<TRow, any> = ColumnDef<TRow, any>
>(options: LocalRecordTableSessionOptions<TRow, TColumn>): RecordTableSession<TRow, TColumn> {
  return new RecordTableSession(options);
}

function validateOptions<TRow, TColumn extends ColumnDef<TRow, any>>(
  options: LocalRecordTableSessionOptions<TRow, TColumn>
): void {
  if (options.document !== undefined && options.defaultDocument !== undefined) {
    throw new Error("document and defaultDocument are mutually exclusive");
  }
  if (options.document !== undefined && !options.onDocumentChange) {
    throw new Error("Controlled document requires onDocumentChange");
  }
  if (options.state && Object.keys(options.state).length > 0 && !options.onStateChange) {
    throw new Error("Controlled state requires onStateChange");
  }
}

function mergeState(
  base: TableViewState,
  defaults?: Partial<TableViewState>,
  controlled?: Partial<TableViewState>
): TableViewState {
  return cloneState({ ...base, ...defaults, ...controlled });
}

function defaultStateForColumns<TRow>(columns: readonly ColumnDef<TRow, any>[]): TableViewState {
  return {
    ...DEFAULT_TABLE_VIEW_STATE,
    columnVisibility: Object.fromEntries(columns
      .filter((column) => column.visible === false)
      .map((column) => [column.id, false])),
    columnPinning: {
      left: columns.filter((column) => column.pin === "left").map((column) => column.id),
      right: columns.filter((column) => column.pin === "right").map((column) => column.id)
    }
  };
}

function mergeControlledState(current: TableViewState, controlled?: Partial<TableViewState>): TableViewState {
  return cloneState({ ...current, ...controlled });
}

function cloneState(state: TableViewState): TableViewState {
  return {
    ...state,
    sorting: [...state.sorting],
    grouping: [...state.grouping],
    aggregates: [...state.aggregates],
    selectedRowIds: [...state.selectedRowIds],
    expandedRowIds: [...state.expandedRowIds],
    columnOrder: [...state.columnOrder],
    columnVisibility: { ...state.columnVisibility },
    columnWidths: { ...state.columnWidths },
    columnPinning: { left: [...state.columnPinning.left], right: [...state.columnPinning.right] }
  };
}

function cloneDocument(document: TableMetadataDocument): TableMetadataDocument {
  return {
    version: 1,
    cells: Object.fromEntries(Object.entries(document.cells).map(([key, value]) => [key, cloneCellMetadata(value)])),
    calculatedColumns: document.calculatedColumns.map((item) => ({ ...item })),
    namedStyles: document.namedStyles.map((item) => ({ ...item, format: { ...item.format } }))
  };
}

function cloneCellMetadata(metadata: TableCellMetadata): TableCellMetadata {
  return {
    ...metadata,
    ...(metadata.format ? { format: { ...metadata.format } } : {}),
    ...(metadata.validation
      ? {
          validation: metadata.validation.kind === "list"
            ? { ...metadata.validation, values: [...metadata.validation.values] }
            : { ...metadata.validation }
        }
      : {})
  };
}

function queryFromState(state: TableViewState): QueryRequest {
  return {
    sorting: state.sorting,
    filter: state.filter,
    grouping: state.grouping,
    aggregates: state.aggregates,
    pagination: state.pagination,
    tree: { expandedRowIds: state.expandedRowIds }
  };
}

function sanitizeLocalViewState<TRow>(
  state: TableViewState,
  columns: readonly ColumnDef<TRow, any>[],
  treeSource: boolean
): TableViewState {
  const columnIds = new Set(columns.map((column) => column.id));
  const pagination = state.pagination.kind === "cursor"
    || state.pagination.kind === "infinite"
    || (treeSource && state.pagination.kind !== "none")
    ? { kind: "none" } as const
    : state.pagination;
  return {
    ...state,
    sorting: state.sorting.filter((sort) => columnIds.has(sort.columnId)),
    filter: filterUsesKnownColumns(state.filter, columnIds) ? state.filter : null,
    grouping: treeSource
      ? []
      : state.grouping.filter((grouping) => columnIds.has(grouping.columnId)),
    aggregates: state.aggregates.filter((aggregate) => columnIds.has(aggregate.columnId)),
    pagination
  };
}

function pruneDeletedRowsFromState(
  state: TableViewState,
  deleted: ReadonlySet<string>
): TableViewState {
  const selection = state.selection
    && (deleted.has(state.selection.anchor.rowId) || deleted.has(state.selection.focus.rowId))
    ? null
    : state.selection;
  return {
    ...state,
    selection,
    selectedRowIds: state.selectedRowIds.filter((rowId) => !deleted.has(rowId)),
    expandedRowIds: state.expandedRowIds.filter((rowId) => !deleted.has(rowId))
  };
}

function filterUsesKnownColumns(
  filter: QueryRequest["filter"],
  columnIds: ReadonlySet<string>
): boolean {
  if (!filter) return true;
  if (filter.kind === "logical") {
    return filter.operands.every((operand) => filterUsesKnownColumns(operand, columnIds));
  }
  if (filter.kind === "not") return filterUsesKnownColumns(filter.operand, columnIds);
  return columnIds.has(filter.columnId);
}

function isQueryStateIntent<TRow>(intent: TableIntent<TRow>): boolean {
  return intent.type === "set-sorting"
    || intent.type === "set-filter"
    || intent.type === "set-grouping"
    || intent.type === "set-aggregates"
    || intent.type === "set-pagination";
}

function intentFeature<TRow>(intent: TableIntent<TRow>) {
  switch (intent.type) {
    case "edit-cells": case "clear-cells": case "insert-rows": case "delete-rows": return "edit" as const;
    case "update-cell-metadata": return "metadata" as const;
    case "set-sorting": return "sort" as const;
    case "set-filter": return "filter" as const;
    case "set-grouping": return "group" as const;
    case "set-aggregates": return "aggregate" as const;
    case "set-pagination": return "pagination" as const;
    case "undo": case "redo": return "undo" as const;
    default: return null;
  }
}

function editIntent<TRow>(
  reason: "edit-cells" | "clear-cells",
  edits: readonly { rowId: string; columnId: string; rawText: string }[]
): TableIntent<TRow> {
  return reason === "edit-cells"
    ? { type: "edit-cells", edits }
    : { type: "clear-cells", cells: edits.map(({ rowId, columnId }) => ({ rowId, columnId })) };
}

function compatibleValue(value: unknown, dataType: AnyColumn<unknown>["dataType"]): boolean {
  if (value === null || dataType === undefined || dataType === "custom") return true;
  if (dataType === "date" || dataType === "datetime") return typeof value === "string";
  if (dataType === "number") return typeof value === "number" && Number.isFinite(value);
  if (dataType === "boolean") return typeof value === "boolean";
  return typeof value === "string";
}

function normalizeLocalValue(
  value: unknown,
  dataType: AnyColumn<unknown>["dataType"]
): { ok: true; value: unknown } | { ok: false } {
  if (value === null || (dataType !== "date" && dataType !== "datetime")) {
    return { ok: true, value };
  }

  const serial = typeof value === "number"
    ? value
    : typeof value === "string"
      ? parseExcelTemporalInput(value)?.serial
      : undefined;
  if (serial === undefined || serial === 60) return { ok: false };
  const date = excelSerialToDate(serial);
  if (!date) return { ok: false };
  return {
    ok: true,
    value: dataType === "date" ? date.toISOString().slice(0, 10) : date.toISOString()
  };
}

function validateMetadataValue(value: unknown, validation: TableCellMetadata["validation"], rowId: string, columnId: string): TableCellIssue[] {
  if (!validation) return [];
  if ((value === null || value === "") && validation.allowBlank !== false) return [];
  let valid = true;
  if (validation.kind === "list") valid = typeof value === "string" && validation.values.includes(value);
  else if (validation.kind === "number") valid = typeof value === "number"
    && (validation.min === undefined || value >= validation.min)
    && (validation.max === undefined || value <= validation.max);
  else valid = typeof value === "string"
    && (validation.min === undefined || value.length >= validation.min)
    && (validation.max === undefined || value.length <= validation.max);
  return valid ? [] : [{ code: "TABLE_METADATA_VALIDATION", message: "Cell value does not satisfy validation", rowId, columnId }];
}

function patchFormula(document: TableMetadataDocument, key: string, formula: string | undefined): TableMetadataDocument {
  const current = { ...(document.cells[key] ?? {}) };
  if (formula) current.formula = formula;
  else delete current.formula;
  const cells = { ...document.cells };
  if (Object.keys(current).length === 0) delete cells[key];
  else cells[key] = current;
  return { ...document, cells };
}

function findRowIndex<TRow>(rows: readonly TRow[], rowId: string, getRowId: (row: TRow) => string): number {
  return rows.findIndex((row) => getRowId(row).trim() === rowId);
}

function replaceAt<T>(values: readonly T[], index: number, value: T): readonly T[] {
  const next = [...values];
  next[index] = value;
  return next;
}

function createHistoryOperation<TRow>(
  beforeRows: readonly TRow[],
  afterRows: readonly TRow[],
  beforeDocument: TableMetadataDocument,
  afterDocument: TableMetadataDocument,
  getRowId: (row: TRow) => string
): LocalSessionOperation<TRow> {
  const beforeById = new Map(beforeRows.map((row) => [getRowId(row).trim(), row]));
  const afterById = new Map(afterRows.map((row) => [getRowId(row).trim(), row]));
  const inserted = afterRows.flatMap((row, index) => {
    const rowId = getRowId(row).trim();
    return beforeById.has(rowId) ? [] : [{ rowId, row, index }];
  });
  const deleted = beforeRows.flatMap((row) => {
    const rowId = getRowId(row).trim();
    return afterById.has(rowId) ? [] : [rowId];
  });
  const replaced = afterRows.flatMap((row) => {
    const rowId = getRowId(row).trim();
    const previous = beforeById.get(rowId);
    return beforeById.has(rowId) && !Object.is(previous, row)
      ? [{ rowId, previous: previous as TRow, row }]
      : [];
  });

  let rows: RowHistoryOperation<TRow> = { kind: "none" };
  if (inserted.length > 0) rows = { kind: "insert", records: inserted };
  else if (deleted.length > 0) rows = { kind: "delete", rowIds: deleted };
  else if (replaced.length > 0) rows = { kind: "replace", records: replaced };

  const metadata: MetadataHistoryOperation[] = [];
  for (const key of new Set([...Object.keys(beforeDocument.cells), ...Object.keys(afterDocument.cells)])) {
    const before = beforeDocument.cells[key];
    const after = afterDocument.cells[key];
    if (JSON.stringify(before) !== JSON.stringify(after)) {
      metadata.push({ key, value: after ? cloneCellMetadata(after) : null });
    }
  }
  return { rows, metadata };
}

function applyHistoryOperation<TRow>(
  currentRows: readonly TRow[],
  currentDocument: TableMetadataDocument,
  operation: LocalSessionOperation<TRow>,
  getRowId: (row: TRow) => string
): { rows: readonly TRow[]; document: TableMetadataDocument } {
  let rows = [...currentRows];
  if (operation.rows.kind === "replace") {
    const replacements = new Map(operation.rows.records.map((record) => [record.rowId, record]));
    rows = rows.map((row) => {
      const replacement = replacements.get(getRowId(row).trim());
      return replacement ? mergeChangedFields(row, replacement.previous, replacement.row) : row;
    });
  } else if (operation.rows.kind === "delete") {
    const deleted = new Set(operation.rows.rowIds);
    rows = rows.filter((row) => !deleted.has(getRowId(row).trim()));
  } else if (operation.rows.kind === "insert") {
    const present = new Set(rows.map((row) => getRowId(row).trim()));
    for (const record of [...operation.rows.records].sort((left, right) => left.index - right.index)) {
      if (present.has(record.rowId)) continue;
      rows.splice(Math.min(record.index, rows.length), 0, record.row);
      present.add(record.rowId);
    }
  }

  if (operation.metadata.length === 0) return { rows, document: currentDocument };
  const cells = { ...currentDocument.cells };
  for (const change of operation.metadata) {
    if (change.value === null) delete cells[change.key];
    else cells[change.key] = cloneCellMetadata(change.value);
  }
  return { rows, document: { ...currentDocument, cells } };
}

function rowsUpdater<TRow>(before: readonly TRow[], after: readonly TRow[], getRowId: (row: TRow) => string): RowUpdater<TRow> {
  const beforeById = new Map(before.map((row) => [getRowId(row), row]));
  const afterById = new Map(after.map((row) => [getRowId(row), row]));
  const deletedIds = new Set([...beforeById.keys()].filter((id) => !afterById.has(id)));
  const inserted = after.filter((row) => !beforeById.has(getRowId(row)));
  const changes = new Map<string, { replacement: TRow; patch?: Readonly<Record<string, unknown>> }>();

  for (const replacement of after) {
    const id = getRowId(replacement);
    const original = beforeById.get(id);
    if (original === undefined || Object.is(original, replacement)) continue;

    if (isPlainRecord(original) && isPlainRecord(replacement)) {
      const patch: Record<string, unknown> = {};
      for (const key of new Set([...Object.keys(original), ...Object.keys(replacement)])) {
        if (!Object.is(original[key], replacement[key])) patch[key] = replacement[key];
      }
      changes.set(id, { replacement, patch });
    } else {
      changes.set(id, { replacement });
    }
  }

  return (previous) => {
    const next = previous.filter((row) => !deletedIds.has(getRowId(row))).map((row) => {
    const change = changes.get(getRowId(row));
    if (!change) return row;
    return change.patch && isPlainRecord(row)
      ? { ...row, ...change.patch } as TRow
      : change.replacement;
    });

    const present = new Set(next.map(getRowId));
    const desiredIds = after.map(getRowId);
    for (const row of inserted) {
      const rowId = getRowId(row);
      if (present.has(rowId)) continue;
      const desiredIndex = desiredIds.indexOf(rowId);
      const previousId = [...desiredIds.slice(0, desiredIndex)].reverse().find((id) => present.has(id));
      const nextId = desiredIds.slice(desiredIndex + 1).find((id) => present.has(id));
      const index = previousId !== undefined
        ? next.findIndex((candidate) => getRowId(candidate) === previousId) + 1
        : nextId !== undefined
          ? next.findIndex((candidate) => getRowId(candidate) === nextId)
          : next.length;
      next.splice(Math.max(0, index), 0, row);
      present.add(rowId);
    }
    return next;
  };
}

function isPlainRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value) && !(value instanceof Date);
}

function mergeChangedFields<TRow>(current: TRow, previous: TRow, next: TRow): TRow {
  if (!isPlainRecord(current) || !isPlainRecord(previous) || !isPlainRecord(next)) return next;
  const merged: Record<string, unknown> = { ...current };
  for (const key of new Set([...Object.keys(previous), ...Object.keys(next)])) {
    if (Object.is(previous[key], next[key])) continue;
    if (Object.prototype.hasOwnProperty.call(next, key)) merged[key] = next[key];
    else delete merged[key];
  }
  return merged as TRow;
}

function rowsEqualByReference<T>(left: readonly T[], right: readonly T[]): boolean {
  return left.length === right.length && left.every((value, index) => Object.is(value, right[index]));
}

function firstDuplicate(values: readonly string[]): string | null {
  const seen = new Set<string>();
  for (const value of values) {
    if (seen.has(value)) return value;
    seen.add(value);
  }
  return null;
}

function safeParseMetadataKey(key: string): { rowId: string; columnId: string } | null {
  try {
    return parseTableMetadataKey(key);
  } catch {
    return null;
  }
}

function documentUpdater(before: TableMetadataDocument, after: TableMetadataDocument): TableMetadataUpdater {
  const keys = [...new Set([...Object.keys(before.cells), ...Object.keys(after.cells)])]
    .filter((key) => JSON.stringify(before.cells[key]) !== JSON.stringify(after.cells[key]));
  const calculatedColumnsChanged = JSON.stringify(before.calculatedColumns) !== JSON.stringify(after.calculatedColumns);
  const namedStylesChanged = JSON.stringify(before.namedStyles) !== JSON.stringify(after.namedStyles);
  return (previous) => {
    const cells = { ...previous.cells };
    for (const key of keys) {
      if (key in after.cells) cells[key] = cloneCellMetadata(after.cells[key]);
      else delete cells[key];
    }
    return {
      ...previous,
      cells,
      ...(calculatedColumnsChanged
        ? { calculatedColumns: after.calculatedColumns.map((column) => ({ ...column })) }
        : {}),
      ...(namedStylesChanged
        ? { namedStyles: after.namedStyles.map((style) => ({ ...style, format: { ...style.format } })) }
        : {})
    };
  };
}

function documentEqual(left: TableMetadataDocument, right: TableMetadataDocument): boolean {
  return JSON.stringify(left) === JSON.stringify(right);
}

function stateEqual(left: TableViewState, right: TableViewState): boolean {
  return JSON.stringify(left) === JSON.stringify(right);
}

function stateSliceEqual(left: unknown, right: unknown): boolean {
  return JSON.stringify(left) === JSON.stringify(right);
}

function applyControlledStateDiff(
  state: TableViewState,
  before: TableViewState,
  after: TableViewState,
  controlled: ReadonlySet<keyof TableViewState>
): TableViewState {
  let next = state;
  for (const key of controlled) {
    if (!stateSliceEqual(before[key], after[key])) next = { ...next, [key]: after[key] };
  }
  return next;
}

function pinColumn(state: TableViewState, columnId: string, pin: "left" | "right" | false): TableViewState {
  const left = state.columnPinning.left.filter((id) => id !== columnId);
  const right = state.columnPinning.right.filter((id) => id !== columnId);
  if (pin === "left") left.push(columnId);
  if (pin === "right") right.push(columnId);
  return { ...state, columnPinning: { left, right } };
}

function hasGroupingPaginationConflict(state: TableViewState): boolean {
  return state.grouping.length > 0 && state.pagination.kind !== "none";
}

function sameColumnReferences<TRow>(left: readonly ColumnDef<TRow, any>[], right: readonly ColumnDef<TRow, any>[]): boolean {
  return left.length === right.length && left.every((column, index) => {
    const candidate = right[index];
    const columnRecord = column as unknown as Readonly<Record<string, unknown>>;
    const candidateRecord = candidate as unknown as Readonly<Record<string, unknown>>;
    const keys = Object.keys(columnRecord);
    return keys.length === Object.keys(candidateRecord).length
      && keys.every((key) => Object.prototype.hasOwnProperty.call(candidateRecord, key)
        && Object.is(columnRecord[key], candidateRecord[key]));
  });
}

function rowIdList<TRow>(rows: readonly TRow[], getRowId: (row: TRow) => string): string[] {
  return rows.map((row) => getRowId(row).trim());
}

function dedupeIssues(issues: readonly TableCellIssue[]): TableCellIssue[] {
  return [...new Map(issues.map((issue) => [JSON.stringify([issue.code, issue.rowId ?? "", issue.columnId ?? ""]), issue])).values()];
}

function formatDefault(value: unknown): string {
  if (isErrorValue(value)) return value.code;
  if (value === null || value === undefined) return "";
  if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
  return String(value);
}

function isErrorValue(value: unknown): value is { kind: "error"; code: string } {
  return typeof value === "object" && value !== null && "kind" in value && value.kind === "error" && "code" in value;
}

function groupOrAggregateCell<TRow>(
  rows: readonly QueryRow<TRow>[],
  rowId: string,
  columnId: string,
  aggregateRequests: readonly TableAggregateRequest[]
) {
  const row = rows.find((candidate) => candidate.id === rowId);
  const aggregateId = aggregateRequests.find((request) => request.columnId === columnId)?.id;
  let value: unknown = null;
  if (row?.kind === "group") {
    value = row.columnId === columnId
      ? row.key.type === "null" ? null : row.key.value
      : aggregateId === undefined ? null : row.aggregates[aggregateId];
  }
  if (row?.kind === "aggregate") value = aggregateId === undefined ? null : row.aggregates[aggregateId];
  return {
    rowId,
    columnId,
    storedValue: value,
    evaluatedValue: value,
    displayValue: formatDefault(value),
    metadata: {},
    editable: false,
    issues: []
  };
}

function exportColumns<TRow, TColumn extends ColumnDef<TRow, any>>(
  columns: readonly TColumn[],
  state: TableViewState
): TColumn[] {
  const byId = new Map(columns.map((column) => [column.id, column]));
  const ordered = [...state.columnOrder, ...columns.map((column) => column.id).filter((id) => !state.columnOrder.includes(id))];
  return ordered.flatMap((id) => {
    const column = byId.get(id);
    return column && state.columnVisibility[id] !== false && column.kind !== "display" ? [column] : [];
  });
}

function csvEscape(value: string): string {
  return /[",\r\n]/.test(value) ? `"${value.replace(/"/g, '""')}"` : value;
}

function formatExportText(value: unknown): string {
  return formatDefault(value);
}

function exportFileName(fileName: string | undefined, format: ExportOptions["format"]): string {
  const trimmed = fileName?.trim() || "table";
  return trimmed.toLowerCase().endsWith(`.${format}`) ? trimmed : `${trimmed}.${format}`;
}

function toExcelValue(value: unknown): ExcelJS.CellValue {
  if (value === null || value === undefined || typeof value === "string" || typeof value === "number" || typeof value === "boolean" || value instanceof Date) {
    return value ?? null;
  }
  return formatDefault(value);
}

function now(): number {
  return typeof globalThis.performance?.now === "function" ? globalThis.performance.now() : Date.now();
}
