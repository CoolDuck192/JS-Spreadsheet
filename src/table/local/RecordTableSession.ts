import ExcelJS from "exceljs";
import type { CommandResult } from "../../core/commands/types";
import { parseCellInput } from "../../core/values/parseCellInput";
import { createLocalTableCapabilities, resolveTableOperationStates, type TableFeatureConfiguration } from "../core/capabilities";
import { createCommandIdFactory, type CommandIdFactory } from "../core/commandId";
import { normalizeColumns } from "../core/columnHelper";
import type { PaginationRequest, QueryRequest, QueryRow } from "../core/query";
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
import { createTableMetadataKey } from "./tableMetadata";

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
type HistoryEntry<TRow> = {
  beforeRows: readonly TRow[];
  afterRows: readonly TRow[];
  beforeDocument: TableMetadataDocument;
  afterDocument: TableMetadataDocument;
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
  private historyPast: HistoryEntry<TRow>[] = [];
  private historyFuture: HistoryEntry<TRow>[] = [];
  private invalidControlledState = false;
  private sessionIssues: TableCellIssue[] = [];

  constructor(options: LocalRecordTableSessionOptions<TRow, TColumn>) {
    validateOptions(options);
    this.options = options;
    this.columns = normalizeColumns(options.columns) as readonly TColumn[];
    this.columnsById = new Map(this.columns.map((column) => [column.id, column]));
    this.rows = [...options.source.rows];
    this.document = cloneDocument(options.document ?? options.defaultDocument ?? EMPTY_TABLE_DOCUMENT);
    this.state = mergeState(DEFAULT_TABLE_VIEW_STATE, options.defaultState, options.state);
    this.controlledRows = Boolean(options.source.onRowsChange);
    this.controlledDocument = options.document !== undefined;
    this.controlledStateKeys = new Set(Object.keys(options.state ?? {}) as (keyof TableViewState)[]);
    this.sourceRowsReference = options.source.rows;
    this.resetKey = options.source.resetKey;
    this.commandIdFactory = options.commandIdFactory ?? createCommandIdFactory();
    this.validateHistoryLimit(options.historyLimit);
    this.validateRows(this.rows);
    this.handleInvalidControlledState();
  }

  updateOptions(options: LocalRecordTableSessionOptions<TRow, TColumn>): void {
    validateOptions(options);
    this.validateHistoryLimit(options.historyLimit);
    const previousControlledRows = this.controlledRows;
    const nextControlledRows = Boolean(options.source.onRowsChange);
    const modeChanged = previousControlledRows !== nextControlledRows;
    let externalChanged = false;

    if (modeChanged) {
      this.rows = [...options.source.rows];
      this.clearHistory();
      this.dropInvalidSelection(options.source, this.rows);
      externalChanged = true;
    } else if (nextControlledRows && options.source.rows !== this.sourceRowsReference) {
      const sameIdentityOrder = rowIdList(this.rows, options.source.getRowId).join("\u0000")
        === rowIdList(options.source.rows, options.source.getRowId).join("\u0000");
      this.rows = [...options.source.rows];
      if (!sameIdentityOrder) {
        this.clearHistory();
        this.dropInvalidSelection(options.source, this.rows);
      }
      externalChanged = true;
    } else if (!nextControlledRows && options.source.resetKey !== this.resetKey) {
      this.rows = [...options.source.rows];
      this.clearHistory();
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
      this.clearHistory();
      externalChanged = true;
    } else if (nextControlledDocument && options.document !== this.options.document) {
      this.document = cloneDocument(options.document!);
      externalChanged = true;
    }

    const nextStateKeys = new Set(Object.keys(options.state ?? {}) as (keyof TableViewState)[]);
    const controlledCandidate = mergeControlledState(this.state, options.state);
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
      if (!stateEqual(this.state, controlledCandidate)) {
        this.state = controlledCandidate;
        externalChanged = true;
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
      canUndo: this.historyPast.length > 0,
      canRedo: this.historyFuture.length > 0,
      pageInfo: model.pageInfo,
      getCell: (rowId, columnId) => {
        const column = this.columnsById.get(columnId);
        if (!column) throw new Error(`Unknown column id: ${columnId}`);
        const row = model.dataRowsById.get(rowId);
        if (!row) return groupOrAggregateCell(snapshotRows, rowId, columnId);
        const evaluation = evaluate(row, rowId, columnId);
        const metadata = this.document.cells[createTableMetadataKey(rowId, columnId)] ?? {};
        const context = this.columnContext(row, rowId, columnId, evaluate);
        const formatted = evaluation.displayValue === undefined && column.format && !isErrorValue(evaluation.evaluatedValue)
          ? safeInvokeTableExtension("format", () => column.format!(evaluation.evaluatedValue, context))
          : null;
        const formatIssue = formatted && !formatted.ok ? [{ ...formatted.issue, rowId, columnId }] : [];
        return {
          rowId,
          columnId,
          storedValue: evaluation.storedValue,
          evaluatedValue: evaluation.evaluatedValue,
          displayValue: evaluation.displayValue
            ?? (formatted?.ok ? formatted.value : formatDefault(evaluation.evaluatedValue)),
          ...(metadata.formula ? { formula: metadata.formula } : {}),
          metadata,
          editable: this.isEditable(column, context, metadata).editable,
          issues: dedupeIssues([...evaluation.issues, ...formatIssue])
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
    if (intent.type === "undo") return this.replayHistory("undo", commandId, startedAt);
    if (intent.type === "redo") return this.replayHistory("redo", commandId, startedAt);
    if (intent.type === "refresh") return this.finishNoChange(commandId, intent, startedAt);

    const feature = intentFeature(intent);
    if (feature && !this.getSnapshot().operationStates[feature].enabled) {
      return this.rejectUnsupported(commandId, intent, startedAt, this.getSnapshot().operationStates[feature].reason);
    }

    if (intent.type === "edit-cells" || intent.type === "clear-cells") {
      const edits = intent.type === "edit-cells"
        ? intent.edits
        : intent.cells.map((cell) => ({ ...cell, rawText: "" }));
      return this.editCells(edits, intent.type, commandId, startedAt);
    }
    if (intent.type === "update-cell-metadata") {
      return this.updateMetadata(intent.updates, commandId, startedAt);
    }
    if (intent.type === "insert-rows" || intent.type === "delete-rows") {
      return this.rejectUnsupported(commandId, intent, startedAt, "Structural row edits are not implemented");
    }

    return this.updateViewState(intent, commandId, startedAt);
  }

  async refresh(): Promise<void> {
    const commandId = this.commandIdFactory();
    this.emitDiagnostic(commandId, "refresh", now(), false, "command", 0, 0);
  }

  async undo(): Promise<CommandResult> {
    const commandId = this.commandIdFactory();
    return this.replayHistory("undo", commandId, now());
  }

  async redo(): Promise<CommandResult> {
    const commandId = this.commandIdFactory();
    return this.replayHistory("redo", commandId, now());
  }

  async export(options: ExportOptions): Promise<ExportArtifact> {
    const commandId = this.commandIdFactory();
    const startedAt = now();
    if (this.destroyed) throw new Error("Table session is destroyed");
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
    let candidateRows = this.rows;
    let candidateDocument = this.document;
    let changed = false;
    for (const edit of edits) {
      const rowIndex = findRowIndex(candidateRows, edit.rowId, this.options.source.getRowId);
      if (rowIndex < 0) return this.reject("validation", [{ code: "TABLE_ROW_NOT_FOUND", message: "Row not found", rowId: edit.rowId }], commandId, reason, startedAt, edits.length, edits.length);
      const row = candidateRows[rowIndex];
      const column = this.columnsById.get(edit.columnId);
      if (!column) return this.reject("validation", [{ code: "TABLE_COLUMN_NOT_FOUND", message: "Column not found", rowId: edit.rowId, columnId: edit.columnId }], commandId, reason, startedAt, edits.length, edits.length);
      const context = this.columnContext(row, edit.rowId, edit.columnId);
      const metadataKey = createTableMetadataKey(edit.rowId, edit.columnId);
      const metadata = candidateDocument.cells[metadataKey] ?? {};
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
        if (!this.options.formulaService) return this.rejectUnsupported(commandId, { type: reason, ...(reason === "edit-cells" ? { edits } : { cells: edits }) } as TableIntent<TRow>, startedAt, "Formula service is not configured");
        const formulaResult = this.invokeFormula(formula, row, edit.rowId, edit.columnId);
        if ("issues" in formulaResult) return this.reject("validation", formulaResult.issues, commandId, reason, startedAt, edits.length, edits.length);
        value = formulaResult.value;
        evaluated = formulaResult.value;
      } else {
        value = input.stored;
        evaluated = value;
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
      if (!("update" in column) || typeof column.update !== "function") {
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
      const updated = safeInvokeTableExtension("update", () => column.update!(row, value));
      if (!updated.ok) return this.reject("unsupported", [{ ...updated.issue, rowId: edit.rowId, columnId: edit.columnId }], commandId, reason, startedAt, edits.length, edits.length);
      candidateRows = replaceAt(candidateRows, rowIndex, updated.value);
      candidateDocument = patchFormula(candidateDocument, metadataKey, formula);
      changed = true;
    }
    if (!changed) return this.finishNoChange(commandId, { type: reason } as TableIntent<TRow>, startedAt);
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
    cellCount: number
  ): Promise<CommandResult> {
    const entry: HistoryEntry<TRow> = {
      beforeRows: this.rows,
      afterRows: rows,
      beforeDocument: this.document,
      afterDocument: document
    };
    this.rows = rows;
    this.document = document;
    this.historyPast.push(entry);
    const limit = this.options.historyLimit ?? 100;
    if (this.historyPast.length > limit) this.historyPast.splice(0, this.historyPast.length - limit);
    this.historyFuture = [];
    this.revision += 1;
    this.invalidate();
    const context = this.context(commandId, reason);
    if (this.controlledRows) this.safeHostCallback(() => this.options.source.onRowsChange?.(rowsUpdater(entry.beforeRows, rows, this.options.source.getRowId), context));
    if (this.controlledDocument && !documentEqual(entry.beforeDocument, document)) {
      this.safeHostCallback(() => this.options.onDocumentChange?.(documentUpdater(entry.beforeDocument, document), context));
    }
    this.publish();
    this.emitDiagnostic(commandId, reason, startedAt, true, "command", rowCount, cellCount);
    return { status: "committed", revision: String(this.revision), changed: true };
  }

  private async replayHistory(kind: "undo" | "redo", commandId: string, startedAt: number): Promise<CommandResult> {
    if (this.destroyed) return this.rejectUnsupported(commandId, { type: kind } as TableIntent<TRow>, startedAt, "Session is destroyed");
    const source = kind === "undo" ? this.historyPast : this.historyFuture;
    const entry = source.pop();
    if (!entry) return this.finishNoChange(commandId, { type: kind } as TableIntent<TRow>, startedAt);
    if (kind === "undo") this.historyFuture.push(entry);
    else this.historyPast.push(entry);
    const beforeRows = this.rows;
    const beforeDocument = this.document;
    this.rows = kind === "undo" ? entry.beforeRows : entry.afterRows;
    this.document = kind === "undo" ? entry.beforeDocument : entry.afterDocument;
    this.revision += 1;
    this.invalidate();
    const context = this.context(commandId, kind);
    if (this.controlledRows) this.safeHostCallback(() => this.options.source.onRowsChange?.(rowsUpdater(beforeRows, this.rows, this.options.source.getRowId), context));
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
    const projected = scope === "currentView"
      ? snapshot.rows
      : this.buildUnpaginatedModel().items;
    const dataRows = projected.filter((row): row is Extract<QueryRow<TRow>, { kind: "data" }> => row.kind === "data");
    return {
      columns: orderedColumns.map((column) => ({ id: column.id, header: typeof column.header === "string" ? column.header : column.id })),
      rows: dataRows.map((row) => orderedColumns.map((column) => snapshot.getCell(row.id, column.id).evaluatedValue))
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
    if (!this.options.onStateChange) return;
    const corrected = { ...candidate, pagination: { kind: "none" } as const };
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

  private clearHistory(): void {
    this.historyPast = [];
    this.historyFuture = [];
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
    cells: Object.fromEntries(Object.entries(document.cells).map(([key, value]) => [key, { ...value }])),
    calculatedColumns: document.calculatedColumns.map((item) => ({ ...item })),
    namedStyles: document.namedStyles.map((item) => ({ ...item, format: { ...item.format } }))
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

function intentFeature<TRow>(intent: TableIntent<TRow>) {
  switch (intent.type) {
    case "edit-cells": case "clear-cells": return "edit" as const;
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

function compatibleValue(value: unknown, dataType: AnyColumn<unknown>["dataType"]): boolean {
  if (value === null || dataType === undefined || dataType === "custom") return true;
  if (dataType === "number" || dataType === "date" || dataType === "datetime") return typeof value === "number" && Number.isFinite(value);
  if (dataType === "boolean") return typeof value === "boolean";
  return typeof value === "string";
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

function rowsUpdater<TRow>(before: readonly TRow[], after: readonly TRow[], getRowId: (row: TRow) => string): RowUpdater<TRow> {
  const beforeById = new Map(before.map((row) => [getRowId(row), row]));
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

  return (previous) => previous.map((row) => {
    const change = changes.get(getRowId(row));
    if (!change) return row;
    return change.patch && isPlainRecord(row)
      ? { ...row, ...change.patch } as TRow
      : change.replacement;
  });
}

function isPlainRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value) && !(value instanceof Date);
}

function documentUpdater(before: TableMetadataDocument, after: TableMetadataDocument): TableMetadataUpdater {
  const keys = new Set([...Object.keys(before.cells), ...Object.keys(after.cells)]);
  return (previous) => {
    const cells = { ...previous.cells };
    for (const key of keys) {
      if (key in after.cells) cells[key] = after.cells[key];
      else delete cells[key];
    }
    return { ...previous, cells, calculatedColumns: after.calculatedColumns, namedStyles: after.namedStyles };
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
    const leftAccessor = "accessor" in column ? column.accessor : undefined;
    const rightAccessor = "accessor" in candidate ? candidate.accessor : undefined;
    const leftUpdate = "update" in column ? column.update : undefined;
    const rightUpdate = "update" in candidate ? candidate.update : undefined;
    const leftCalculate = "calculate" in column ? column.calculate : undefined;
    const rightCalculate = "calculate" in candidate ? candidate.calculate : undefined;
    return column.id === candidate.id
      && leftAccessor === rightAccessor
      && leftUpdate === rightUpdate
      && leftCalculate === rightCalculate;
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

function groupOrAggregateCell<TRow>(rows: readonly QueryRow<TRow>[], rowId: string, columnId: string) {
  const row = rows.find((candidate) => candidate.id === rowId);
  let value: unknown = null;
  if (row?.kind === "group") value = row.columnId === columnId ? row.key.type === "null" ? null : row.key.value : row.aggregates[columnId];
  if (row?.kind === "aggregate") value = row.aggregates[columnId];
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
