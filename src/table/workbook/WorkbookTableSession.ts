import type { CommandResult } from "../../core/commands/types";
import type { WorkbookSession } from "../../core/workbook/WorkbookSession";
import type { WorkbookCommand } from "../../core/workbook/commands";
import { getStructuredTableBodyRange } from "../../core/workbook/structuredTables";
import { isStructuredTableRowVisible } from "../../core/workbook/structuredTableFilter";
import {
  getCellComment,
  getCellContent,
  getCellFormat,
  getCellReadOnly,
  getCellValidation,
  getColumnWidth,
  isColumnHidden
} from "../../lib/workbook";
import { formatCellAddress } from "../../lib/addressing";
import { formatDisplayValue } from "../../lib/displayFormat";
import {
  resolveTableOperationStates,
  type TableCapabilities
} from "../core/capabilities";
import type { QueryRow } from "../core/query";
import type {
  ColumnDef,
  ExportArtifact,
  ExportOptions,
  TableCellMetadata,
  TableCellMetadataUpdate,
  TableCellRef,
  TableCellSnapshot,
  TableIntent,
  TableSelection,
  TableSession,
  TableViewSnapshot,
  TableViewState
} from "../core/types";
import type {
  CellFormat,
  DataValidationRule,
  StructuredTable,
  StructuredTableColumn
} from "../../types";

export type WorkbookTableRow = Readonly<{
  id: string;
  tableId: string;
  sheetId: string;
  sheetRow: number;
}>;

export interface WorkbookTableSession extends TableSession<WorkbookTableRow> {
  readonly tableId: string;
}

const WORKBOOK_TABLE_CAPABILITIES = Object.freeze({
  sort: { executor: "client", scope: "completeDataset" },
  filter: { executor: "client", scope: "completeDataset" },
  group: false,
  aggregate: false,
  pagination: false,
  edit: { executor: "client", scope: "completeDataset" },
  bulkEdit: { executor: "client", scope: "completeDataset" },
  metadata: { executor: "client", scope: "completeDataset" },
  validation: { executor: "client", scope: "completeDataset" },
  formula: "fullLocalDataset",
  subscription: true,
  undo: { executor: "client", scope: "completeDataset" },
  export: false
} satisfies TableCapabilities);

const OPERATION_STATES = resolveTableOperationStates(WORKBOOK_TABLE_CAPABILITIES, {});

export function createWorkbookTableSession(
  workbookSession: WorkbookSession,
  tableId: string,
  onDestroy?: () => void
): WorkbookTableSession {
  let destroyed = false;
  let localRevision = 0;
  let cachedParentRevision = "";
  let cachedSnapshot: TableViewSnapshot<WorkbookTableRow> | null = null;
  let selectedRowIds: readonly string[] = [];
  let expandedRowIds: readonly string[] = [];
  let localSelection: TableSelection | null | undefined;
  let lastWorksheetSelection = workbookSession.getSnapshot().selection;
  let localColumnOrder: readonly string[] | null = null;
  let localPinning: { left: readonly string[]; right: readonly string[] } = { left: [], right: [] };
  const columnDefinitions = new Map<string, {
    column: StructuredTableColumn;
    width: number;
    definition: ColumnDef<WorkbookTableRow>;
  }>();
  const listeners = new Set<() => void>();

  const parentUnsubscribe = workbookSession.subscribe(() => {
    const nextWorksheetSelection = workbookSession.getSnapshot().selection;
    if (nextWorksheetSelection !== lastWorksheetSelection) localSelection = undefined;
    lastWorksheetSelection = nextWorksheetSelection;
    cachedSnapshot = null;
    publish();
  });

  const session: WorkbookTableSession = {
    tableId,
    getSnapshot,
    subscribe(listener) {
      if (destroyed) return () => {};
      listeners.add(listener);
      return () => listeners.delete(listener);
    },
    async dispatch(intent) {
      if (destroyed) return unsupported("Workbook table session is destroyed");
      const parent = workbookSession.getSnapshot();
      const table = parent.workbook.tables.find((candidate) => candidate.id === tableId);
      if (!table) return rejected("validation", "TABLE_NOT_FOUND", "Structured table does not exist");

      switch (intent.type) {
        case "edit-cells":
          return workbookSession.dispatch({ type: "table.editCells", tableId, edits: intent.edits });
        case "clear-cells":
          return workbookSession.dispatch({
            type: "table.editCells",
            tableId,
            edits: intent.cells.map((cell) => ({ ...cell, rawText: "" }))
          });
        case "set-selection":
          if (intent.selection === null) {
            return commitLocal(() => { localSelection = null; });
          }
          localSelection = undefined;
          return dispatchSelection(table, intent.selection);
        case "set-row-selection":
          if (intent.rowIds.some((rowId) => !table.rowIds.includes(rowId))) {
            return rejected("validation", "TABLE_ROW_NOT_FOUND", "Selected table row does not exist");
          }
          return commitLocal(() => { selectedRowIds = [...intent.rowIds]; });
        case "set-row-expanded":
          if (!table.rowIds.includes(intent.rowId)) {
            return rejected("validation", "TABLE_ROW_NOT_FOUND", "Expanded table row does not exist");
          }
          return commitLocal(() => {
            const next = new Set(expandedRowIds);
            if (intent.expanded) next.add(intent.rowId);
            else next.delete(intent.rowId);
            expandedRowIds = [...next];
          });
        case "set-sorting":
          return workbookSession.dispatch({ type: "table.sort", tableId, sorting: intent.sorting });
        case "set-filter":
          return workbookSession.dispatch({
            type: "table.setFilter",
            tableId,
            ...(intent.filter === null ? {} : { filter: intent.filter })
          });
        case "set-grouping":
          return unsupported("Grouping is not supported for workbook tables");
        case "set-aggregates":
          return unsupported("Ad-hoc aggregation is not supported for workbook tables");
        case "set-pagination":
          return unsupported("Pagination is not supported for workbook tables");
        case "update-cell-metadata":
          return dispatchMetadata(table, intent.updates);
        case "set-column-order": {
          const ids = new Set(table.columns.map((column) => column.id));
          if (
            intent.columnIds.length !== ids.size
            || new Set(intent.columnIds).size !== ids.size
            || intent.columnIds.some((id) => !ids.has(id))
          ) {
            return rejected("validation", "TABLE_COLUMN_NOT_FOUND", "Column order must contain every table column exactly once");
          }
          return commitLocal(() => { localColumnOrder = [...intent.columnIds]; });
        }
        case "resize-column": {
          const column = table.columns.find((candidate) => candidate.id === intent.columnId);
          if (!column) return rejected("validation", "TABLE_COLUMN_NOT_FOUND", "Table column does not exist");
          return workbookSession.dispatch({
            type: "columns.resize",
            sheetId: table.sheetId,
            columns: [column.sheetColumn],
            width: intent.width
          });
        }
        case "set-column-visibility": {
          const column = table.columns.find((candidate) => candidate.id === intent.columnId);
          if (!column) return rejected("validation", "TABLE_COLUMN_NOT_FOUND", "Table column does not exist");
          return workbookSession.dispatch({
            type: "columns.hidden.set",
            sheetId: table.sheetId,
            columns: [column.sheetColumn],
            hidden: !intent.visible
          });
        }
        case "set-column-pinning":
          if (!table.columns.some((column) => column.id === intent.columnId)) {
            return rejected("validation", "TABLE_COLUMN_NOT_FOUND", "Table column does not exist");
          }
          return commitLocal(() => {
            const left = localPinning.left.filter((id) => id !== intent.columnId);
            const right = localPinning.right.filter((id) => id !== intent.columnId);
            if (intent.pin === "left") left.push(intent.columnId);
            if (intent.pin === "right") right.push(intent.columnId);
            localPinning = { left, right };
          });
        case "insert-rows":
          if (!("count" in intent)) return unsupported("Workbook table insertion accepts a row count, not row objects");
          return workbookSession.dispatch({
            type: "table.insertRows",
            tableId,
            count: intent.count,
            ...(intent.beforeRowId === undefined ? {} : { beforeRowId: intent.beforeRowId }),
            ...(intent.afterRowId === undefined ? {} : { afterRowId: intent.afterRowId })
          });
        case "delete-rows":
          if (intent.rowIds.some((rowId) => !table.rowIds.includes(rowId))) {
            return rejected("validation", "TABLE_ROW_NOT_FOUND", "Deleted table row does not exist");
          }
          return workbookSession.dispatch({ type: "table.deleteRows", tableId, rowIds: intent.rowIds });
        case "undo":
          return workbookSession.dispatch({ type: "history.undo" });
        case "redo":
          return workbookSession.dispatch({ type: "history.redo" });
        case "refresh":
          cachedSnapshot = null;
          return { status: "committed", revision: getSnapshot().revision, changed: false };
        case "reload-authoritative":
        case "retry-with-revision":
          return unsupported("Remote reconciliation intents are not supported for workbook tables");
      }
    },
    async refresh() {
      if (!destroyed) cachedSnapshot = null;
    },
    async undo() {
      return session.dispatch({ type: "undo" });
    },
    async redo() {
      return session.dispatch({ type: "redo" });
    },
    async export(_options: ExportOptions): Promise<ExportArtifact> {
      throw new Error("Workbook table export is not available until native XLSX table export is enabled");
    },
    destroy() {
      if (destroyed) return;
      destroyed = true;
      parentUnsubscribe();
      listeners.clear();
      cachedSnapshot = null;
      columnDefinitions.clear();
      onDestroy?.();
    }
  };

  return session;

  function getSnapshot(): TableViewSnapshot<WorkbookTableRow> {
    const parent = workbookSession.getSnapshot();
    const revision = `${parent.revision}:${localRevision}`;
    if (cachedSnapshot && cachedParentRevision === revision) return cachedSnapshot;
    cachedParentRevision = revision;
    const table = parent.workbook.tables.find((candidate) => candidate.id === tableId);
    if (!table) {
      cachedSnapshot = missingTableSnapshot(revision, parent.canUndo, parent.canRedo);
      return cachedSnapshot;
    }
    const body = getStructuredTableBodyRange(table);
    const allRows: WorkbookTableRow[] = table.rowIds.map((id, index) => ({
      id,
      tableId,
      sheetId: table.sheetId,
      sheetRow: (body?.start.row ?? table.range.start.row) + index
    }));
    const visibleRows = allRows.filter((row) => isStructuredTableRowVisible(
      parent.workbook,
      table,
      row.sheetRow,
      workbookSession.getCellEvaluation
    ));
    const rows: QueryRow<WorkbookTableRow>[] = visibleRows.map((row) => ({
      kind: "data",
      id: row.id,
      original: row,
      depth: 0
    }));
    const currentColumnIds = new Set(table.columns.map((column) => column.id));
    for (const columnId of columnDefinitions.keys()) {
      if (!currentColumnIds.has(columnId)) columnDefinitions.delete(columnId);
    }
    const columns = table.columns.map((column): ColumnDef<WorkbookTableRow> => {
      const width = getColumnWidth(parent.workbook, table.sheetId, column.sheetColumn);
      const cached = columnDefinitions.get(column.id);
      if (cached?.column === column && cached.width === width) return cached.definition;
      const definition: ColumnDef<WorkbookTableRow> = {
        id: column.id,
        header: column.name,
        dataType: column.dataType,
        accessor: (row) => getCell(row.id, column.id).evaluatedValue,
        editable: !column.calculatedFormula,
        sortable: true,
        filterable: true,
        width
      };
      columnDefinitions.set(column.id, { column, width, definition });
      return definition;
    });
    const selection = localSelection === undefined
      ? selectionFromWorkbook(table, parent.selection)
      : localSelection;
    const columnOrder = localColumnOrder?.filter((id) => table.columns.some((column) => column.id === id));
    const state: TableViewState = {
      sorting: table.sort ?? [],
      filter: table.filter ?? null,
      grouping: [],
      aggregates: [],
      pagination: { kind: "none" },
      selection,
      selectedRowIds: selectedRowIds.filter((id) => table.rowIds.includes(id)),
      expandedRowIds: expandedRowIds.filter((id) => table.rowIds.includes(id)),
      columnOrder: columnOrder?.length === table.columns.length
        ? columnOrder
        : table.columns.map((column) => column.id),
      columnVisibility: Object.fromEntries(table.columns.map((column) => [
        column.id,
        !isColumnHidden(parent.workbook, table.sheetId, column.sheetColumn)
      ])),
      columnWidths: Object.fromEntries(table.columns.map((column) => [
        column.id,
        getColumnWidth(parent.workbook, table.sheetId, column.sheetColumn)
      ])),
      columnPinning: {
        left: localPinning.left.filter((id) => table.columns.some((column) => column.id === id)),
        right: localPinning.right.filter((id) => table.columns.some((column) => column.id === id))
      }
    };
    cachedSnapshot = {
      revision,
      rows,
      columns,
      rowCount: rows.length,
      totalRowCount: { kind: "known", value: table.rowIds.length },
      completeness: "completeDataset",
      state,
      selection,
      status: { phase: "ready" },
      issues: [],
      capabilities: WORKBOOK_TABLE_CAPABILITIES,
      operationStates: OPERATION_STATES,
      pendingOperations: [],
      conflicts: [],
      canUndo: parent.canUndo,
      canRedo: parent.canRedo,
      pageInfo: { kind: "none", total: { kind: "known", value: table.rowIds.length } },
      getCell,
      getRowIndex(rowId) {
        return rows.findIndex((row) => row.id === rowId);
      },
      getColumnIndex(columnId) {
        return state.columnOrder.indexOf(columnId);
      }
    };
    return cachedSnapshot;
  }

  function getCell(rowId: string, columnId: string): TableCellSnapshot {
    const parent = workbookSession.getSnapshot();
    const table = parent.workbook.tables.find((candidate) => candidate.id === tableId);
    const body = table ? getStructuredTableBodyRange(table) : null;
    const rowIndex = table?.rowIds.indexOf(rowId) ?? -1;
    const column = table?.columns.find((candidate) => candidate.id === columnId);
    if (!table || !body || rowIndex < 0 || !column) {
      return {
        rowId,
        columnId,
        storedValue: null,
        evaluatedValue: null,
        displayValue: "",
        metadata: {},
        editable: false,
        issues: [{ code: "TABLE_NOT_FOUND", message: "Structured table cell does not exist", rowId, columnId }]
      };
    }
    const address = formatCellAddress({ row: body.start.row + rowIndex, column: column.sheetColumn });
    const storedValue = getCellContent(parent.workbook, table.sheetId, address);
    const evaluatedValue = workbookSession.getCellEvaluation(table.sheetId, address);
    const format = getCellFormat(parent.workbook, table.sheetId, address);
    const metadata = cellMetadata(parent.workbook, table, address, format, storedValue);
    const displayRaw = evaluatedValue === null
      ? ""
      : typeof evaluatedValue === "object"
        ? evaluatedValue.code
        : typeof evaluatedValue === "boolean"
          ? evaluatedValue ? "TRUE" : "FALSE"
          : String(evaluatedValue);
    return {
      rowId,
      columnId,
      storedValue,
      evaluatedValue,
      displayValue: formatDisplayValue(displayRaw, format),
      ...(typeof storedValue === "string" && storedValue.startsWith("=") ? { formula: storedValue } : {}),
      metadata,
      editable: !column.calculatedFormula && !getCellReadOnly(parent.workbook, table.sheetId, address),
      issues: []
    };
  }

  function dispatchSelection(table: StructuredTable, selection: TableSelection): CommandResult {
    const start = resolveCell(table, selection.anchor);
    const end = resolveCell(table, selection.focus);
    if (!start || !end) return rejected("validation", "TABLE_CELL_NOT_FOUND", "Selected table cell does not exist");
    return workbookSession.dispatch({ type: "selection.set", selection: { start, end } });
  }

  function dispatchMetadata(
    table: StructuredTable,
    updates: readonly TableCellMetadataUpdate[]
  ): CommandResult {
    const commands: WorkbookCommand[] = [];
    const formulaEdits: Array<{ rowId: string; columnId: string; rawText: string }> = [];
    for (const update of updates) {
      const coord = resolveCell(table, update);
      if (!coord) return rejected("validation", "TABLE_CELL_NOT_FOUND", "Metadata target does not exist");
      const range = { start: coord, end: coord };
      if (update.patch.format) {
        commands.push({
          type: "range.format",
          sheetId: table.sheetId,
          range,
          format: fromTableFormat(update.patch.format)
        });
      }
      if ("validation" in update.patch) {
        commands.push(update.patch.validation
          ? { type: "range.validation.set", sheetId: table.sheetId, range, rule: fromTableValidation(update.patch.validation) }
          : { type: "range.validation.clear", sheetId: table.sheetId, range });
      }
      if ("comment" in update.patch) {
        commands.push({
          type: "cell.comment.set",
          sheetId: table.sheetId,
          address: formatCellAddress(coord),
          comment: update.patch.comment ?? null
        });
      }
      if ("readOnly" in update.patch) {
        commands.push({
          type: "range.readOnly.set",
          sheetId: table.sheetId,
          range,
          readOnly: update.patch.readOnly === true
        });
      }
      if ("formula" in update.patch) {
        formulaEdits.push({
          rowId: update.rowId,
          columnId: update.columnId,
          rawText: update.patch.formula ?? ""
        });
      }
    }
    if (formulaEdits.length > 0) commands.push({ type: "table.editCells", tableId, edits: formulaEdits });
    if (commands.length === 0) return { status: "committed", revision: getSnapshot().revision, changed: false };
    return workbookSession.dispatch({ type: "transaction", commands });
  }

  function resolveCell(table: StructuredTable, cell: TableCellRef) {
    const body = getStructuredTableBodyRange(table);
    const rowIndex = table.rowIds.indexOf(cell.rowId);
    const column = table.columns.find((candidate) => candidate.id === cell.columnId);
    return !body || rowIndex < 0 || !column
      ? null
      : { row: body.start.row + rowIndex, column: column.sheetColumn };
  }

  function commitLocal(change: () => void): CommandResult {
    change();
    localRevision += 1;
    cachedSnapshot = null;
    publish();
    return { status: "committed", revision: getSnapshot().revision, changed: true };
  }

  function publish(): void {
    for (const listener of [...listeners]) {
      try { listener(); } catch { /* Host listeners are isolated. */ }
    }
  }
}

function selectionFromWorkbook(table: StructuredTable, selection: { start: { row: number; column: number }; end: { row: number; column: number } }): TableSelection | null {
  const body = getStructuredTableBodyRange(table);
  if (!body) return null;
  const map = (coord: { row: number; column: number }): TableCellRef | null => {
    const rowIndex = coord.row - body.start.row;
    const column = table.columns.find((candidate) => candidate.sheetColumn === coord.column);
    return rowIndex < 0 || rowIndex >= table.rowIds.length || !column
      ? null
      : { rowId: table.rowIds[rowIndex], columnId: column.id };
  };
  const anchor = map(selection.start);
  const focus = map(selection.end);
  return anchor && focus ? { anchor, focus } : null;
}

function cellMetadata(
  workbook: ReturnType<WorkbookSession["getSnapshot"]>["workbook"],
  table: StructuredTable,
  address: string,
  format: CellFormat,
  storedValue: unknown
): TableCellMetadata {
  const validation = getCellValidation(workbook, table.sheetId, address);
  const comment = getCellComment(workbook, table.sheetId, address);
  return {
    ...(Object.keys(format).length === 0 ? {} : { format: toTableFormat(format) }),
    ...(validation ? { validation: toTableValidation(validation) } : {}),
    ...(comment === null ? {} : { comment }),
    ...(typeof storedValue === "string" && storedValue.startsWith("=") ? { formula: storedValue } : {}),
    ...(getCellReadOnly(workbook, table.sheetId, address) ? { readOnly: true } : {})
  };
}

function toTableFormat(format: CellFormat): NonNullable<TableCellMetadata["format"]> {
  const { borders: _borders, numberFormat, ...rest } = format;
  return {
    ...rest,
    ...(numberFormat === undefined
      ? {}
      : { numberFormat: numberFormat === "dateTime" ? "datetime" : numberFormat })
  };
}

function fromTableFormat(format: NonNullable<TableCellMetadata["format"]>): Partial<CellFormat> {
  const { numberFormat, ...rest } = format;
  return {
    ...rest,
    ...(numberFormat === undefined
      ? {}
      : { numberFormat: numberFormat === "datetime" ? "dateTime" : numberFormat })
  };
}

function toTableValidation(validation: DataValidationRule): NonNullable<TableCellMetadata["validation"]> {
  if (validation.type === "list") return { kind: "list", values: [...validation.values], allowBlank: validation.allowBlank };
  return { kind: validation.type, min: validation.min, max: validation.max, allowBlank: validation.allowBlank };
}

function fromTableValidation(validation: NonNullable<TableCellMetadata["validation"]>): DataValidationRule {
  if (validation.kind === "list") return { type: "list", values: [...validation.values], allowBlank: validation.allowBlank };
  return { type: validation.kind, min: validation.min, max: validation.max, allowBlank: validation.allowBlank };
}

function missingTableSnapshot(
  revision: string,
  canUndo: boolean,
  canRedo: boolean
): TableViewSnapshot<WorkbookTableRow> {
  const issue = { code: "TABLE_NOT_FOUND", message: "Structured table does not exist" };
  const state: TableViewState = {
    sorting: [], filter: null, grouping: [], aggregates: [], pagination: { kind: "none" },
    selection: null, selectedRowIds: [], expandedRowIds: [], columnOrder: [],
    columnVisibility: {}, columnWidths: {}, columnPinning: { left: [], right: [] }
  };
  return {
    revision,
    rows: [],
    columns: [],
    rowCount: 0,
    totalRowCount: { kind: "known", value: 0 },
    completeness: "completeDataset",
    state,
    selection: null,
    status: { phase: "error", message: issue.message },
    issues: [issue],
    capabilities: WORKBOOK_TABLE_CAPABILITIES,
    operationStates: OPERATION_STATES,
    pendingOperations: [],
    conflicts: [],
    canUndo,
    canRedo,
    pageInfo: { kind: "none", total: { kind: "known", value: 0 } },
    getCell(rowId, columnId) {
      return {
        rowId, columnId, storedValue: null, evaluatedValue: null, displayValue: "",
        metadata: {}, editable: false, issues: [{ ...issue, rowId, columnId }]
      };
    },
    getRowIndex: () => -1,
    getColumnIndex: () => -1
  };
}

function rejected(
  reason: "validation" | "permission" | "unsupported",
  code: string,
  message: string
): CommandResult {
  return { status: "rejected", reason, issues: [{ code, message }] };
}

function unsupported(message: string): CommandResult {
  return rejected("unsupported", "TABLE_OPERATION_UNSUPPORTED", message);
}
