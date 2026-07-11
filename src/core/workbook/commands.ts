import type {
  BorderPreset,
  CellCoord,
  CellFormat,
  CellRange,
  ConditionalFormatRule,
  DataValidationRule,
  NamedRange,
  SheetChart,
  SheetFilter,
  WorkbookModel
} from "../../types";
import { formatCellAddress, getRangeAddresses, normalizeRange, parseCellAddress } from "../../lib/addressing";
import type { ComputedCellValue } from "../../lib/formulaEngine";
import {
  addSheet,
  autoFillRange,
  clearCellComments,
  clearConditionalFormatRules,
  clearDirectCellFormats,
  clearCellFormats,
  clearCellHyperlinks,
  clearRange,
  clearRangeAll,
  clearSheetFilter,
  clearSheetFilters,
  copyRichRange,
  defineNamedRange,
  deleteSheet,
  deleteSheetChart,
  duplicateSheet,
  fillDown,
  fillRight,
  getCellContent,
  getCellFormat,
  getCellReadOnly,
  getCellValidation,
  getSheetFilters,
  mergeCells,
  moveRichRange,
  moveSheet,
  pasteRichRange,
  pasteMatrix,
  previewRichPaste,
  removeConditionalFormatRule,
  removeDuplicateRows,
  removeNamedRange,
  renameSheet,
  setActiveSheet,
  setCellBorders,
  setCellComment,
  setCellContent,
  setCellFormat,
  setCellHyperlink,
  setCellValidation,
  setColumnWidth,
  setColumnsHidden,
  setRangeReadOnly,
  setRowHeight,
  setRowsHidden,
  setSheetFreezePanes,
  setSheetHidden,
  setSheetProtection,
  setSheetTabColor,
  sortRange,
  unmergeCells,
  type RichClipboardCell,
  type RichClipboardRange,
  type RichPasteMode
} from "../../lib/workbook";
import { parseCellInput } from "../values/parseCellInput";
import { validateCellCandidate } from "../../lib/validation";
import type { TableIssue } from "../commands/types";
import { createRandomId, type IdGenerator, type IdKind } from "../ids";
import {
  getStructuredTableAtCell,
  getStructuredTableBodyRange,
  reconcileStructuredTableContentWrites,
  reduceStructuredTableCommand,
  type StructuredTableCommand
} from "./structuredTables";
import {
  isWorksheetStructureCommand,
  reduceWorksheetStructureCommand,
  type WorksheetStructureCommand
} from "./worksheetStructure";
import {
  isMigratableCellContent,
  migrateConditionalFormatRule,
  migrateSheetFilter,
  migrateValidationRule
} from "./migrateWorkbook";

export type SerializableRichClipboardRange = {
  readonly range: CellRange;
  readonly cells: readonly (readonly RichClipboardCell[])[];
};

/**
 * Declares an exact ID expected from the session allocator during one outer
 * transaction. `occurrence` is zero-based independently for each ID kind and
 * includes allocations made by nested child commands. Reservations belong on
 * the dispatched outer transaction only; the real allocator output must match.
 *
 * Explicit reservations are required for commands that refer to an ID created
 * earlier in the same transaction because an opaque stateful allocator cannot
 * be predicted or rewound safely during semantic preflight.
 */
export type WorkbookIdReservation = Readonly<{
  kind: IdKind;
  occurrence: number;
  id: string;
}>;

export type WorkbookCommand =
  | {
      type: "transaction";
      commands: readonly WorkbookCommand[];
      /** Exact generated IDs needed by later children in this transaction. */
      idReservations?: readonly WorkbookIdReservation[];
    }
  | { type: "selection.set"; selection: CellRange }
  | { type: "cell.set"; sheetId: string; address: string; input: string }
  | { type: "cell.comment.set"; sheetId: string; address: string; comment: string | null }
  | { type: "cell.hyperlink.set"; sheetId: string; address: string; hyperlink: string | null }
  | {
      type: "range.clear";
      sheetId: string;
      range: CellRange;
      mode: "contents" | "formats" | "comments" | "hyperlinks" | "all";
    }
  | { type: "range.format"; sheetId: string; range: CellRange; format: Partial<CellFormat> }
  | { type: "range.directFormat.clear"; sheetId: string; range: CellRange }
  | { type: "range.format.replace"; sheetId: string; range: CellRange; format: Partial<CellFormat> }
  | { type: "range.borders"; sheetId: string; range: CellRange; preset: BorderPreset }
  | { type: "range.validation.set"; sheetId: string; range: CellRange; rule: DataValidationRule }
  | { type: "range.validation.clear"; sheetId: string; range: CellRange }
  | {
      type: "range.conditionalFormat.add";
      sheetId: string;
      range: CellRange;
      rule: ConditionalFormatRule;
    }
  | { type: "range.conditionalFormat.remove"; sheetId: string; ruleId: string }
  | { type: "range.conditionalFormat.clear"; sheetId: string; range: CellRange }
  | { type: "range.readOnly.set"; sheetId: string; range: CellRange; readOnly: boolean }
  | { type: "range.merge" | "range.unmerge"; sheetId: string; range: CellRange }
  | { type: "range.fill"; sheetId: string; range: CellRange; direction: "down" | "right" }
  | { type: "range.autoFill"; sheetId: string; source: CellRange; target: CellRange }
  | {
      type: "range.sort";
      sheetId: string;
      range: CellRange;
      direction: "asc" | "desc";
      sortColumn?: number;
    }
  | { type: "range.removeDuplicates"; sheetId: string; range: CellRange }
  | {
      type: "clipboard.paste";
      sheetId: string;
      target: CellCoord;
      payload: SerializableRichClipboardRange;
      mode: RichPasteMode;
    }
  | {
      type: "clipboard.move";
      sourceSheetId: string;
      source: CellRange;
      targetSheetId: string;
      target: CellCoord;
    }
  | {
      type: "clipboard.pasteMatrix";
      sheetId: string;
      target: CellCoord;
      matrix: readonly (readonly string[])[];
    }
  | WorksheetStructureCommand
  | { type: "rows.resize"; sheetId: string; rows: readonly number[]; height: number }
  | { type: "columns.resize"; sheetId: string; columns: readonly number[]; width: number }
  | { type: "rows.hidden.set"; sheetId: string; rows: readonly number[]; hidden: boolean }
  | { type: "columns.hidden.set"; sheetId: string; columns: readonly number[]; hidden: boolean }
  | { type: "sheet.add"; name?: string }
  | {
      type: "sheet.createFromMatrix";
      sheetId: string;
      name: string;
      rows: readonly (readonly string[])[];
      formats?: readonly Readonly<{ range: CellRange; format: Partial<CellFormat> }>[];
      columnWidths?: readonly number[];
      freeze?: Readonly<{ rows: 0 | 1; columns: 0 | 1 }>;
    }
  | { type: "sheet.replaceWithRows"; sheetId: string; rows: readonly (readonly string[])[] }
  | { type: "sheet.rename"; sheetId: string; name: string }
  | { type: "sheet.duplicate" | "sheet.delete" | "sheet.activate"; sheetId: string }
  | { type: "sheet.move"; sheetId: string; targetIndex: number }
  | { type: "sheet.hidden.set"; sheetId: string; hidden: boolean }
  | { type: "sheet.tabColor.set"; sheetId: string; color: string }
  | { type: "sheet.freeze.set"; sheetId: string; rows: number; columns: number }
  | { type: "sheet.protection.set"; sheetId: string; protected: boolean }
  | { type: "sheet.filter.set"; sheetId: string; filter: SheetFilter }
  | { type: "sheet.filter.clear"; sheetId: string; column?: number }
  | { type: "sheet.chart.add"; sheetId: string; chart: SheetChart }
  | { type: "sheet.chart.delete"; sheetId: string; chartId: string }
  | { type: "namedRange.define"; namedRange: NamedRange }
  | { type: "namedRange.remove"; name: string }
  | { type: "history.undo" | "history.redo" }
  | {
      type: "persistence.status";
      status: "idle" | "saving" | "failed";
      operation?: "load" | "save";
      message?: string;
    }
  | { type: "workbook.replace"; workbook: WorkbookModel; history: "commit" | "reset" }
  | StructuredTableCommand;

export type WorkbookMutationContext = {
  evaluateCell(workbook: WorkbookModel, sheetId: string, address: string): ComputedCellValue;
  createId?: IdGenerator;
};

export type WorkbookMutationResult =
  | { status: "applied"; workbook: WorkbookModel }
  | {
      status: "rejected";
      reason: "validation" | "permission" | "unsupported";
      issues?: readonly TableIssue[];
    };

export function applyWorkbookMutation(
  workbook: WorkbookModel,
  command: WorkbookCommand,
  context: WorkbookMutationContext
): WorkbookMutationResult {
  if (isStructuredTableCommand(command)) {
    const reduction = reduceStructuredTableCommand(workbook, command, {
      createId: context.createId ?? createRandomId,
      getCellEvaluation(sheetId, address) {
        return context.evaluateCell(workbook, sheetId, address);
      }
    });
    if (reduction.status === "rejected") {
      return { status: "rejected", reason: "validation", issues: reduction.issues };
    }
    if (command.type === "table.editCells" && reduction.status === "committed") {
      const table = workbook.tables.find((candidate) => candidate.id === command.tableId);
      const body = table ? getStructuredTableBodyRange(table) : null;
      if (table && body) {
        const addresses = command.edits.flatMap((edit) => {
          const rowIndex = table.rowIds.indexOf(edit.rowId);
          const column = table.columns.find((candidate) => candidate.id === edit.columnId);
          return rowIndex < 0 || !column
            ? []
            : [formatCellAddress({ row: body.start.row + rowIndex, column: column.sheetColumn })];
        });
        return validatedMutation(reduction.workbook, table.sheetId, addresses, context);
      }
    }
    return applied(reduction.workbook);
  }
  if (isWorksheetStructureCommand(command)) {
    const reduction = reduceWorksheetStructureCommand(workbook, command, {
      createId: context.createId ?? createRandomId
    });
    return reduction.status === "committed"
      ? applied(reduction.workbook)
      : {
          status: "rejected",
          reason: reduction.reason,
          issues: reduction.issues
        };
  }
  switch (command.type) {
    case "cell.set": {
      const coordinate = parseCellAddress(command.address);
      const permission = writableAddresses(workbook, command.sheetId, [command.address]);
      if (permission) {
        return permission;
      }
      const table = getStructuredTableAtCell(workbook, command.sheetId, coordinate);
      if (table) {
        const column = table.columns.find((candidate) => candidate.sheetColumn === coordinate.column);
        if (column && table.headerRow && coordinate.row === table.range.start.row) {
          const reduction = reduceStructuredTableCommand(workbook, {
            type: "table.renameColumn",
            tableId: table.id,
            columnId: column.id,
            name: command.input
          }, {
            createId: context.createId ?? createRandomId,
            getCellEvaluation: (sheetId, address) => context.evaluateCell(workbook, sheetId, address)
          });
          return reduction.status === "rejected"
            ? { status: "rejected", reason: "validation", issues: reduction.issues }
            : applied(reduction.workbook);
        }
        const body = getStructuredTableBodyRange(table);
        if (column?.calculatedFormula && body && coordinate.row >= body.start.row && coordinate.row <= body.end.row) {
          return {
            status: "rejected",
            reason: "permission",
            issues: [{
              code: "TABLE_CALCULATED_COLUMN_READ_ONLY",
              message: "Calculated table columns are read-only",
              sheetId: command.sheetId,
              address: command.address
            }]
          };
        }
      }
      const parsed = parseCellInput(command.input);
      let candidate = setCellContent(workbook, command.sheetId, command.address, parsed.stored);
      if (
        parsed.inferredNumberFormat
        && getCellFormat(workbook, command.sheetId, command.address).numberFormat === undefined
      ) {
        candidate = setCellFormat(
          candidate,
          command.sheetId,
          { start: coordinate, end: coordinate },
          { numberFormat: parsed.inferredNumberFormat }
        );
      }
      return validatedContentMutation(candidate, command.sheetId, [command.address], context);
    }
    case "cell.comment.set": {
      parseCellAddress(command.address);
      const permission = writableAddresses(workbook, command.sheetId, [command.address]);
      return permission ?? applied(setCellComment(
        workbook,
        command.sheetId,
        command.address,
        command.comment
      ));
    }
    case "cell.hyperlink.set": {
      parseCellAddress(command.address);
      const permission = writableAddresses(workbook, command.sheetId, [command.address]);
      return permission ?? applied(setCellHyperlink(
        workbook,
        command.sheetId,
        command.address,
        command.hyperlink
      ));
    }
    case "range.clear": {
      const range = checkedRange(command.range);
      const addresses = getRangeAddresses(range);
      const mutatesContent = command.mode === "contents" || command.mode === "all";
      const permission = mutatesContent
        ? contentMutationPermission(workbook, command.sheetId, addresses)
        : writableAddresses(workbook, command.sheetId, addresses);
      if (permission) {
        return permission;
      }
      let candidate = workbook;
      if (command.mode === "contents") {
        candidate = clearRange(candidate, command.sheetId, range);
      } else if (command.mode === "formats") {
        candidate = clearCellFormats(candidate, command.sheetId, range);
      } else if (command.mode === "comments") {
        candidate = clearCellComments(candidate, command.sheetId, range);
      } else if (command.mode === "hyperlinks") {
        candidate = clearCellHyperlinks(candidate, command.sheetId, range);
      } else {
        candidate = clearRangeAll(candidate, command.sheetId, range);
        candidate = clearCellComments(candidate, command.sheetId, range);
        candidate = clearCellHyperlinks(candidate, command.sheetId, range);
        candidate = setCellValidation(candidate, command.sheetId, range, null);
      }
      return command.mode === "contents" || command.mode === "all"
        ? validatedContentMutation(candidate, command.sheetId, addresses, context)
        : applied(candidate);
    }
    case "range.format": {
      const range = checkedRange(command.range);
      const permission = writableAddresses(workbook, command.sheetId, getRangeAddresses(range));
      return permission ?? applied(setCellFormat(workbook, command.sheetId, range, command.format));
    }
    case "range.directFormat.clear": {
      const range = checkedRange(command.range);
      const permission = writableAddresses(workbook, command.sheetId, getRangeAddresses(range));
      return permission ?? applied(clearDirectCellFormats(workbook, command.sheetId, range));
    }
    case "range.format.replace": {
      const range = checkedRange(command.range);
      const permission = writableAddresses(workbook, command.sheetId, getRangeAddresses(range));
      if (permission) {
        return permission;
      }
      const cleared = clearDirectCellFormats(workbook, command.sheetId, range);
      return applied(setCellFormat(cleared, command.sheetId, range, command.format));
    }
    case "range.borders": {
      const range = checkedRange(command.range);
      const permission = writableAddresses(workbook, command.sheetId, getRangeAddresses(range));
      return permission ?? applied(setCellBorders(workbook, command.sheetId, range, command.preset));
    }
    case "range.validation.set": {
      const range = checkedRange(command.range);
      const rangeIssue = rangeBoundsValidation(
        workbook,
        command.sheetId,
        range,
        "Validation range must be inside the sheet"
      );
      if (rangeIssue) return rangeIssue;
      const ruleIssue = validationRuleValidation(command.rule, command.sheetId);
      if (ruleIssue) return ruleIssue;
      const permission = writableAddresses(workbook, command.sheetId, getRangeAddresses(range));
      return permission ?? applied(setCellValidation(workbook, command.sheetId, range, command.rule));
    }
    case "range.validation.clear": {
      const range = checkedRange(command.range);
      const permission = writableAddresses(workbook, command.sheetId, getRangeAddresses(range));
      return permission ?? applied(setCellValidation(workbook, command.sheetId, range, null));
    }
    case "range.conditionalFormat.add": {
      const range = checkedRange(command.range);
      const rangeIssue = rangeBoundsValidation(
        workbook,
        command.sheetId,
        range,
        "Conditional format range must be inside the sheet"
      );
      if (rangeIssue) return rangeIssue;
      const sheet = workbook.sheets.find((candidate) => candidate.id === command.sheetId)!;
      const rule = migrateConditionalFormatRule({ ...command.rule, range }, sheet);
      if (!rule) {
        return {
          status: "rejected",
          reason: "validation",
          issues: [{
            code: "conditionalFormat.rule.invalid",
            message: "Conditional format rule must have a nonblank id and supported condition and format",
            sheetId: command.sheetId
          }]
        };
      }
      const permission = writableAddresses(workbook, command.sheetId, getRangeAddresses(range));
      return permission ?? applied(addConditionalFormat(workbook, command.sheetId, range, rule));
    }
    case "range.conditionalFormat.remove": {
      const sheet = workbook.sheets.find((candidate) => candidate.id === command.sheetId);
      const rule = sheet?.conditionalFormats.find((candidate) => candidate.id === command.ruleId);
      const permission = rule
        ? writableAddresses(workbook, command.sheetId, getRangeAddresses(rule.range))
        : null;
      return permission ?? applied(removeConditionalFormatRule(workbook, command.sheetId, command.ruleId));
    }
    case "range.conditionalFormat.clear": {
      const range = checkedRange(command.range);
      const permission = writableAddresses(workbook, command.sheetId, getRangeAddresses(range));
      return permission ?? applied(clearConditionalFormatRules(workbook, command.sheetId, range));
    }
    case "range.readOnly.set":
      return applied(setRangeReadOnly(
        workbook,
        command.sheetId,
        checkedRange(command.range),
        command.readOnly
      ));
    case "range.merge": {
      const range = checkedRange(command.range);
      if (workbook.tables.some((table) => table.sheetId === command.sheetId && rangesIntersect(table.range, range))) {
        return {
          status: "rejected",
          reason: "validation",
          issues: [{ code: "TABLE_PARTIAL_STRUCTURAL_EDIT", message: "Merged cells cannot intersect a table" }]
        };
      }
      const permission = writableAddresses(workbook, command.sheetId, getRangeAddresses(range));
      return permission ?? applied(mergeCells(workbook, command.sheetId, range));
    }
    case "range.unmerge": {
      const range = checkedRange(command.range);
      if (workbook.tables.some((table) => table.sheetId === command.sheetId && rangesIntersect(table.range, range))) {
        return {
          status: "rejected",
          reason: "validation",
          issues: [{ code: "TABLE_PARTIAL_STRUCTURAL_EDIT", message: "Merged cells cannot intersect a table" }]
        };
      }
      const permission = writableAddresses(workbook, command.sheetId, getRangeAddresses(range));
      return permission ?? applied(unmergeCells(workbook, command.sheetId, range));
    }
    case "range.fill": {
      const range = checkedRange(command.range);
      const addresses = fillDestinationAddresses(range, command.direction);
      const permission = contentMutationPermission(workbook, command.sheetId, addresses);
      if (permission) {
        return permission;
      }
      const candidate = command.direction === "down"
        ? fillDown(workbook, command.sheetId, range)
        : fillRight(workbook, command.sheetId, range);
      return validatedContentMutation(candidate, command.sheetId, addresses, context);
    }
    case "range.autoFill": {
      const source = checkedRange(command.source);
      const target = checkedRange(command.target);
      const sourceAddresses = new Set(getRangeAddresses(source));
      const addresses = getRangeAddresses(target).filter((address) => !sourceAddresses.has(address));
      const permission = contentMutationPermission(workbook, command.sheetId, addresses);
      if (permission) {
        return permission;
      }
      const candidate = autoFillRange(workbook, command.sheetId, command.source, target);
      return validatedContentMutation(candidate, command.sheetId, addresses, context);
    }
    case "range.sort": {
      const range = checkedRange(command.range);
      const headerPermission = tableHeaderMutationPermission(
        workbook,
        command.sheetId,
        range,
        "Sorting cannot include a structured-table header"
      );
      if (headerPermission) {
        return headerPermission;
      }
      const bodyPermission = tableBodyMutationPermission(
        workbook,
        command.sheetId,
        range,
        "Use table.sort to preserve structured-table row IDs"
      );
      if (bodyPermission) return bodyPermission;
      const addresses = getRangeAddresses(range);
      const permission = contentMutationPermission(workbook, command.sheetId, addresses);
      if (permission) {
        return permission;
      }
      const candidate = sortRange(workbook, command.sheetId, range, {
        direction: command.direction,
        sortColumn: command.sortColumn,
        readValue: (address) => context.evaluateCell(workbook, command.sheetId, address)
      });
      return validatedContentMutation(candidate, command.sheetId, addresses, context);
    }
    case "range.removeDuplicates": {
      const range = checkedRange(command.range);
      const headerPermission = tableHeaderMutationPermission(
        workbook,
        command.sheetId,
        range,
        "Duplicate removal cannot include a structured-table header"
      );
      if (headerPermission) {
        return headerPermission;
      }
      const bodyPermission = tableBodyMutationPermission(
        workbook,
        command.sheetId,
        range,
        "Duplicate removal is unsupported inside a structured table"
      );
      if (bodyPermission) return bodyPermission;
      const addresses = getRangeAddresses(range);
      const permission = contentMutationPermission(workbook, command.sheetId, addresses);
      if (permission) {
        return permission;
      }
      return validatedContentMutation(
        removeDuplicateRows(workbook, command.sheetId, range).workbook,
        command.sheetId,
        addresses,
        context
      );
    }
    case "clipboard.paste": {
      const modeIssue = clipboardPasteModeValidation(command.mode, command.sheetId);
      if (modeIssue) return modeIssue;
      const startAddress = formatCellAddress(checkedCoordinate(command.target));
      const clipboard = mutableClipboard(command.payload);
      const contentIssue = clipboardContentValidation(clipboard, command.mode, command.sheetId);
      if (contentIssue) return contentIssue;
      if (command.mode === "all" || command.mode === "transpose") {
        const ruleIssue = clipboardValidationRulesValidation(clipboard, command.sheetId);
        if (ruleIssue) return ruleIssue;
      }
      const addresses = previewRichPaste(clipboard, startAddress, { mode: command.mode })
        .map((cell) => cell.address);
      const mutatesContent = command.mode !== "formats";
      const permission = mutatesContent
        ? contentMutationPermission(workbook, command.sheetId, addresses)
        : writableAddresses(workbook, command.sheetId, addresses);
      if (permission) {
        return permission;
      }
      const candidate = pasteRichRange(workbook, command.sheetId, startAddress, clipboard, {
        mode: command.mode
      });
      return mutatesContent
        ? validatedContentMutation(candidate, command.sheetId, addresses, context)
        : validatedMutation(candidate, command.sheetId, addresses, context);
    }
    case "clipboard.move": {
      const source = checkedRange(command.source);
      const targetAddress = formatCellAddress(checkedCoordinate(command.target));
      const clipboard = copyRichRange(workbook, command.sourceSheetId, source);
      const targetAddresses = previewRichPaste(clipboard, targetAddress).map((cell) => cell.address);
      const targetAddressSet = command.sourceSheetId === command.targetSheetId
        ? new Set(targetAddresses)
        : null;
      const sourceAddresses = getRangeAddresses(source)
        .filter((address) => !targetAddressSet?.has(address));
      const writes = combineContentWrites([
        { sheetId: command.sourceSheetId, addresses: sourceAddresses },
        { sheetId: command.targetSheetId, addresses: targetAddresses }
      ]);
      for (const write of writes) {
        const permission = contentMutationPermission(workbook, write.sheetId, write.addresses);
        if (permission) return permission;
      }
      const candidate = moveRichRange(
        workbook,
        command.sourceSheetId,
        command.targetSheetId,
        targetAddress,
        clipboard
      );
      return validatedContentMutations(candidate, writes, context);
    }
    case "clipboard.pasteMatrix": {
      const target = checkedCoordinate(command.target);
      const matrix = cloneStringMatrix(command.matrix);
      const addresses = matrixAddresses(target, matrix);
      const permission = contentMutationPermission(workbook, command.sheetId, addresses);
      if (permission) {
        return permission;
      }
      const candidate = pasteMatrix(
        workbook,
        command.sheetId,
        formatCellAddress(target),
        matrix
      );
      return validatedContentMutation(candidate, command.sheetId, addresses, context);
    }
    case "rows.resize": {
      checkedDimension(command.height);
      let candidate = workbook;
      for (const row of checkedIndexes(command.rows)) {
        candidate = setRowHeight(candidate, command.sheetId, row, command.height);
      }
      return applied(candidate);
    }
    case "columns.resize": {
      checkedDimension(command.width);
      let candidate = workbook;
      for (const column of checkedIndexes(command.columns)) {
        candidate = setColumnWidth(candidate, command.sheetId, column, command.width);
      }
      return applied(candidate);
    }
    case "rows.hidden.set": {
      let candidate = workbook;
      for (const row of checkedIndexes(command.rows)) {
        candidate = setRowsHidden(candidate, command.sheetId, row, row, command.hidden);
      }
      return applied(candidate);
    }
    case "columns.hidden.set": {
      let candidate = workbook;
      for (const column of checkedIndexes(command.columns)) {
        candidate = setColumnsHidden(candidate, command.sheetId, column, column, command.hidden);
      }
      return applied(candidate);
    }
    case "sheet.add":
      return applied(addSheet(workbook, command.name));
    case "sheet.createFromMatrix": {
      if (!command.sheetId.trim()) {
        return rejectedUnsupported("sheet.id.invalid", "Generated sheet id is required");
      }
      const rows = cloneStringMatrix(command.rows);
      let candidate = addSheet(workbook, command.name);
      if (candidate.activeSheetId !== command.sheetId) {
        return rejectedUnsupported("sheet.id.conflict", "Generated sheet id is stale");
      }
      candidate = pasteMatrix(candidate, command.sheetId, "A1", rows);
      for (const entry of command.formats ?? []) {
        candidate = setCellFormat(candidate, command.sheetId, checkedRange(entry.range), entry.format);
      }
      for (const [column, width] of (command.columnWidths ?? []).entries()) {
        checkedDimension(width);
        candidate = setColumnWidth(candidate, command.sheetId, column, width);
      }
      if (command.freeze) {
        checkedFreezeCount(command.freeze.rows);
        checkedFreezeCount(command.freeze.columns);
        candidate = setSheetFreezePanes(candidate, command.sheetId, {
          freezeTopRow: command.freeze.rows === 1,
          freezeFirstColumn: command.freeze.columns === 1
        });
      }
      return applied(candidate);
    }
    case "sheet.replaceWithRows":
      return applied(replaceSheetWithRows(workbook, command.sheetId, cloneStringMatrix(command.rows)));
    case "sheet.rename":
      return applied(renameSheet(workbook, command.sheetId, command.name));
    case "sheet.duplicate":
      return applied(duplicateSheet(workbook, command.sheetId, context.createId ?? createRandomId));
    case "sheet.delete":
      return applied(deleteSheet(workbook, command.sheetId));
    case "sheet.activate":
      return applied(setActiveSheet(workbook, command.sheetId));
    case "sheet.move":
      checkedIndex(command.targetIndex);
      return applied(moveSheet(workbook, command.sheetId, command.targetIndex));
    case "sheet.hidden.set":
      return applied(setSheetHidden(workbook, command.sheetId, command.hidden));
    case "sheet.tabColor.set":
      return applied(setSheetTabColor(workbook, command.sheetId, command.color));
    case "sheet.freeze.set":
      checkedFreezeCount(command.rows);
      checkedFreezeCount(command.columns);
      return applied(setSheetFreezePanes(workbook, command.sheetId, {
        freezeTopRow: command.rows === 1,
        freezeFirstColumn: command.columns === 1
      }));
    case "sheet.protection.set":
      return applied(setSheetProtection(workbook, command.sheetId, command.protected));
    case "sheet.filter.set": {
      const range = checkedRange(command.filter.range);
      const rangeIssue = rangeBoundsValidation(
        workbook,
        command.sheetId,
        range,
        "Filter range must be inside the sheet"
      );
      if (rangeIssue) return rangeIssue;
      const sheet = workbook.sheets.find((candidate) => candidate.id === command.sheetId)!;
      const filter = migrateSheetFilter({ ...command.filter, range }, sheet);
      if (!filter) {
        return {
          status: "rejected",
          reason: "validation",
          issues: [{
            code: "filter.invalid",
            message: "Filter id, column, operator, values, and header flag must be valid",
            sheetId: command.sheetId
          }]
        };
      }
      return applied(setSheetFilter(workbook, command.sheetId, filter));
    }
    case "sheet.filter.clear": {
      if (command.column === undefined) {
        return applied(clearSheetFilters(workbook, command.sheetId));
      }
      checkedIndex(command.column);
      let candidate = workbook;
      for (const filter of getSheetFilters(workbook, command.sheetId)) {
        if (filter.column === command.column) {
          candidate = clearSheetFilter(candidate, command.sheetId, filter.range, filter.column);
        }
      }
      return applied(candidate);
    }
    case "sheet.chart.add": {
      const shapeIssue = chartShapeValidation(command.chart, command.sheetId);
      if (shapeIssue) return shapeIssue;
      const range = checkedRange(command.chart.range);
      const rangeIssue = rangeBoundsValidation(
        workbook,
        command.sheetId,
        range,
        "Chart data range must be inside the sheet"
      );
      if (rangeIssue) return rangeIssue;
      const anchor = checkedCoordinate(command.chart.anchor);
      const anchorIssue = coordinateBoundsValidation(
        workbook,
        command.sheetId,
        anchor,
        "Chart anchor must be inside the sheet"
      );
      if (anchorIssue) return anchorIssue;
      return applied(addChart(workbook, command.sheetId, { ...command.chart, range, anchor }));
    }
    case "sheet.chart.delete":
      return applied(deleteSheetChart(workbook, command.sheetId, command.chartId));
    case "namedRange.define": {
      const range = checkedRange(command.namedRange.range);
      const rangeIssue = rangeBoundsValidation(
        workbook,
        command.namedRange.sheetId,
        range,
        "Named range must be inside the sheet"
      );
      if (rangeIssue) return rangeIssue;
      return applied(defineNamedRange(
        workbook,
        command.namedRange.sheetId,
        command.namedRange.name,
        range
      ));
    }
    case "namedRange.remove":
      return applied(removeNamedRange(workbook, command.name));
    case "transaction":
    case "selection.set":
    case "history.undo":
    case "history.redo":
    case "persistence.status":
    case "workbook.replace":
      return {
        status: "rejected",
        reason: "unsupported",
        issues: [{ code: "command.invalid", message: "Command could not be applied" }]
      };
  }
}

export function collectCommandDiagnostics(command: WorkbookCommand): {
  sheetCount: number;
  rangeCount: number;
  cellCount: number;
} {
  const sheets = new Set<string>();
  let rangeCount = 0;
  let cellCount = 0;

  visit(command);
  return { sheetCount: sheets.size, rangeCount, cellCount };

  function visit(current: WorkbookCommand): void {
    if (current.type === "transaction") {
      for (const child of current.commands) {
        visit(child);
      }
      return;
    }

    addSheetId("sheetId" in current ? current.sheetId : undefined);
    addSheetId("sourceSheetId" in current ? current.sourceSheetId : undefined);
    addSheetId("targetSheetId" in current ? current.targetSheetId : undefined);
    if (current.type === "namedRange.define") {
      addSheetId(current.namedRange.sheetId);
    }

    if (current.type === "selection.set") {
      addRange(current.selection);
      return;
    }
    if ("range" in current && current.range) {
      addRange(current.range);
    }
    if (current.type === "range.autoFill") {
      addRange(current.source);
      addRange(current.target);
    } else if (current.type === "clipboard.move") {
      addRange(current.source);
    } else if (current.type === "cell.set" || current.type === "cell.comment.set" || current.type === "cell.hyperlink.set") {
      cellCount += 1;
    } else if (current.type === "clipboard.paste") {
      cellCount += current.payload.cells.reduce((count, row) => count + row.length, 0);
    } else if (current.type === "clipboard.pasteMatrix") {
      cellCount += current.matrix.reduce((count, row) => count + row.length, 0);
    } else if (current.type === "sheet.createFromMatrix" || current.type === "sheet.replaceWithRows") {
      cellCount += current.rows.reduce((count, row) => count + row.length, 0);
    } else if (current.type === "namedRange.define") {
      addRange(current.namedRange.range);
    }
  }

  function addSheetId(sheetId: string | undefined): void {
    if (sheetId) {
      sheets.add(sheetId);
    }
  }

  function addRange(range: CellRange): void {
    rangeCount += 1;
    cellCount += rangeArea(range);
  }
}

function applied(workbook: WorkbookModel): WorkbookMutationResult {
  return { status: "applied", workbook };
}

type ContentWrite = {
  readonly sheetId: string;
  readonly addresses: readonly string[];
};

function fillDestinationAddresses(
  range: CellRange,
  direction: "down" | "right"
): string[] {
  if (direction === "down") {
    return range.start.row === range.end.row
      ? []
      : getRangeAddresses({
          start: { row: range.start.row + 1, column: range.start.column },
          end: range.end
        });
  }
  return range.start.column === range.end.column
    ? []
    : getRangeAddresses({
        start: { row: range.start.row, column: range.start.column + 1 },
        end: range.end
      });
}

function combineContentWrites(writes: readonly ContentWrite[]): ContentWrite[] {
  const addressesBySheet = new Map<string, Set<string>>();
  for (const write of writes) {
    const addresses = addressesBySheet.get(write.sheetId) ?? new Set<string>();
    for (const address of write.addresses) addresses.add(address);
    addressesBySheet.set(write.sheetId, addresses);
  }
  return [...addressesBySheet].map(([sheetId, addresses]) => ({
    sheetId,
    addresses: [...addresses]
  }));
}

function contentMutationPermission(
  workbook: WorkbookModel,
  sheetId: string,
  addresses: readonly string[]
): WorkbookMutationResult | null {
  const permission = writableAddresses(workbook, sheetId, addresses);
  if (permission) return permission;

  for (const address of new Set(addresses)) {
    const coordinate = parseCellAddress(address);
    const table = getStructuredTableAtCell(workbook, sheetId, coordinate);
    if (!table) continue;
    const column = table.columns.find((candidate) => candidate.sheetColumn === coordinate.column);
    const body = getStructuredTableBodyRange(table);
    if (column?.calculatedFormula && body
      && coordinate.row >= body.start.row
      && coordinate.row <= body.end.row) {
      return {
        status: "rejected",
        reason: "permission",
        issues: [{
          code: "TABLE_CALCULATED_COLUMN_READ_ONLY",
          message: "Calculated table columns are read-only",
          sheetId,
          address
        }]
      };
    }
  }
  return null;
}

function tableHeaderMutationPermission(
  workbook: WorkbookModel,
  sheetId: string,
  range: CellRange,
  message: string
): WorkbookMutationResult | null {
  const intersectsHeader = workbook.tables.some((table) => table.sheetId === sheetId
    && table.headerRow
    && rangesIntersect(range, {
      start: { row: table.range.start.row, column: table.range.start.column },
      end: { row: table.range.start.row, column: table.range.end.column }
    }));
  return intersectsHeader
    ? {
        status: "rejected",
        reason: "validation",
        issues: [{ code: "TABLE_HEADER_MUTATION_UNSUPPORTED", message, sheetId }]
      }
    : null;
}

function tableBodyMutationPermission(
  workbook: WorkbookModel,
  sheetId: string,
  range: CellRange,
  message: string
): WorkbookMutationResult | null {
  const intersectsBody = workbook.tables.some((table) => {
    if (table.sheetId !== sheetId) return false;
    const body = getStructuredTableBodyRange(table);
    return body !== null && rangesIntersect(range, body);
  });
  return intersectsBody
    ? {
        status: "rejected",
        reason: "validation",
        issues: [{ code: "TABLE_ROW_ID_MUTATION_UNSUPPORTED", message, sheetId }]
      }
    : null;
}

function validatedContentMutation(
  workbook: WorkbookModel,
  sheetId: string,
  addresses: readonly string[],
  context: WorkbookMutationContext
): WorkbookMutationResult {
  return validatedContentMutations(workbook, [{ sheetId, addresses }], context);
}

function validatedContentMutations(
  workbook: WorkbookModel,
  writes: readonly ContentWrite[],
  context: WorkbookMutationContext
): WorkbookMutationResult {
  let reconciled = workbook;
  for (const write of writes) {
    const reduction = reconcileStructuredTableContentWrites(
      reconciled,
      write.sheetId,
      [...new Set(write.addresses)].map(parseCellAddress)
    );
    if (reduction.status === "rejected") {
      return { status: "rejected", reason: "validation", issues: reduction.issues };
    }
    reconciled = reduction.workbook;
  }
  for (const write of writes) {
    const validation = validatedMutation(reconciled, write.sheetId, write.addresses, context);
    if (validation.status === "rejected") return validation;
  }
  return applied(reconciled);
}

function validatedMutation(
  workbook: WorkbookModel,
  sheetId: string,
  addresses: readonly string[],
  context: WorkbookMutationContext
): WorkbookMutationResult {
  for (const address of new Set(addresses)) {
    const rule = getCellValidation(workbook, sheetId, address);
    if (!rule) {
      continue;
    }
    const parsed = getCellContent(workbook, sheetId, address);
    const formula = typeof parsed === "string" && parsed.startsWith("=") ? parsed : undefined;
    const result = validateCellCandidate({
      raw: parsed === null ? "" : String(parsed),
      parsed,
      evaluated: formula ? context.evaluateCell(workbook, sheetId, address) : parsed,
      ...(formula ? { formula } : {})
    }, rule);
    if (!result.valid) {
      return {
        status: "rejected",
        reason: "validation",
        issues: [{
          code: "validation.failed",
          message: result.message,
          sheetId,
          address
        }]
      };
    }
  }
  return applied(workbook);
}

function writableAddresses(
  workbook: WorkbookModel,
  sheetId: string,
  addresses: readonly string[]
): WorkbookMutationResult | null {
  for (const address of addresses) {
    if (getCellReadOnly(workbook, sheetId, address)) {
      return {
        status: "rejected",
        reason: "permission",
        issues: [{
          code: "permission.readOnly",
          message: "Cell is read-only",
          sheetId,
          address
        }]
      };
    }
  }
  return null;
}

function rangeBoundsValidation(
  workbook: WorkbookModel,
  sheetId: string,
  range: CellRange,
  message: string
): WorkbookMutationResult | null {
  const sheet = workbook.sheets.find((candidate) => candidate.id === sheetId);
  if (sheet
    && range.start.row >= 0
    && range.start.column >= 0
    && range.end.row >= range.start.row
    && range.end.column >= range.start.column
    && range.end.row < sheet.rowCount
    && range.end.column < sheet.columnCount) {
    return null;
  }
  return {
    status: "rejected",
    reason: "validation",
    issues: [{ code: "range.outOfBounds", message, sheetId }]
  };
}

function coordinateBoundsValidation(
  workbook: WorkbookModel,
  sheetId: string,
  coordinate: CellCoord,
  message: string
): WorkbookMutationResult | null {
  const sheet = workbook.sheets.find((candidate) => candidate.id === sheetId);
  if (sheet
    && coordinate.row >= 0
    && coordinate.column >= 0
    && coordinate.row < sheet.rowCount
    && coordinate.column < sheet.columnCount) {
    return null;
  }
  return {
    status: "rejected",
    reason: "validation",
    issues: [{ code: "coordinate.outOfBounds", message, sheetId }]
  };
}

function validationRuleValidation(
  rule: unknown,
  sheetId: string
): WorkbookMutationResult | null {
  return migrateValidationRule(rule)
    ? null
    : {
        status: "rejected",
        reason: "validation",
        issues: [{
          code: "validation.rule.invalid",
          message: "Validation rule must have a supported shape and valid bounds",
          sheetId
        }]
      };
}

function clipboardPasteModeValidation(
  mode: unknown,
  sheetId: string
): WorkbookMutationResult | null {
  return mode === "all" || mode === "values" || mode === "formats" || mode === "transpose"
    ? null
    : {
        status: "rejected",
        reason: "validation",
        issues: [{
          code: "clipboard.mode.invalid",
          message: "Clipboard paste mode is invalid",
          sheetId
        }]
      };
}

function clipboardContentValidation(
  clipboard: RichClipboardRange,
  mode: RichPasteMode,
  sheetId: string
): WorkbookMutationResult | null {
  if (mode === "formats") return null;
  for (const row of clipboard.cells) {
    for (const cell of row) {
      const content = mode === "values" ? cell.displayContent : cell.content;
      if (!isMigratableCellContent(content)) {
        return {
          status: "rejected",
          reason: "validation",
          issues: [{
            code: "clipboard.content.invalid",
            message: "Clipboard content must be a finite number, text, boolean, or blank",
            sheetId
          }]
        };
      }
    }
  }
  return null;
}

function clipboardValidationRulesValidation(
  clipboard: RichClipboardRange,
  sheetId: string
): WorkbookMutationResult | null {
  for (const row of clipboard.cells) {
    for (const cell of row) {
      if (cell.validation !== null && !migrateValidationRule(cell.validation)) {
        return validationRuleValidation(cell.validation, sheetId);
      }
    }
  }
  return null;
}

function chartShapeValidation(
  chart: unknown,
  sheetId: string
): WorkbookMutationResult | null {
  const candidate = typeof chart === "object" && chart !== null && !Array.isArray(chart)
    ? chart as Record<string, unknown>
    : null;
  const valid = candidate !== null
    && typeof candidate.id === "string"
    && candidate.id.trim().length > 0
    && typeof candidate.title === "string"
    && (candidate.type === "bar" || candidate.type === "line" || candidate.type === "pie");
  return valid
    ? null
    : {
        status: "rejected",
        reason: "validation",
        issues: [{
          code: "chart.invalid",
          message: "Chart id, title, and type must be valid",
          sheetId
        }]
      };
}

function checkedRange(range: CellRange): CellRange {
  checkedCoordinate(range.start);
  checkedCoordinate(range.end);
  return normalizeRange(range);
}

function checkedCoordinate(coordinate: CellCoord): CellCoord {
  checkedIndex(coordinate.row);
  checkedIndex(coordinate.column);
  return coordinate;
}

function checkedIndexes(indexes: readonly number[]): readonly number[] {
  for (const index of indexes) {
    checkedIndex(index);
  }
  return indexes;
}

function checkedIndex(index: number): void {
  if (!Number.isInteger(index) || index < 0) {
    throw new Error("Index must be a non-negative integer");
  }
}

function checkedDimension(value: number): void {
  if (!Number.isFinite(value)) {
    throw new Error("Dimension must be finite");
  }
}

function checkedFreezeCount(value: number): void {
  if (value !== 0 && value !== 1) {
    throw new Error("Freeze count must be zero or one");
  }
}

function cloneStringMatrix(matrix: readonly (readonly string[])[]): string[][] {
  return matrix.map((row) => row.map((value) => String(value)));
}

function matrixAddresses(target: CellCoord, matrix: readonly (readonly string[])[]): string[] {
  const addresses: string[] = [];
  for (let rowOffset = 0; rowOffset < matrix.length; rowOffset += 1) {
    for (let columnOffset = 0; columnOffset < matrix[rowOffset].length; columnOffset += 1) {
      addresses.push(formatCellAddress({
        row: target.row + rowOffset,
        column: target.column + columnOffset
      }));
    }
  }
  return addresses;
}

function replaceSheetWithRows(
  workbook: WorkbookModel,
  sheetId: string,
  rows: readonly (readonly string[])[]
): WorkbookModel {
  const rowCount = Math.max(100, rows.length);
  const columnCount = Math.max(26, ...rows.map((row) => row.length));
  const cells: Record<string, string> = {};
  rows.forEach((row, rowIndex) => {
    row.forEach((value, columnIndex) => {
      if (value !== "") {
        cells[formatCellAddress({ row: rowIndex, column: columnIndex })] = value;
      }
    });
  });

  const next = updateSheetModel(workbook, sheetId, (sheet) => ({
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
  }));
  const namedRanges = (next.namedRanges ?? []).filter((namedRange) => namedRange.sheetId !== sheetId);
  const tables = next.tables.filter((table) => table.sheetId !== sheetId);
  return namedRanges.length === (next.namedRanges ?? []).length && tables.length === next.tables.length
    ? next
    : { ...next, namedRanges, tables };
}

function rejectedUnsupported(code: string, message: string): WorkbookMutationResult {
  return { status: "rejected", reason: "unsupported", issues: [{ code, message }] };
}

function isStructuredTableCommand(command: WorkbookCommand): command is StructuredTableCommand {
  return command.type.startsWith("table.");
}

function rangesIntersect(left: CellRange, right: CellRange): boolean {
  const normalizedLeft = normalizeRange(left);
  const normalizedRight = normalizeRange(right);
  return normalizedLeft.start.row <= normalizedRight.end.row
    && normalizedLeft.end.row >= normalizedRight.start.row
    && normalizedLeft.start.column <= normalizedRight.end.column
    && normalizedLeft.end.column >= normalizedRight.start.column;
}

function mutableClipboard(clipboard: SerializableRichClipboardRange): RichClipboardRange {
  return {
    range: {
      start: { ...clipboard.range.start },
      end: { ...clipboard.range.end }
    },
    cells: clipboard.cells.map((row) => row.map((cell) => cloneClipboardCell(cell)))
  };
}

function cloneClipboardCell(cell: RichClipboardCell): RichClipboardCell {
  return {
    ...cell,
    format: { ...cell.format },
    validation: cell.validation
      ? {
          ...cell.validation,
          ...(cell.validation.type === "list" ? { values: [...cell.validation.values] } : {})
        }
      : null
  };
}

function addConditionalFormat(
  workbook: WorkbookModel,
  sheetId: string,
  range: CellRange,
  rule: ConditionalFormatRule
): WorkbookModel {
  const normalized = checkedRange(range);
  return updateSheetModel(workbook, sheetId, (sheet) => {
    const nextRule: ConditionalFormatRule = {
      id: rule.id,
      range: normalized,
      condition: { ...rule.condition },
      format: { ...rule.format }
    };
    const existingIndex = sheet.conditionalFormats.findIndex((item) => item.id === rule.id);
    if (existingIndex >= 0 && persistedEqual(sheet.conditionalFormats[existingIndex], nextRule)) {
      return sheet;
    }
    const conditionalFormats = [...sheet.conditionalFormats];
    if (existingIndex >= 0) {
      conditionalFormats[existingIndex] = nextRule;
    } else {
      conditionalFormats.push(nextRule);
    }
    return { ...sheet, conditionalFormats };
  });
}

function setSheetFilter(workbook: WorkbookModel, sheetId: string, filter: SheetFilter): WorkbookModel {
  const range = checkedRange(filter.range);
  checkedIndex(filter.column);
  return updateSheetModel(workbook, sheetId, (sheet) => {
    const nextFilter: SheetFilter = {
      ...filter,
      range,
      values: filter.values ? [...filter.values] : undefined
    };
    const isSameSlot = (item: SheetFilter) =>
      item.id === filter.id || (rangesEqual(item.range, range) && item.column === filter.column);
    const matching = sheet.filters.filter(isSameSlot);
    if (matching.length === 1 && persistedEqual(matching[0], nextFilter)) {
      return sheet;
    }
    const insertionIndex = sheet.filters.findIndex(isSameSlot);
    const filters = sheet.filters.filter((item) => !isSameSlot(item));
    filters.splice(insertionIndex < 0 ? filters.length : insertionIndex, 0, nextFilter);
    return { ...sheet, autoFilterRange: range, filters };
  });
}

function addChart(workbook: WorkbookModel, sheetId: string, chart: SheetChart): WorkbookModel {
  const range = checkedRange(chart.range);
  checkedCoordinate(chart.anchor);
  return updateSheetModel(workbook, sheetId, (sheet) => {
    const nextChart: SheetChart = {
      ...chart,
      title: chart.title.trim() || "Chart",
      range,
      anchor: { ...chart.anchor }
    };
    const existingIndex = sheet.charts.findIndex((item) => item.id === chart.id);
    if (existingIndex >= 0 && persistedEqual(sheet.charts[existingIndex], nextChart)) {
      return sheet;
    }
    const charts = [...sheet.charts];
    if (existingIndex >= 0) {
      charts[existingIndex] = nextChart;
    } else {
      charts.push(nextChart);
    }
    return { ...sheet, charts };
  });
}

function updateSheetModel(
  workbook: WorkbookModel,
  sheetId: string,
  update: (sheet: WorkbookModel["sheets"][number]) => WorkbookModel["sheets"][number]
): WorkbookModel {
  const index = workbook.sheets.findIndex((sheet) => sheet.id === sheetId);
  if (index < 0) {
    throw new Error("Unknown sheet");
  }
  const current = workbook.sheets[index];
  const next = update(current);
  if (next === current) {
    return workbook;
  }
  const sheets = [...workbook.sheets];
  sheets[index] = next;
  return { ...workbook, sheets };
}

function rangesEqual(left: CellRange, right: CellRange): boolean {
  const normalizedLeft = normalizeRange(left);
  const normalizedRight = normalizeRange(right);
  return normalizedLeft.start.row === normalizedRight.start.row
    && normalizedLeft.start.column === normalizedRight.start.column
    && normalizedLeft.end.row === normalizedRight.end.row
    && normalizedLeft.end.column === normalizedRight.end.column;
}

function rangeArea(range: CellRange): number {
  const normalized = normalizeRange(range);
  const rows = Math.max(0, normalized.end.row - normalized.start.row + 1);
  const columns = Math.max(0, normalized.end.column - normalized.start.column + 1);
  const area = rows * columns;
  return Number.isSafeInteger(area) ? area : Number.MAX_SAFE_INTEGER;
}

function persistedEqual(left: unknown, right: unknown): boolean {
  return JSON.stringify(left) === JSON.stringify(right);
}
