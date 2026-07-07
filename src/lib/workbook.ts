import type {
  BorderPreset,
  CellBorders,
  CellBorderSide,
  CellContent,
  CellFormat,
  CellRange,
  ConditionalFormatRule,
  DataValidationRule,
  HistoryState,
  NamedRange,
  SheetChart,
  SheetMerge,
  SheetModel,
  SheetFilter,
  SheetProtection,
  WorkbookModel
} from "../types";
import {
  columnIndexToName,
  columnNameToIndex,
  formatCellAddress,
  getRangeAddresses,
  normalizeRange,
  parseCellAddress
} from "./addressing";
import {
  DEFAULT_COLUMN_WIDTH,
  DEFAULT_ROW_HEIGHT,
  clampColumnWidth,
  clampRowHeight
} from "./sheetDimensions";
import { translateFormulaReferences } from "./formulaReferences";

const DEFAULT_ROWS = 100;
const DEFAULT_COLUMNS = 26;
const DEFAULT_BORDER: CellBorderSide = { style: "thin", color: "#64748b" };
const HEX_COLOR_PATTERN = /^#[0-9a-f]{6}$/i;
const MS_PER_DAY = 24 * 60 * 60 * 1000;
const AUTO_FILL_NAMED_LISTS = [
  ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"],
  ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"],
  ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"],
  ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"]
] as const;

type StructureOperation = {
  axis: "row" | "column";
  mode: "insert" | "delete";
  index: number;
  count: number;
};

type AutoFillSourceCell = {
  coord: { row: number; column: number };
  content: CellContent;
  format: CellFormat;
  validation: DataValidationRule | null;
};

type AutoFillNamedListSeries = {
  values: readonly string[];
  start: number;
  step: number;
};

type RangeRowSnapshot = {
  originalIndex: number;
  key: string;
  cells: CellContent[];
  formats: CellFormat[];
  validations: (DataValidationRule | null)[];
  comments: (string | null)[];
  hyperlinks: (string | null)[];
};

export type RichClipboardCell = {
  sourceAddress: string;
  content: CellContent;
  displayContent: CellContent;
  format: CellFormat;
  validation: DataValidationRule | null;
  comment: string | null;
  hyperlink: string | null;
};

export type RichClipboardRange = {
  range: CellRange;
  cells: RichClipboardCell[][];
};

export type RichPasteCellPreview = {
  address: string;
  content: CellContent;
  format: CellFormat;
  validation: DataValidationRule | null;
  comment: string | null;
  hyperlink: string | null;
};

export type RichPasteMode = "all" | "values" | "formats" | "transpose";

export type CopyRichRangeOptions = {
  getDisplayValue?: (address: string) => CellContent;
};

export type PasteRichRangeOptions = {
  mode?: RichPasteMode;
  translateFormulas?: boolean;
};

export type RemoveDuplicateRowsResult = {
  workbook: WorkbookModel;
  removedCount: number;
};

export type CellMergeInfo =
  | {
      role: "anchor";
      range: CellRange;
      rowSpan: number;
      columnSpan: number;
    }
  | {
      role: "covered";
      range: CellRange;
      anchor: { row: number; column: number };
      rowSpan: number;
      columnSpan: number;
    };

export function createBlankWorkbook(): WorkbookModel {
  const sheet = createSheet("sheet-1", "Sheet1");
  return {
    version: 1,
    activeSheetId: sheet.id,
    sheets: [sheet],
    namedRanges: []
  };
}

export function getActiveSheet(workbook: WorkbookModel): SheetModel {
  return getSheet(workbook, workbook.activeSheetId);
}

export function getCellContent(workbook: WorkbookModel, sheetId: string, address: string): CellContent {
  return getSheet(workbook, sheetId).cells[normalizeAddress(address)] ?? null;
}

export function getCellFormat(workbook: WorkbookModel, sheetId: string, address: string): CellFormat {
  const sheet = getSheet(workbook, sheetId);
  return cloneCellFormat((sheet.formats ?? {})[normalizeAddress(address)] ?? {});
}

export function getColumnWidth(workbook: WorkbookModel, sheetId: string, column: number): number {
  const width = (getSheet(workbook, sheetId).columnWidths ?? {})[String(Math.max(0, Math.floor(column)))];
  return width === undefined ? DEFAULT_COLUMN_WIDTH : clampColumnWidth(width);
}

export function getRowHeight(workbook: WorkbookModel, sheetId: string, row: number): number {
  const height = (getSheet(workbook, sheetId).rowHeights ?? {})[String(Math.max(0, Math.floor(row)))];
  return height === undefined ? DEFAULT_ROW_HEIGHT : clampRowHeight(height);
}

export function isColumnHidden(workbook: WorkbookModel, sheetId: string, column: number): boolean {
  return Boolean((getSheet(workbook, sheetId).hiddenColumns ?? {})[String(Math.max(0, Math.floor(column)))]);
}

export function isRowHidden(workbook: WorkbookModel, sheetId: string, row: number): boolean {
  return Boolean((getSheet(workbook, sheetId).hiddenRows ?? {})[String(Math.max(0, Math.floor(row)))]);
}

export function isSheetHidden(workbook: WorkbookModel, sheetId: string): boolean {
  return Boolean(getSheet(workbook, sheetId).isHidden);
}

export function getSheetTabColor(workbook: WorkbookModel, sheetId: string): string | undefined {
  return normalizeHexColor(getSheet(workbook, sheetId).tabColor);
}

export function getCellComment(workbook: WorkbookModel, sheetId: string, address: string): string | null {
  return (getSheet(workbook, sheetId).comments ?? {})[normalizeAddress(address)] ?? null;
}

export function getCellHyperlink(workbook: WorkbookModel, sheetId: string, address: string): string | null {
  return (getSheet(workbook, sheetId).hyperlinks ?? {})[normalizeAddress(address)] ?? null;
}

export function getCellReadOnly(workbook: WorkbookModel, sheetId: string, address: string): boolean {
  const protection = getSheetProtection(getSheet(workbook, sheetId));
  const normalizedAddress = normalizeAddress(address);
  return Boolean(protection.lockedCells[normalizedAddress]) || (protection.isProtected && !protection.unlockedCells[normalizedAddress]);
}

export function getCellValidation(workbook: WorkbookModel, sheetId: string, address: string): DataValidationRule | null {
  const rule = (getSheet(workbook, sheetId).validations ?? {})[normalizeAddress(address)];
  return rule ? cloneValidationRule(rule) : null;
}

export function getCellConditionalFormatRules(
  workbook: WorkbookModel,
  sheetId: string,
  address: string
): ConditionalFormatRule[] {
  const coord = parseCellAddress(normalizeAddress(address));
  return (getSheet(workbook, sheetId).conditionalFormats ?? [])
    .filter((rule) => isCoordInRange(coord, normalizeRange(rule.range)))
    .map(cloneConditionalFormatRule);
}

export function getSheetFilters(workbook: WorkbookModel, sheetId: string): SheetFilter[] {
  return (getSheet(workbook, sheetId).filters ?? []).map(cloneSheetFilter);
}

export function getSheetAutoFilterRange(workbook: WorkbookModel, sheetId: string): CellRange | null {
  const range = getSheet(workbook, sheetId).autoFilterRange;
  return range ? cloneRange(range) : null;
}

export function getSheetCharts(workbook: WorkbookModel, sheetId: string): SheetChart[] {
  return (getSheet(workbook, sheetId).charts ?? []).map(cloneSheetChart);
}

export function getSheetMerges(workbook: WorkbookModel, sheetId: string): SheetMerge[] {
  return (getSheet(workbook, sheetId).merges ?? []).map(cloneSheetMerge);
}

export function getCellMerge(workbook: WorkbookModel, sheetId: string, address: string): CellMergeInfo | null {
  const coord = parseCellAddress(normalizeAddress(address));
  const merge = (getSheet(workbook, sheetId).merges ?? []).find((candidate) =>
    isCoordInRange(coord, normalizeRange(candidate.range))
  );
  if (!merge) {
    return null;
  }

  const range = normalizeRange(merge.range);
  const rowSpan = range.end.row - range.start.row + 1;
  const columnSpan = range.end.column - range.start.column + 1;
  const base = {
    range: cloneRange(range),
    rowSpan,
    columnSpan
  };

  return coord.row === range.start.row && coord.column === range.start.column
    ? { role: "anchor", ...base }
    : { role: "covered", ...base, anchor: { ...range.start } };
}

export function getNamedRange(workbook: WorkbookModel, name: string): NamedRange | null {
  const normalizedName = normalizeNamedRangeLookup(name);
  const namedRange = (workbook.namedRanges ?? []).find((range) => normalizeNamedRangeLookup(range.name) === normalizedName);
  return namedRange ? cloneNamedRange(namedRange) : null;
}

export function getNamedRangeForSelection(
  workbook: WorkbookModel,
  sheetId: string,
  range: CellRange
): NamedRange | null {
  const normalizedRange = normalizeRange(range);
  const namedRange = (workbook.namedRanges ?? []).find(
    (candidate) => candidate.sheetId === sheetId && rangesEqual(candidate.range, normalizedRange)
  );
  return namedRange ? cloneNamedRange(namedRange) : null;
}

export function setCellContent(
  workbook: WorkbookModel,
  sheetId: string,
  address: string,
  content: CellContent
): WorkbookModel {
  const normalizedAddress = normalizeAddress(address);
  const coord = parseCellAddress(normalizedAddress);

  return updateSheet(workbook, sheetId, (sheet) => {
    const nextContent = content === "" ? null : content;
    const previousContent = sheet.cells[normalizedAddress] ?? null;
    if (previousContent === nextContent) {
      return sheet;
    }

    const cells = { ...sheet.cells };
    if (nextContent === null) {
      delete cells[normalizedAddress];
    } else {
      cells[normalizedAddress] = nextContent;
    }

    return {
      ...sheet,
      rowCount: Math.max(sheet.rowCount, coord.row + 1),
      columnCount: Math.max(sheet.columnCount, coord.column + 1),
      cells
    };
  });
}

export function clearRange(workbook: WorkbookModel, sheetId: string, range: CellRange): WorkbookModel {
  return updateSheet(workbook, sheetId, (sheet) => {
    const cells = { ...sheet.cells };
    const hyperlinks = { ...(sheet.hyperlinks ?? {}) };
    let changed = false;
    for (const address of getRangeAddresses(range)) {
      if (address in cells) {
        delete cells[address];
        changed = true;
      }
      if (address in hyperlinks) {
        delete hyperlinks[address];
        changed = true;
      }
    }
    if (!changed) {
      return sheet;
    }
    return { ...sheet, cells, hyperlinks };
  });
}

export function clearRangeAll(workbook: WorkbookModel, sheetId: string, range: CellRange): WorkbookModel {
  const normalizedClearRange = normalizeRange(range);
  return updateSheet(workbook, sheetId, (sheet) => {
    const cells = { ...sheet.cells };
    const hyperlinks = { ...(sheet.hyperlinks ?? {}) };
    const formats = { ...(sheet.formats ?? {}) };
    let changed = false;

    for (const address of getRangeAddresses(range)) {
      if (address in cells) {
        delete cells[address];
        changed = true;
      }
      if (address in hyperlinks) {
        delete hyperlinks[address];
        changed = true;
      }
      if (address in formats) {
        delete formats[address];
        changed = true;
      }
    }

    const conditionalFormats = sheet.conditionalFormats ?? [];
    const nextConditionalFormats = conditionalFormats.filter(
      (rule) => !rangesIntersect(normalizeRange(rule.range), normalizedClearRange)
    );
    if (nextConditionalFormats.length !== conditionalFormats.length) {
      changed = true;
    }

    if (!changed) {
      return sheet;
    }

    return {
      ...sheet,
      cells,
      hyperlinks,
      formats,
      conditionalFormats: nextConditionalFormats.map(cloneConditionalFormatRule)
    };
  });
}

export function clearCellFormats(workbook: WorkbookModel, sheetId: string, range: CellRange): WorkbookModel {
  const normalizedClearRange = normalizeRange(range);
  return updateSheet(workbook, sheetId, (sheet) => {
    const formats = { ...(sheet.formats ?? {}) };
    let formatsChanged = false;
    for (const address of getRangeAddresses(range)) {
      if (address in formats) {
        delete formats[address];
        formatsChanged = true;
      }
    }

    const conditionalFormats = sheet.conditionalFormats ?? [];
    const nextConditionalFormats = conditionalFormats.filter(
      (rule) => !rangesIntersect(normalizeRange(rule.range), normalizedClearRange)
    );
    const conditionalFormatsChanged = nextConditionalFormats.length !== conditionalFormats.length;

    if (!formatsChanged && !conditionalFormatsChanged) {
      return sheet;
    }
    return {
      ...sheet,
      formats,
      conditionalFormats: nextConditionalFormats.map(cloneConditionalFormatRule)
    };
  });
}

export function clearDirectCellFormats(workbook: WorkbookModel, sheetId: string, range: CellRange): WorkbookModel {
  return updateSheet(workbook, sheetId, (sheet) => {
    const formats = { ...(sheet.formats ?? {}) };
    let changed = false;

    for (const address of getRangeAddresses(range)) {
      if (address in formats) {
        delete formats[address];
        changed = true;
      }
    }

    return changed ? { ...sheet, formats } : sheet;
  });
}

export function clearCellHyperlinks(workbook: WorkbookModel, sheetId: string, range: CellRange): WorkbookModel {
  return updateSheet(workbook, sheetId, (sheet) => {
    const hyperlinks = { ...(sheet.hyperlinks ?? {}) };
    let changed = false;
    for (const address of getRangeAddresses(range)) {
      if (address in hyperlinks) {
        delete hyperlinks[address];
        changed = true;
      }
    }
    if (!changed) {
      return sheet;
    }
    return { ...sheet, hyperlinks };
  });
}

export function setCellFormat(
  workbook: WorkbookModel,
  sheetId: string,
  range: CellRange,
  format: CellFormat
): WorkbookModel {
  return updateSheet(workbook, sheetId, (sheet) => {
    const formats = { ...(sheet.formats ?? {}) };
    let changed = false;
    for (const address of getRangeAddresses(range)) {
      const nextFormat = {
        ...(formats[address] ?? {}),
        ...format
      };
      if (!cellFormatsEqual(formats[address] ?? {}, nextFormat)) {
        formats[address] = nextFormat;
        changed = true;
      }
    }
    if (!changed) {
      return sheet;
    }
    return { ...sheet, formats };
  });
}

export function setCellBorders(
  workbook: WorkbookModel,
  sheetId: string,
  range: CellRange,
  preset: BorderPreset,
  color = DEFAULT_BORDER.color
): WorkbookModel {
  const normalized = normalizeRange(range);
  const border: CellBorderSide = { style: "thin", color };

  return updateSheet(workbook, sheetId, (sheet) => {
    const formats = { ...(sheet.formats ?? {}) };
    let changed = false;

    for (const address of getRangeAddresses(normalized)) {
      const coord = parseCellAddress(address);
      const previousFormat = formats[address] ?? {};
      const nextFormat =
        preset === "none"
          ? withoutBorders(previousFormat)
          : withBorders(previousFormat, normalized, coord, preset, border);

      if (cellFormatsEqual(previousFormat, nextFormat)) {
        continue;
      }

      changed = true;
      if (isEmptyFormat(nextFormat)) {
        delete formats[address];
      } else {
        formats[address] = nextFormat;
      }
    }

    if (!changed) {
      return sheet;
    }
    return { ...sheet, formats };
  });
}

export function setColumnWidth(workbook: WorkbookModel, sheetId: string, column: number, width: number): WorkbookModel {
  const normalizedColumn = Math.max(0, Math.floor(column));
  const nextWidth = clampColumnWidth(width);

  return updateSheet(workbook, sheetId, (sheet) => {
    const columnWidths = { ...(sheet.columnWidths ?? {}) };
    const key = String(normalizedColumn);
    const previousWidth = columnWidths[key] ?? DEFAULT_COLUMN_WIDTH;
    if (previousWidth === nextWidth) {
      return sheet;
    }

    if (nextWidth === DEFAULT_COLUMN_WIDTH) {
      delete columnWidths[key];
    } else {
      columnWidths[key] = nextWidth;
    }

    return { ...sheet, columnWidths };
  });
}

export function setRowHeight(workbook: WorkbookModel, sheetId: string, row: number, height: number): WorkbookModel {
  const normalizedRow = Math.max(0, Math.floor(row));
  const nextHeight = clampRowHeight(height);

  return updateSheet(workbook, sheetId, (sheet) => {
    const rowHeights = { ...(sheet.rowHeights ?? {}) };
    const key = String(normalizedRow);
    const previousHeight = rowHeights[key] ?? DEFAULT_ROW_HEIGHT;
    if (previousHeight === nextHeight) {
      return sheet;
    }

    if (nextHeight === DEFAULT_ROW_HEIGHT) {
      delete rowHeights[key];
    } else {
      rowHeights[key] = nextHeight;
    }

    return { ...sheet, rowHeights };
  });
}

export function setColumnsHidden(
  workbook: WorkbookModel,
  sheetId: string,
  startColumn: number,
  endColumn: number,
  hidden: boolean
): WorkbookModel {
  return updateSheet(workbook, sheetId, (sheet) => {
    const nextHiddenColumns = setIndexFlags(sheet.hiddenColumns ?? {}, startColumn, endColumn, hidden, sheet.columnCount);
    return indexFlagsEqual(sheet.hiddenColumns ?? {}, nextHiddenColumns) ? sheet : { ...sheet, hiddenColumns: nextHiddenColumns };
  });
}

export function setRowsHidden(
  workbook: WorkbookModel,
  sheetId: string,
  startRow: number,
  endRow: number,
  hidden: boolean
): WorkbookModel {
  return updateSheet(workbook, sheetId, (sheet) => {
    const nextHiddenRows = setIndexFlags(sheet.hiddenRows ?? {}, startRow, endRow, hidden, sheet.rowCount);
    return indexFlagsEqual(sheet.hiddenRows ?? {}, nextHiddenRows) ? sheet : { ...sheet, hiddenRows: nextHiddenRows };
  });
}

export function clearHiddenRowsAndColumns(workbook: WorkbookModel, sheetId: string): WorkbookModel {
  return updateSheet(workbook, sheetId, (sheet) => {
    if (Object.keys(sheet.hiddenRows ?? {}).length === 0 && Object.keys(sheet.hiddenColumns ?? {}).length === 0) {
      return sheet;
    }
    return { ...sheet, hiddenRows: {}, hiddenColumns: {} };
  });
}

export function setSheetHidden(workbook: WorkbookModel, sheetId: string, hidden: boolean): WorkbookModel {
  const sheet = getSheet(workbook, sheetId);
  const nextHidden = hidden === true;
  if (Boolean(sheet.isHidden) === nextHidden) {
    return workbook;
  }

  if (nextHidden && visibleSheets(workbook.sheets).length <= 1) {
    return workbook;
  }

  const sheets = workbook.sheets.map((candidate) =>
    candidate.id === sheetId ? { ...candidate, isHidden: nextHidden } : candidate
  );
  const activeSheetId =
    workbook.activeSheetId === sheetId && nextHidden ? firstVisibleSheetId(sheets) ?? workbook.activeSheetId : workbook.activeSheetId;

  return {
    ...workbook,
    activeSheetId,
    sheets
  };
}

export function unhideAllSheets(workbook: WorkbookModel): WorkbookModel {
  if (!workbook.sheets.some((sheet) => sheet.isHidden === true)) {
    return workbook;
  }

  return {
    ...workbook,
    sheets: workbook.sheets.map((sheet) => (sheet.isHidden ? { ...sheet, isHidden: false } : sheet))
  };
}

export function setSheetTabColor(workbook: WorkbookModel, sheetId: string, color: string): WorkbookModel {
  const nextColor = normalizeHexColor(color);

  return updateSheet(workbook, sheetId, (sheet) => {
    const previousColor = normalizeHexColor(sheet.tabColor);
    if (previousColor === nextColor) {
      return sheet.tabColor === nextColor ? sheet : { ...sheet, tabColor: nextColor };
    }
    return { ...sheet, tabColor: nextColor };
  });
}

export function setCellComment(
  workbook: WorkbookModel,
  sheetId: string,
  address: string,
  comment: string | null
): WorkbookModel {
  const normalizedAddress = normalizeAddress(address);
  const coord = parseCellAddress(normalizedAddress);
  const nextComment = comment?.trim() ?? "";

  return updateSheet(workbook, sheetId, (sheet) => {
    const comments = { ...(sheet.comments ?? {}) };
    const previousComment = comments[normalizedAddress] ?? "";

    if (previousComment === nextComment) {
      return sheet;
    }

    if (nextComment) {
      comments[normalizedAddress] = nextComment;
    } else {
      delete comments[normalizedAddress];
    }

    return {
      ...sheet,
      rowCount: Math.max(sheet.rowCount, coord.row + 1),
      columnCount: Math.max(sheet.columnCount, coord.column + 1),
      comments
    };
  });
}

export function clearCellComments(workbook: WorkbookModel, sheetId: string, range: CellRange): WorkbookModel {
  return updateSheet(workbook, sheetId, (sheet) => {
    const comments = { ...(sheet.comments ?? {}) };
    let changed = false;
    for (const address of getRangeAddresses(range)) {
      if (address in comments) {
        delete comments[address];
        changed = true;
      }
    }
    if (!changed) {
      return sheet;
    }
    return { ...sheet, comments };
  });
}

export function setCellHyperlink(
  workbook: WorkbookModel,
  sheetId: string,
  address: string,
  url: string | null
): WorkbookModel {
  const normalizedAddress = normalizeAddress(address);
  const coord = parseCellAddress(normalizedAddress);
  const nextUrl = url?.trim() ?? "";

  return updateSheet(workbook, sheetId, (sheet) => {
    const hyperlinks = { ...(sheet.hyperlinks ?? {}) };
    const previousUrl = hyperlinks[normalizedAddress] ?? "";

    if (previousUrl === nextUrl) {
      return sheet;
    }

    if (nextUrl) {
      hyperlinks[normalizedAddress] = nextUrl;
    } else {
      delete hyperlinks[normalizedAddress];
    }

    return {
      ...sheet,
      rowCount: nextUrl ? Math.max(sheet.rowCount, coord.row + 1) : sheet.rowCount,
      columnCount: nextUrl ? Math.max(sheet.columnCount, coord.column + 1) : sheet.columnCount,
      hyperlinks
    };
  });
}

export function setSheetProtection(workbook: WorkbookModel, sheetId: string, isProtected: boolean): WorkbookModel {
  return updateSheet(workbook, sheetId, (sheet) => {
    const protection = getSheetProtection(sheet);
    if (protection.isProtected === isProtected) {
      return sheet;
    }

    return {
      ...sheet,
      protection: {
        ...protection,
        isProtected
      }
    };
  });
}

export function setSheetFreezePanes(
  workbook: WorkbookModel,
  sheetId: string,
  panes: { freezeTopRow: boolean; freezeFirstColumn: boolean }
): WorkbookModel {
  return updateSheet(workbook, sheetId, (sheet) => {
    const freezeTopRow = panes.freezeTopRow === true;
    const freezeFirstColumn = panes.freezeFirstColumn === true;
    if (Boolean(sheet.freezeTopRow) === freezeTopRow && Boolean(sheet.freezeFirstColumn) === freezeFirstColumn) {
      return sheet;
    }

    return {
      ...sheet,
      freezeTopRow,
      freezeFirstColumn
    };
  });
}

export function setRangeReadOnly(
  workbook: WorkbookModel,
  sheetId: string,
  range: CellRange,
  readOnly: boolean
): WorkbookModel {
  return updateSheet(workbook, sheetId, (sheet) => {
    const protection = getSheetProtection(sheet);
    const lockedCells = { ...protection.lockedCells };
    const unlockedCells = { ...protection.unlockedCells };
    let changed = false;

    for (const address of getRangeAddresses(range)) {
      if (readOnly) {
        if (!lockedCells[address]) {
          lockedCells[address] = true;
          changed = true;
        }
        if (unlockedCells[address]) {
          delete unlockedCells[address];
          changed = true;
        }
      } else {
        if (lockedCells[address]) {
          delete lockedCells[address];
          changed = true;
        }
        if (!unlockedCells[address]) {
          unlockedCells[address] = true;
          changed = true;
        }
      }
    }

    if (!changed) {
      return sheet;
    }

    return {
      ...sheet,
      protection: {
        ...protection,
        lockedCells,
        unlockedCells
      }
    };
  });
}

export function setCellValidation(
  workbook: WorkbookModel,
  sheetId: string,
  range: CellRange,
  rule: DataValidationRule | null
): WorkbookModel {
  return updateSheet(workbook, sheetId, (sheet) => {
    const validations = { ...(sheet.validations ?? {}) };
    let changed = false;

    for (const address of getRangeAddresses(range)) {
      if (rule === null) {
        if (address in validations) {
          delete validations[address];
          changed = true;
        }
        continue;
      }

      const nextRule = cloneValidationRule(rule);
      if (!dataValidationsEqual(validations[address], nextRule)) {
        validations[address] = nextRule;
        changed = true;
      }
    }

    if (!changed) {
      return sheet;
    }

    return { ...sheet, validations };
  });
}

export function addConditionalFormatRule(
  workbook: WorkbookModel,
  sheetId: string,
  range: CellRange,
  rule: Omit<ConditionalFormatRule, "id" | "range">
): WorkbookModel {
  return updateSheet(workbook, sheetId, (sheet) => {
    const conditionalFormats = sheet.conditionalFormats ?? [];
    const nextRule: ConditionalFormatRule = {
      id: nextConditionalFormatId(conditionalFormats),
      range: normalizeRange(range),
      condition: { ...rule.condition },
      format: { ...rule.format }
    };

    return { ...sheet, conditionalFormats: [...conditionalFormats.map(cloneConditionalFormatRule), nextRule] };
  });
}

export function clearConditionalFormatRules(workbook: WorkbookModel, sheetId: string, range: CellRange): WorkbookModel {
  const normalizedClearRange = normalizeRange(range);
  return updateSheet(workbook, sheetId, (sheet) => {
    const conditionalFormats = sheet.conditionalFormats ?? [];
    const nextRules = conditionalFormats.filter((rule) => !rangesIntersect(normalizeRange(rule.range), normalizedClearRange));
    if (nextRules.length === conditionalFormats.length) {
      return sheet;
    }

    return { ...sheet, conditionalFormats: nextRules.map(cloneConditionalFormatRule) };
  });
}

export function removeConditionalFormatRule(workbook: WorkbookModel, sheetId: string, ruleId: string): WorkbookModel {
  return updateSheet(workbook, sheetId, (sheet) => {
    const conditionalFormats = sheet.conditionalFormats ?? [];
    const nextRules = conditionalFormats.filter((rule) => rule.id !== ruleId);
    if (nextRules.length === conditionalFormats.length) {
      return sheet;
    }

    return { ...sheet, conditionalFormats: nextRules.map(cloneConditionalFormatRule) };
  });
}

export function addSheetFilter(
  workbook: WorkbookModel,
  sheetId: string,
  range: CellRange,
  filter: Omit<SheetFilter, "id" | "range">
): WorkbookModel {
  return updateSheet(workbook, sheetId, (sheet) => {
    const filters = sheet.filters ?? [];
    const normalizedRange = normalizeRange(range);
    const nextFilter: SheetFilter = {
      id: nextSheetFilterId(filters),
      range: normalizedRange,
      column: Math.min(Math.max(filter.column, normalizedRange.start.column), normalizedRange.end.column),
      operator: filter.operator,
      value: filter.value,
      values: filter.values ? uniqueFilterValues(filter.values) : undefined,
      hasHeader: filter.hasHeader
    };
    const nextFilters = filters.filter(
      (candidate) => !(rangesEqual(candidate.range, normalizedRange) && candidate.column === nextFilter.column)
    );

    return {
      ...sheet,
      autoFilterRange: cloneRange(normalizedRange),
      filters: [...nextFilters.map(cloneSheetFilter), nextFilter]
    };
  });
}

export function clearSheetFilter(
  workbook: WorkbookModel,
  sheetId: string,
  range: CellRange,
  column: number
): WorkbookModel {
  return updateSheet(workbook, sheetId, (sheet) => {
    const normalizedRange = normalizeRange(range);
    const clampedColumn = Math.min(Math.max(column, normalizedRange.start.column), normalizedRange.end.column);
    const filters = sheet.filters ?? [];
    const nextFilters = filters.filter(
      (filter) => !(rangesEqual(filter.range, normalizedRange) && filter.column === clampedColumn)
    );

    if (nextFilters.length === filters.length && sheet.autoFilterRange && rangesEqual(sheet.autoFilterRange, normalizedRange)) {
      return sheet;
    }

    return {
      ...sheet,
      autoFilterRange: cloneRange(normalizedRange),
      filters: nextFilters.map(cloneSheetFilter)
    };
  });
}

export function clearSheetFilters(workbook: WorkbookModel, sheetId: string): WorkbookModel {
  return updateSheet(workbook, sheetId, (sheet) => {
    if ((sheet.filters ?? []).length === 0 && !sheet.autoFilterRange) {
      return sheet;
    }

    return { ...sheet, autoFilterRange: undefined, filters: [] };
  });
}

export function addSheetChart(
  workbook: WorkbookModel,
  sheetId: string,
  range: CellRange,
  chart: Omit<SheetChart, "id" | "range">
): WorkbookModel {
  return updateSheet(workbook, sheetId, (sheet) => {
    const charts = sheet.charts ?? [];
    const nextChart: SheetChart = {
      id: nextSheetChartId(charts),
      range: normalizeRange(range),
      anchor: { ...chart.anchor },
      title: chart.title.trim() || "Chart",
      type: chart.type
    };

    return { ...sheet, charts: [...charts.map(cloneSheetChart), nextChart] };
  });
}

export function deleteSheetChart(workbook: WorkbookModel, sheetId: string, chartId: string): WorkbookModel {
  return updateSheet(workbook, sheetId, (sheet) => {
    const charts = sheet.charts ?? [];
    const nextCharts = charts.filter((chart) => chart.id !== chartId);
    if (nextCharts.length === charts.length) {
      return sheet;
    }

    return { ...sheet, charts: nextCharts.map(cloneSheetChart) };
  });
}

export function mergeCells(workbook: WorkbookModel, sheetId: string, range: CellRange): WorkbookModel {
  const normalizedRange = normalizeRange(range);
  const isSingleCell =
    normalizedRange.start.row === normalizedRange.end.row && normalizedRange.start.column === normalizedRange.end.column;
  if (isSingleCell) {
    return workbook;
  }

  return updateSheet(workbook, sheetId, (sheet) => {
    const merges = sheet.merges ?? [];
    const nextMerge: SheetMerge = {
      id: nextSheetMergeId(merges),
      range: cloneRange(normalizedRange)
    };
    const nextMerges = [
      ...merges.filter((merge) => !rangesIntersect(normalizeRange(merge.range), normalizedRange)).map(cloneSheetMerge),
      nextMerge
    ];
    const cells = { ...sheet.cells };
    for (const address of getRangeAddresses(normalizedRange)) {
      const coord = parseCellAddress(address);
      if (coord.row !== normalizedRange.start.row || coord.column !== normalizedRange.start.column) {
        delete cells[address];
      }
    }

    return {
      ...sheet,
      rowCount: Math.max(sheet.rowCount, normalizedRange.end.row + 1),
      columnCount: Math.max(sheet.columnCount, normalizedRange.end.column + 1),
      cells,
      merges: nextMerges
    };
  });
}

export function unmergeCells(workbook: WorkbookModel, sheetId: string, range: CellRange): WorkbookModel {
  const normalizedRange = normalizeRange(range);
  return updateSheet(workbook, sheetId, (sheet) => {
    const merges = sheet.merges ?? [];
    const nextMerges = merges.filter((merge) => !rangesIntersect(normalizeRange(merge.range), normalizedRange));
    if (nextMerges.length === merges.length) {
      return sheet;
    }

    return {
      ...sheet,
      merges: nextMerges.map(cloneSheetMerge)
    };
  });
}

export function isValidNamedRangeName(name: string): boolean {
  const trimmedName = name.trim();
  if (!/^[A-Za-z_][A-Za-z0-9_]*$/.test(trimmedName)) {
    return false;
  }

  try {
    parseCellAddress(trimmedName);
    return false;
  } catch {
    return true;
  }
}

export function defineNamedRange(
  workbook: WorkbookModel,
  sheetId: string,
  name: string,
  range: CellRange
): WorkbookModel {
  getSheet(workbook, sheetId);
  const trimmedName = name.trim();
  if (!isValidNamedRangeName(trimmedName)) {
    return workbook;
  }

  const normalizedRange = normalizeRange(range);
  const namedRanges = workbook.namedRanges ?? [];
  const existingIndex = namedRanges.findIndex(
    (namedRange) => normalizeNamedRangeLookup(namedRange.name) === normalizeNamedRangeLookup(trimmedName)
  );
  const nextNamedRange: NamedRange = {
    name: trimmedName,
    sheetId,
    range: cloneRange(normalizedRange)
  };

  if (existingIndex >= 0) {
    const existingRange = namedRanges[existingIndex];
    if (
      existingRange.name === nextNamedRange.name &&
      existingRange.sheetId === nextNamedRange.sheetId &&
      rangesEqual(existingRange.range, nextNamedRange.range)
    ) {
      return workbook;
    }

    return {
      ...workbook,
      namedRanges: namedRanges.map((namedRange, index) =>
        index === existingIndex ? nextNamedRange : cloneNamedRange(namedRange)
      )
    };
  }

  return {
    ...workbook,
    namedRanges: [...namedRanges.map(cloneNamedRange), nextNamedRange]
  };
}

export function removeNamedRange(workbook: WorkbookModel, name: string): WorkbookModel {
  const normalizedName = normalizeNamedRangeLookup(name);
  const namedRanges = workbook.namedRanges ?? [];
  const nextNamedRanges = namedRanges.filter(
    (namedRange) => normalizeNamedRangeLookup(namedRange.name) !== normalizedName
  );

  if (nextNamedRanges.length === namedRanges.length) {
    return workbook;
  }

  return {
    ...workbook,
    namedRanges: nextNamedRanges.map(cloneNamedRange)
  };
}

export function copyRange(workbook: WorkbookModel, sheetId: string, range: CellRange): CellContent[][] {
  const normalized = normalizeRange(range);
  const matrix: CellContent[][] = [];

  for (let row = normalized.start.row; row <= normalized.end.row; row += 1) {
    const values: CellContent[] = [];
    for (let column = normalized.start.column; column <= normalized.end.column; column += 1) {
      values.push(getCellContent(workbook, sheetId, formatCellAddress({ row, column })));
    }
    matrix.push(values);
  }

  return matrix;
}

export function copyRichRange(
  workbook: WorkbookModel,
  sheetId: string,
  range: CellRange,
  options: CopyRichRangeOptions = {}
): RichClipboardRange {
  const sheet = getSheet(workbook, sheetId);
  const normalized = normalizeRange(range);
  const cells: RichClipboardCell[][] = [];

  for (let row = normalized.start.row; row <= normalized.end.row; row += 1) {
    const rowCells: RichClipboardCell[] = [];
    for (let column = normalized.start.column; column <= normalized.end.column; column += 1) {
      const sourceAddress = formatCellAddress({ row, column });
      const validation = (sheet.validations ?? {})[sourceAddress];
      rowCells.push({
        sourceAddress,
        content: sheet.cells[sourceAddress] ?? null,
        displayContent: options.getDisplayValue?.(sourceAddress) ?? sheet.cells[sourceAddress] ?? null,
        format: cloneCellFormat((sheet.formats ?? {})[sourceAddress] ?? {}),
        validation: validation ? cloneValidationRule(validation) : null,
        comment: (sheet.comments ?? {})[sourceAddress] ?? null,
        hyperlink: (sheet.hyperlinks ?? {})[sourceAddress] ?? null
      });
    }
    cells.push(rowCells);
  }

  return { range: cloneRange(normalized), cells };
}

export function previewRichPaste(
  clipboard: RichClipboardRange,
  startAddress: string,
  options: PasteRichRangeOptions = {}
): RichPasteCellPreview[] {
  const start = parseCellAddress(startAddress);
  const mode = options.mode ?? "all";
  const isTransposed = mode === "transpose";
  const translateFormulas = options.translateFormulas ?? true;

  return clipboard.cells.flatMap((rowCells, rowOffset) =>
    rowCells.map((cell, columnOffset) => {
      const target = isTransposed
        ? { row: start.row + columnOffset, column: start.column + rowOffset }
        : { row: start.row + rowOffset, column: start.column + columnOffset };
      const source = parseCellAddress(cell.sourceAddress);
      return {
        address: formatCellAddress(target),
        content:
          mode === "values"
            ? cell.displayContent
            : translateFormulas
              ? translateFillContent(cell.content, {
                  rowOffset: target.row - source.row,
                  columnOffset: target.column - source.column
                })
              : cell.content,
        format: cloneCellFormat(cell.format),
        validation: cell.validation ? cloneValidationRule(cell.validation) : null,
        comment: cell.comment,
        hyperlink: cell.hyperlink
      };
    })
  );
}

export function pasteRichRange(
  workbook: WorkbookModel,
  sheetId: string,
  startAddress: string,
  clipboard: RichClipboardRange,
  options: PasteRichRangeOptions = {}
): WorkbookModel {
  const mode = options.mode ?? "all";
  const preview = previewRichPaste(clipboard, startAddress, options);

  return updateSheet(workbook, sheetId, (sheet) => {
    const cells = { ...sheet.cells };
    const formats = { ...(sheet.formats ?? {}) };
    const validations = { ...(sheet.validations ?? {}) };
    const comments = { ...(sheet.comments ?? {}) };
    const hyperlinks = { ...(sheet.hyperlinks ?? {}) };
    let changed = false;
    let rowCount = sheet.rowCount;
    let columnCount = sheet.columnCount;

    for (const item of preview) {
      const coord = parseCellAddress(item.address);
      const nextContent = item.content === "" ? null : item.content;
      const previousContent = cells[item.address] ?? null;
      const nextFormat = cloneCellFormat(item.format);
      const nextComment = item.comment?.trim() ?? "";
      const nextHyperlink = item.hyperlink?.trim() ?? "";

      rowCount = Math.max(rowCount, coord.row + 1);
      columnCount = Math.max(columnCount, coord.column + 1);

      if (mode !== "formats" && previousContent !== nextContent) {
        changed = true;
        if (nextContent === null) {
          delete cells[item.address];
        } else {
          cells[item.address] = nextContent;
        }
      }

      if (mode !== "values" && !cellFormatsEqual(formats[item.address] ?? {}, nextFormat)) {
        changed = true;
        if (isEmptyFormat(nextFormat)) {
          delete formats[item.address];
        } else {
          formats[item.address] = nextFormat;
        }
      }

      if (mode !== "values" && mode !== "formats" && !validationRulesEqual(validations[item.address], item.validation)) {
        changed = true;
        if (item.validation) {
          validations[item.address] = cloneValidationRule(item.validation);
        } else {
          delete validations[item.address];
        }
      }

      if (mode !== "values" && mode !== "formats" && (comments[item.address] ?? "") !== nextComment) {
        changed = true;
        if (nextComment) {
          comments[item.address] = nextComment;
        } else {
          delete comments[item.address];
        }
      }

      if (mode !== "values" && mode !== "formats" && (hyperlinks[item.address] ?? "") !== nextHyperlink) {
        changed = true;
        if (nextHyperlink) {
          hyperlinks[item.address] = nextHyperlink;
        } else {
          delete hyperlinks[item.address];
        }
      }
    }

    if (!changed && rowCount === sheet.rowCount && columnCount === sheet.columnCount) {
      return sheet;
    }

    return {
      ...sheet,
      rowCount,
      columnCount,
      cells,
      formats,
      validations,
      comments,
      hyperlinks
    };
  });
}

export function moveRichRange(
  workbook: WorkbookModel,
  sourceSheetId: string,
  targetSheetId: string,
  startAddress: string,
  clipboard: RichClipboardRange
): WorkbookModel {
  const preview = previewRichPaste(clipboard, startAddress, { translateFormulas: false });
  const preservedSourceAddresses =
    sourceSheetId === targetSheetId ? new Set(preview.map((item) => item.address)) : new Set<string>();
  const pasted = pasteRichRange(workbook, targetSheetId, startAddress, clipboard, { translateFormulas: false });

  return clearMovedSourceRange(pasted, sourceSheetId, clipboard.range, preservedSourceAddresses);
}

function clearMovedSourceRange(
  workbook: WorkbookModel,
  sheetId: string,
  range: CellRange,
  preservedAddresses: Set<string>
): WorkbookModel {
  return updateSheet(workbook, sheetId, (sheet) => {
    const cells = { ...sheet.cells };
    const formats = { ...(sheet.formats ?? {}) };
    const validations = { ...(sheet.validations ?? {}) };
    const comments = { ...(sheet.comments ?? {}) };
    const hyperlinks = { ...(sheet.hyperlinks ?? {}) };
    let changed = false;

    for (const address of getRangeAddresses(range)) {
      if (preservedAddresses.has(address)) {
        continue;
      }
      if (address in cells) {
        delete cells[address];
        changed = true;
      }
      if (address in formats) {
        delete formats[address];
        changed = true;
      }
      if (address in validations) {
        delete validations[address];
        changed = true;
      }
      if (address in comments) {
        delete comments[address];
        changed = true;
      }
      if (address in hyperlinks) {
        delete hyperlinks[address];
        changed = true;
      }
    }

    if (!changed) {
      return sheet;
    }

    return { ...sheet, cells, formats, validations, comments, hyperlinks };
  });
}

export function pasteMatrix(
  workbook: WorkbookModel,
  sheetId: string,
  startAddress: string,
  matrix: CellContent[][]
): WorkbookModel {
  const start = parseCellAddress(startAddress);

  return updateSheet(workbook, sheetId, (sheet) => {
    const cells = { ...sheet.cells };
    let changed = false;
    let rowCount = sheet.rowCount;
    let columnCount = sheet.columnCount;

    matrix.forEach((rowValues, rowOffset) => {
      rowValues.forEach((content, columnOffset) => {
        const row = start.row + rowOffset;
        const column = start.column + columnOffset;
        const address = formatCellAddress({ row, column });
        const nextContent = content === "" ? null : content;
        const previousContent = cells[address] ?? null;

        rowCount = Math.max(rowCount, row + 1);
        columnCount = Math.max(columnCount, column + 1);

        if (previousContent === nextContent) {
          return;
        }

        changed = true;
        if (nextContent === null) {
          delete cells[address];
        } else {
          cells[address] = nextContent;
        }
      });
    });

    if (!changed && rowCount === sheet.rowCount && columnCount === sheet.columnCount) {
      return sheet;
    }

    return {
      ...sheet,
      rowCount,
      columnCount,
      cells
    };
  });
}

export function insertRows(workbook: WorkbookModel, sheetId: string, rowIndex: number, count = 1): WorkbookModel {
  return shiftSheetStructure(workbook, sheetId, { axis: "row", mode: "insert", index: rowIndex, count });
}

export function deleteRows(workbook: WorkbookModel, sheetId: string, rowIndex: number, count = 1): WorkbookModel {
  return shiftSheetStructure(workbook, sheetId, { axis: "row", mode: "delete", index: rowIndex, count });
}

export function insertColumns(workbook: WorkbookModel, sheetId: string, columnIndex: number, count = 1): WorkbookModel {
  return shiftSheetStructure(workbook, sheetId, { axis: "column", mode: "insert", index: columnIndex, count });
}

export function deleteColumns(workbook: WorkbookModel, sheetId: string, columnIndex: number, count = 1): WorkbookModel {
  return shiftSheetStructure(workbook, sheetId, { axis: "column", mode: "delete", index: columnIndex, count });
}

export function fillDown(workbook: WorkbookModel, sheetId: string, range: CellRange): WorkbookModel {
  const normalized = normalizeRange(range);
  let nextWorkbook = workbook;

  for (let column = normalized.start.column; column <= normalized.end.column; column += 1) {
    const source = getCellContent(workbook, sheetId, formatCellAddress({ row: normalized.start.row, column }));
    for (let row = normalized.start.row + 1; row <= normalized.end.row; row += 1) {
      nextWorkbook = setCellContent(
        nextWorkbook,
        sheetId,
        formatCellAddress({ row, column }),
        translateFillContent(source, { rowOffset: row - normalized.start.row, columnOffset: 0 })
      );
    }
  }

  return nextWorkbook;
}

export function fillRight(workbook: WorkbookModel, sheetId: string, range: CellRange): WorkbookModel {
  const normalized = normalizeRange(range);
  let nextWorkbook = workbook;

  for (let row = normalized.start.row; row <= normalized.end.row; row += 1) {
    const source = getCellContent(workbook, sheetId, formatCellAddress({ row, column: normalized.start.column }));
    for (let column = normalized.start.column + 1; column <= normalized.end.column; column += 1) {
      nextWorkbook = setCellContent(
        nextWorkbook,
        sheetId,
        formatCellAddress({ row, column }),
        translateFillContent(source, { rowOffset: 0, columnOffset: column - normalized.start.column })
      );
    }
  }

  return nextWorkbook;
}

export function autoFillRange(
  workbook: WorkbookModel,
  sheetId: string,
  sourceRange: CellRange,
  targetRange: CellRange
): WorkbookModel {
  const source = normalizeRange(sourceRange);
  const target = normalizeRange(targetRange);

  if (isDownAutoFill(source, target)) {
    return autoFillDown(workbook, sheetId, source, target);
  }

  if (isRightAutoFill(source, target)) {
    return autoFillRight(workbook, sheetId, source, target);
  }

  return workbook;
}

export function sortRange(
  workbook: WorkbookModel,
  sheetId: string,
  range: CellRange,
  direction: "asc" | "desc",
  sortColumn?: number
): WorkbookModel {
  const normalized = normalizeRange(range);
  if (normalized.start.row === normalized.end.row) {
    return workbook;
  }

  return updateSheet(workbook, sheetId, (sheet) => {
    const sortableRange = sortBodyRange(sheet, normalized);
    if (sortableRange.start.row === sortableRange.end.row) {
      return sheet;
    }
    const keyColumn = Math.min(
      Math.max(sortColumn ?? normalized.start.column, sortableRange.start.column),
      sortableRange.end.column
    );

    const rows = Array.from({ length: sortableRange.end.row - sortableRange.start.row + 1 }, (_, rowOffset) => {
      const row = sortableRange.start.row + rowOffset;
      return {
        originalIndex: rowOffset,
        key: sheet.cells[formatCellAddress({ row, column: keyColumn })] ?? null,
        cells: Array.from({ length: sortableRange.end.column - sortableRange.start.column + 1 }, (_, columnOffset) => {
          const column = sortableRange.start.column + columnOffset;
          return sheet.cells[formatCellAddress({ row, column })] ?? null;
        }),
        formats: Array.from({ length: sortableRange.end.column - sortableRange.start.column + 1 }, (_, columnOffset) => {
          const column = sortableRange.start.column + columnOffset;
          return cloneCellFormat((sheet.formats ?? {})[formatCellAddress({ row, column })] ?? {});
        }),
        validations: Array.from({ length: sortableRange.end.column - sortableRange.start.column + 1 }, (_, columnOffset) => {
          const column = sortableRange.start.column + columnOffset;
          const rule = (sheet.validations ?? {})[formatCellAddress({ row, column })];
          return rule ? cloneValidationRule(rule) : null;
        }),
        comments: Array.from({ length: sortableRange.end.column - sortableRange.start.column + 1 }, (_, columnOffset) => {
          const column = sortableRange.start.column + columnOffset;
          return (sheet.comments ?? {})[formatCellAddress({ row, column })] ?? null;
        }),
        hyperlinks: Array.from({ length: sortableRange.end.column - sortableRange.start.column + 1 }, (_, columnOffset) => {
          const column = sortableRange.start.column + columnOffset;
          return (sheet.hyperlinks ?? {})[formatCellAddress({ row, column })] ?? null;
        })
      };
    });

    const sortedRows = [...rows].sort((left, right) => compareSortRows(left, right, direction));
    if (sortedRows.every((row, index) => row.originalIndex === index)) {
      return sheet;
    }

    const cells = { ...sheet.cells };
    const formats = { ...(sheet.formats ?? {}) };
    const validations = { ...(sheet.validations ?? {}) };
    const comments = { ...(sheet.comments ?? {}) };
    const hyperlinks = { ...(sheet.hyperlinks ?? {}) };
    sortedRows.forEach((row, rowOffset) => {
      row.cells.forEach((content, columnOffset) => {
        const address = formatCellAddress({
          row: sortableRange.start.row + rowOffset,
          column: sortableRange.start.column + columnOffset
        });
        if (content === null) {
          delete cells[address];
        } else {
          cells[address] = content;
        }

        const nextFormat = row.formats[columnOffset];
        if (Object.keys(nextFormat).length > 0) {
          formats[address] = nextFormat;
        } else {
          delete formats[address];
        }

        const nextValidation = row.validations[columnOffset];
        if (nextValidation) {
          validations[address] = cloneValidationRule(nextValidation);
        } else {
          delete validations[address];
        }

        const nextComment = row.comments[columnOffset];
        if (nextComment) {
          comments[address] = nextComment;
        } else {
          delete comments[address];
        }

        const nextHyperlink = row.hyperlinks[columnOffset];
        if (nextHyperlink) {
          hyperlinks[address] = nextHyperlink;
        } else {
          delete hyperlinks[address];
        }
      });
    });

    return { ...sheet, cells, formats, validations, comments, hyperlinks };
  });
}

export function removeDuplicateRows(workbook: WorkbookModel, sheetId: string, range: CellRange): RemoveDuplicateRowsResult {
  const normalized = normalizeRange(range);
  let removedCount = 0;

  const nextWorkbook = updateSheet(workbook, sheetId, (sheet) => {
    const rows = collectRangeRows(sheet, normalized);
    const seen = new Set<string>();
    const uniqueRows: RangeRowSnapshot[] = [];

    for (const row of rows) {
      if (seen.has(row.key)) {
        removedCount += 1;
        continue;
      }
      seen.add(row.key);
      uniqueRows.push(row);
    }

    if (removedCount === 0) {
      return sheet;
    }

    const cells = { ...sheet.cells };
    const formats = { ...(sheet.formats ?? {}) };
    const validations = { ...(sheet.validations ?? {}) };
    const comments = { ...(sheet.comments ?? {}) };
    const hyperlinks = { ...(sheet.hyperlinks ?? {}) };
    const columnCount = normalized.end.column - normalized.start.column + 1;

    for (let rowOffset = 0; rowOffset < rows.length; rowOffset += 1) {
      const sourceRow = uniqueRows[rowOffset] ?? null;
      for (let columnOffset = 0; columnOffset < columnCount; columnOffset += 1) {
        const address = formatCellAddress({
          row: normalized.start.row + rowOffset,
          column: normalized.start.column + columnOffset
        });

        if (!sourceRow) {
          delete cells[address];
          delete formats[address];
          delete validations[address];
          delete comments[address];
          delete hyperlinks[address];
          continue;
        }

        const content = sourceRow.cells[columnOffset] ?? null;
        if (content === null) {
          delete cells[address];
        } else {
          cells[address] = content;
        }

        const nextFormat = sourceRow.formats[columnOffset];
        if (nextFormat && !isEmptyFormat(nextFormat)) {
          formats[address] = cloneCellFormat(nextFormat);
        } else {
          delete formats[address];
        }

        const nextValidation = sourceRow.validations[columnOffset];
        if (nextValidation) {
          validations[address] = cloneValidationRule(nextValidation);
        } else {
          delete validations[address];
        }

        const nextComment = sourceRow.comments[columnOffset]?.trim() ?? "";
        if (nextComment) {
          comments[address] = nextComment;
        } else {
          delete comments[address];
        }

        const nextHyperlink = sourceRow.hyperlinks[columnOffset]?.trim() ?? "";
        if (nextHyperlink) {
          hyperlinks[address] = nextHyperlink;
        } else {
          delete hyperlinks[address];
        }
      }
    }

    return { ...sheet, cells, formats, validations, comments, hyperlinks };
  });

  return { workbook: nextWorkbook, removedCount };
}

export function addSheet(workbook: WorkbookModel, name = `Sheet${workbook.sheets.length + 1}`): WorkbookModel {
  const sheet = createSheet(nextSheetId(workbook), uniqueSheetName(workbook, name));
  return {
    ...workbook,
    activeSheetId: sheet.id,
    sheets: [...workbook.sheets, sheet]
  };
}

export function renameSheet(workbook: WorkbookModel, sheetId: string, name: string): WorkbookModel {
  const trimmedName = name.trim();
  if (!trimmedName) {
    return workbook;
  }

  return updateSheet(workbook, sheetId, (sheet) => {
    const nextName = uniqueSheetName(workbook, trimmedName, sheetId);
    if (sheet.name === nextName) {
      return sheet;
    }
    return { ...sheet, name: nextName };
  });
}

export function duplicateSheet(workbook: WorkbookModel, sheetId: string): WorkbookModel {
  const source = getSheet(workbook, sheetId);
  const copy: SheetModel = {
    ...source,
    id: nextSheetId(workbook),
    name: uniqueSheetName(workbook, `${source.name} Copy`),
    isHidden: false,
    cells: { ...source.cells },
    formats: cloneCellFormats(source.formats ?? {}),
    columnWidths: { ...(source.columnWidths ?? {}) },
    rowHeights: { ...(source.rowHeights ?? {}) },
    hiddenColumns: { ...(source.hiddenColumns ?? {}) },
    hiddenRows: { ...(source.hiddenRows ?? {}) },
    autoFilterRange: source.autoFilterRange ? cloneRange(source.autoFilterRange) : undefined,
    comments: { ...(source.comments ?? {}) },
    hyperlinks: { ...(source.hyperlinks ?? {}) },
    validations: cloneValidations(source.validations ?? {}),
    conditionalFormats: (source.conditionalFormats ?? []).map(cloneConditionalFormatRule),
    filters: (source.filters ?? []).map(cloneSheetFilter),
    charts: (source.charts ?? []).map(cloneSheetChart),
    merges: (source.merges ?? []).map(cloneSheetMerge),
    protection: cloneSheetProtection(getSheetProtection(source))
  };

  return {
    ...workbook,
    activeSheetId: copy.id,
    sheets: [...workbook.sheets, copy]
  };
}

export function deleteSheet(workbook: WorkbookModel, sheetId: string): WorkbookModel {
  if (workbook.sheets.length <= 1) {
    return workbook;
  }

  const sheets = workbook.sheets.filter((sheet) => sheet.id !== sheetId);
  if (sheets.length === workbook.sheets.length) {
    return workbook;
  }

  return {
    ...workbook,
    activeSheetId: workbook.activeSheetId === sheetId ? firstVisibleSheetId(sheets) ?? sheets[0].id : workbook.activeSheetId,
    sheets,
    namedRanges: (workbook.namedRanges ?? []).filter((namedRange) => namedRange.sheetId !== sheetId).map(cloneNamedRange)
  };
}

export function moveSheet(workbook: WorkbookModel, sheetId: string, targetIndex: number): WorkbookModel {
  const currentIndex = workbook.sheets.findIndex((sheet) => sheet.id === sheetId);
  if (currentIndex === -1) {
    return workbook;
  }

  const clampedIndex = clampStructureIndex(targetIndex, 0, workbook.sheets.length - 1);
  if (currentIndex === clampedIndex) {
    return workbook;
  }

  const sheets = [...workbook.sheets];
  const [sheet] = sheets.splice(currentIndex, 1);
  sheets.splice(clampedIndex, 0, sheet);

  return {
    ...workbook,
    sheets
  };
}

export function setActiveSheet(workbook: WorkbookModel, sheetId: string): WorkbookModel {
  const sheet = getSheet(workbook, sheetId);
  if (sheet.isHidden) {
    return workbook;
  }
  if (workbook.activeSheetId === sheetId) {
    return workbook;
  }
  return { ...workbook, activeSheetId: sheetId };
}

// Each history entry pins a workbook snapshot. Snapshots share unchanged sheets
// and cell values structurally, but the edited sheet's cells record (its keys) is
// a fresh copy per edit — on a 100k-cell sheet that is megabytes per entry, so an
// unbounded past grows without limit. 100 undo steps matches Excel's default.
const MAX_UNDO_HISTORY = 100;

export function createHistory(initial: WorkbookModel): HistoryState {
  return {
    past: [],
    present: initial,
    future: []
  };
}

export function commitHistory(history: HistoryState, present: WorkbookModel): HistoryState {
  if (history.present === present) {
    return history;
  }

  const past = [...history.past, history.present];
  return {
    past: past.length > MAX_UNDO_HISTORY ? past.slice(past.length - MAX_UNDO_HISTORY) : past,
    present,
    future: []
  };
}

export function undoHistory(history: HistoryState): HistoryState {
  const previous = history.past.at(-1);
  if (!previous) {
    return history;
  }

  return {
    past: history.past.slice(0, -1),
    present: previous,
    future: [history.present, ...history.future]
  };
}

export function redoHistory(history: HistoryState): HistoryState {
  const next = history.future[0];
  if (!next) {
    return history;
  }

  return {
    past: [...history.past, history.present],
    present: next,
    future: history.future.slice(1)
  };
}

function createSheet(id: string, name: string): SheetModel {
  return {
    id,
    name,
    rowCount: DEFAULT_ROWS,
    columnCount: DEFAULT_COLUMNS,
    isHidden: false,
    tabColor: undefined,
    cells: {},
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
    protection: createDefaultProtection()
  };
}

function visibleSheets(sheets: readonly SheetModel[]): SheetModel[] {
  return sheets.filter((sheet) => sheet.isHidden !== true);
}

function firstVisibleSheetId(sheets: readonly SheetModel[]): string | undefined {
  return visibleSheets(sheets)[0]?.id;
}

function getSheet(workbook: WorkbookModel, sheetId: string): SheetModel {
  const sheet = workbook.sheets.find((candidate) => candidate.id === sheetId);
  if (!sheet) {
    throw new Error(`Unknown sheet: ${sheetId}`);
  }
  return sheet;
}

function updateSheet(
  workbook: WorkbookModel,
  sheetId: string,
  updater: (sheet: SheetModel) => SheetModel
): WorkbookModel {
  let changed = false;
  const sheets = workbook.sheets.map((sheet) => {
    if (sheet.id !== sheetId) {
      return sheet;
    }

    const nextSheet = updater(sheet);
    if (nextSheet !== sheet) {
      changed = true;
    }
    return nextSheet;
  });

  if (!changed) {
    return workbook;
  }

  return {
    ...workbook,
    sheets
  };
}

function shiftSheetStructure(workbook: WorkbookModel, sheetId: string, operation: StructureOperation): WorkbookModel {
  const count = Math.max(1, Math.floor(operation.count));
  const sourceSheet = getSheet(workbook, sheetId);
  const limit = operation.axis === "row" ? sourceSheet.rowCount : sourceSheet.columnCount;
  const index =
    operation.mode === "insert"
      ? clampStructureIndex(operation.index, 0, limit)
      : clampStructureIndex(operation.index, 0, Math.max(limit - 1, 0));
  const normalizedOperation = { ...operation, index, count };

  const nextWorkbook = updateSheet(workbook, sheetId, (sheet) => {
    const nextCells: Record<string, CellContent> = {};
    for (const [address, content] of Object.entries(sheet.cells)) {
      const nextCoord = shiftCoord(parseCellAddress(address), normalizedOperation);
      if (!nextCoord) {
        continue;
      }

      const nextContent = shiftFormulaReferences(content, normalizedOperation);
      if (nextContent !== null) {
        nextCells[formatCellAddress(nextCoord)] = nextContent;
      }
    }

    const nextFormats: Record<string, CellFormat> = {};
    for (const [address, format] of Object.entries(sheet.formats ?? {})) {
      const nextCoord = shiftCoord(parseCellAddress(address), normalizedOperation);
      if (nextCoord) {
        nextFormats[formatCellAddress(nextCoord)] = cloneCellFormat(format);
      }
    }

    const nextColumnWidths = shiftIndexedNumberMap(sheet.columnWidths ?? {}, normalizedOperation, "column", clampColumnWidth);
    const nextRowHeights = shiftIndexedNumberMap(sheet.rowHeights ?? {}, normalizedOperation, "row", clampRowHeight);
    const nextHiddenColumns = shiftIndexFlags(sheet.hiddenColumns ?? {}, normalizedOperation, "column");
    const nextHiddenRows = shiftIndexFlags(sheet.hiddenRows ?? {}, normalizedOperation, "row");

    const nextValidations: Record<string, DataValidationRule> = {};
    for (const [address, rule] of Object.entries(sheet.validations ?? {})) {
      const nextCoord = shiftCoord(parseCellAddress(address), normalizedOperation);
      if (nextCoord) {
        nextValidations[formatCellAddress(nextCoord)] = cloneValidationRule(rule);
      }
    }

    const nextComments: Record<string, string> = {};
    for (const [address, comment] of Object.entries(sheet.comments ?? {})) {
      const nextCoord = shiftCoord(parseCellAddress(address), normalizedOperation);
      if (nextCoord) {
        nextComments[formatCellAddress(nextCoord)] = comment;
      }
    }

    const nextHyperlinks: Record<string, string> = {};
    for (const [address, hyperlink] of Object.entries(sheet.hyperlinks ?? {})) {
      const nextCoord = shiftCoord(parseCellAddress(address), normalizedOperation);
      if (nextCoord) {
        nextHyperlinks[formatCellAddress(nextCoord)] = hyperlink;
      }
    }

    const protection = getSheetProtection(sheet);
    const nextProtection: SheetProtection = {
      ...protection,
      lockedCells: shiftAddressFlags(protection.lockedCells, normalizedOperation),
      unlockedCells: shiftAddressFlags(protection.unlockedCells, normalizedOperation)
    };

    const nextConditionalFormats = (sheet.conditionalFormats ?? []).flatMap((rule) => {
      const shiftedRange = shiftRange(rule.range, normalizedOperation);
      return shiftedRange ? [{ ...cloneConditionalFormatRule(rule), range: shiftedRange }] : [];
    });

    const nextFilters = (sheet.filters ?? []).flatMap((filter) => {
      const shiftedRange = shiftRange(filter.range, normalizedOperation);
      if (!shiftedRange) {
        return [];
      }

      const shiftedColumn =
        normalizedOperation.axis === "column" ? shiftCoord({ row: 0, column: filter.column }, normalizedOperation)?.column : filter.column;
      if (shiftedColumn === undefined || shiftedColumn < shiftedRange.start.column || shiftedColumn > shiftedRange.end.column) {
        return [];
      }

      return [{ ...cloneSheetFilter(filter), range: shiftedRange, column: shiftedColumn }];
    });
    const shiftedAutoFilterRange = sheet.autoFilterRange ? shiftRange(sheet.autoFilterRange, normalizedOperation) : undefined;
    const nextAutoFilterRange = shiftedAutoFilterRange ?? undefined;

    const nextCharts = (sheet.charts ?? []).flatMap((chart) => {
      const shiftedRange = shiftRange(chart.range, normalizedOperation);
      const shiftedAnchor = shiftCoord(chart.anchor, normalizedOperation);
      if (!shiftedRange || !shiftedAnchor) {
        return [];
      }

      return [{ ...cloneSheetChart(chart), range: shiftedRange, anchor: shiftedAnchor }];
    });

    const nextMerges = (sheet.merges ?? []).flatMap((merge) => {
      const shiftedRange = shiftRange(merge.range, normalizedOperation);
      return shiftedRange ? [{ ...cloneSheetMerge(merge), range: shiftedRange }] : [];
    });

    const rowCount =
      operation.axis === "row"
        ? operation.mode === "insert"
          ? sheet.rowCount + count
          : Math.max(DEFAULT_ROWS, sheet.rowCount - count)
        : sheet.rowCount;
    const columnCount =
      operation.axis === "column"
        ? operation.mode === "insert"
          ? sheet.columnCount + count
          : Math.max(DEFAULT_COLUMNS, sheet.columnCount - count)
        : sheet.columnCount;

    return {
      ...sheet,
      rowCount,
      columnCount,
      cells: nextCells,
      formats: nextFormats,
      columnWidths: nextColumnWidths,
      rowHeights: nextRowHeights,
      hiddenColumns: nextHiddenColumns,
      hiddenRows: nextHiddenRows,
      comments: nextComments,
      hyperlinks: nextHyperlinks,
      validations: nextValidations,
      conditionalFormats: nextConditionalFormats,
      autoFilterRange: nextAutoFilterRange,
      filters: nextFilters,
      charts: nextCharts,
      merges: nextMerges,
      protection: nextProtection
    };
  });

  const nextNamedRanges = (nextWorkbook.namedRanges ?? []).flatMap((namedRange) => {
    if (namedRange.sheetId !== sheetId) {
      return [cloneNamedRange(namedRange)];
    }

    const shiftedRange = shiftRange(namedRange.range, normalizedOperation);
    return shiftedRange ? [{ ...cloneNamedRange(namedRange), range: shiftedRange }] : [];
  });

  return { ...nextWorkbook, namedRanges: nextNamedRanges };
}

function shiftRange(range: CellRange, operation: StructureOperation): CellRange | null {
  const normalized = normalizeRange(range);
  const nextStart = shiftCoord(normalized.start, operation);
  const nextEnd = shiftCoord(normalized.end, operation);
  if (!nextStart || !nextEnd) {
    return null;
  }
  return normalizeRange({ start: nextStart, end: nextEnd });
}

function translateFillContent(
  content: CellContent,
  offset: { rowOffset: number; columnOffset: number }
): CellContent {
  return typeof content === "string" && content.startsWith("=") ? translateFormulaReferences(content, offset) : content;
}

function autoFillDown(
  workbook: WorkbookModel,
  sheetId: string,
  source: CellRange,
  target: CellRange
): WorkbookModel {
  let nextWorkbook = workbook;

  for (let column = source.start.column; column <= source.end.column; column += 1) {
    const sourceCells = collectAutoFillSource(workbook, sheetId, source.start.row, source.end.row, (row) => ({ row, column }));
    const namedListSeries = autoFillNamedListSeries(sourceCells.map((cell) => cell.content));
    const dateSeries = isoDateSeries(sourceCells.map((cell) => cell.content));
    const series = numericSeries(sourceCells.map((cell) => cell.content));
    for (let row = source.end.row + 1; row <= target.end.row; row += 1) {
      const sourceIndex = (row - source.start.row) % sourceCells.length;
      const sourceCell = sourceCells[sourceIndex];
      const content =
        namedListSeries !== null
          ? namedListSeries.values[wrapSeriesIndex(namedListSeries.start + namedListSeries.step * (row - source.start.row), namedListSeries.values.length)]
          : dateSeries !== null
          ? isoDateFromDayNumber(dateSeries.start + dateSeries.step * (row - source.start.row))
          : series === null
          ? translateFillContent(sourceCell.content, {
              rowOffset: row - sourceCell.coord.row,
              columnOffset: column - sourceCell.coord.column
            })
          : series.start + series.step * (row - source.start.row);
      const targetAddress = formatCellAddress({ row, column });
      nextWorkbook = setCellContent(nextWorkbook, sheetId, targetAddress, content);
      nextWorkbook = applyAutoFillMetadata(nextWorkbook, sheetId, targetAddress, sourceCell);
    }
  }

  return nextWorkbook;
}

function autoFillRight(
  workbook: WorkbookModel,
  sheetId: string,
  source: CellRange,
  target: CellRange
): WorkbookModel {
  let nextWorkbook = workbook;

  for (let row = source.start.row; row <= source.end.row; row += 1) {
    const sourceCells = collectAutoFillSource(workbook, sheetId, source.start.column, source.end.column, (column) => ({ row, column }));
    const namedListSeries = autoFillNamedListSeries(sourceCells.map((cell) => cell.content));
    const dateSeries = isoDateSeries(sourceCells.map((cell) => cell.content));
    const series = numericSeries(sourceCells.map((cell) => cell.content));
    for (let column = source.end.column + 1; column <= target.end.column; column += 1) {
      const sourceIndex = (column - source.start.column) % sourceCells.length;
      const sourceCell = sourceCells[sourceIndex];
      const content =
        namedListSeries !== null
          ? namedListSeries.values[
              wrapSeriesIndex(namedListSeries.start + namedListSeries.step * (column - source.start.column), namedListSeries.values.length)
            ]
          : dateSeries !== null
          ? isoDateFromDayNumber(dateSeries.start + dateSeries.step * (column - source.start.column))
          : series === null
          ? translateFillContent(sourceCell.content, {
              rowOffset: row - sourceCell.coord.row,
              columnOffset: column - sourceCell.coord.column
            })
          : series.start + series.step * (column - source.start.column);
      const targetAddress = formatCellAddress({ row, column });
      nextWorkbook = setCellContent(nextWorkbook, sheetId, targetAddress, content);
      nextWorkbook = applyAutoFillMetadata(nextWorkbook, sheetId, targetAddress, sourceCell);
    }
  }

  return nextWorkbook;
}

function collectAutoFillSource(
  workbook: WorkbookModel,
  sheetId: string,
  start: number,
  end: number,
  coordAt: (index: number) => { row: number; column: number }
): AutoFillSourceCell[] {
  return Array.from({ length: end - start + 1 }, (_, offset) => {
    const coord = coordAt(start + offset);
    return {
      coord,
      content: getCellContent(workbook, sheetId, formatCellAddress(coord)),
      format: getCellFormat(workbook, sheetId, formatCellAddress(coord)),
      validation: getCellValidation(workbook, sheetId, formatCellAddress(coord))
    };
  });
}

function applyAutoFillMetadata(
  workbook: WorkbookModel,
  sheetId: string,
  targetAddress: string,
  sourceCell: AutoFillSourceCell
): WorkbookModel {
  const coord = parseCellAddress(targetAddress);
  const targetRange = { start: coord, end: coord };
  const sourceFormat = compactCellFormat(sourceCell.format);
  let nextWorkbook = clearDirectCellFormats(workbook, sheetId, targetRange);

  if (!isEmptyFormat(sourceFormat)) {
    nextWorkbook = setCellFormat(nextWorkbook, sheetId, targetRange, sourceFormat);
  }

  return setCellValidation(nextWorkbook, sheetId, targetRange, sourceCell.validation);
}

function isDownAutoFill(source: CellRange, target: CellRange): boolean {
  return (
    target.start.row === source.start.row &&
    target.start.column === source.start.column &&
    target.end.column === source.end.column &&
    target.end.row > source.end.row
  );
}

function isRightAutoFill(source: CellRange, target: CellRange): boolean {
  return (
    target.start.row === source.start.row &&
    target.start.column === source.start.column &&
    target.end.row === source.end.row &&
    target.end.column > source.end.column
  );
}

function numericSeries(contents: CellContent[]): { start: number; step: number } | null {
  if (contents.length < 2) {
    return null;
  }

  const values = contents.map(cellContentToNumber);
  if (values.some((value) => value === null)) {
    return null;
  }

  const numbers = values as number[];
  const step = numbers[1] - numbers[0];
  if (!numbers.slice(1).every((value, index) => value - numbers[index] === step)) {
    return null;
  }

  return { start: numbers[0], step };
}

function autoFillNamedListSeries(contents: CellContent[]): AutoFillNamedListSeries | null {
  const values = contents.map((content) => (typeof content === "string" ? content.trim().toLowerCase() : null));
  if (values.some((value) => value === null || value === "")) {
    return null;
  }

  for (const list of AUTO_FILL_NAMED_LISTS) {
    const indices = values.map((value) => list.findIndex((item) => item.toLowerCase() === value));
    if (indices.some((index) => index < 0)) {
      continue;
    }

    if (indices.length === 1) {
      return { values: list, start: indices[0], step: 1 };
    }

    const step = wrapSeriesIndex(indices[1] - indices[0], list.length);
    const isConsistent = indices.slice(1).every((index, offset) => wrapSeriesIndex(index - indices[offset], list.length) === step);
    if (isConsistent) {
      return { values: list, start: indices[0], step };
    }
  }

  return null;
}

function isoDateSeries(contents: CellContent[]): { start: number; step: number } | null {
  const values = contents.map(cellContentToIsoDayNumber);
  if (values.some((value) => value === null)) {
    return null;
  }

  const dayNumbers = values as number[];
  if (dayNumbers.length === 1) {
    return { start: dayNumbers[0], step: 1 };
  }

  const step = dayNumbers[1] - dayNumbers[0];
  if (!dayNumbers.slice(1).every((value, index) => value - dayNumbers[index] === step)) {
    return null;
  }

  return { start: dayNumbers[0], step };
}

function cellContentToNumber(content: CellContent): number | null {
  if (typeof content === "number" && Number.isFinite(content)) {
    return content;
  }
  if (typeof content !== "string" || content.trim() === "") {
    return null;
  }

  const parsed = Number(content);
  return Number.isFinite(parsed) ? parsed : null;
}

function cellContentToIsoDayNumber(content: CellContent): number | null {
  if (typeof content !== "string") {
    return null;
  }

  const match = content.trim().match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!match) {
    return null;
  }

  const year = Number(match[1]);
  const month = Number(match[2]);
  const day = Number(match[3]);
  const timestamp = Date.UTC(year, month - 1, day);
  const date = new Date(timestamp);
  if (date.getUTCFullYear() !== year || date.getUTCMonth() !== month - 1 || date.getUTCDate() !== day) {
    return null;
  }

  return timestamp / MS_PER_DAY;
}

function isoDateFromDayNumber(dayNumber: number): string {
  return new Date(dayNumber * MS_PER_DAY).toISOString().slice(0, 10);
}

function wrapSeriesIndex(index: number, count: number): number {
  return ((index % count) + count) % count;
}

function shiftCoord(coord: { row: number; column: number }, operation: StructureOperation) {
  const end = operation.index + operation.count;
  if (operation.axis === "row") {
    if (operation.mode === "insert") {
      return { ...coord, row: coord.row >= operation.index ? coord.row + operation.count : coord.row };
    }
    if (coord.row >= operation.index && coord.row < end) {
      return null;
    }
    return { ...coord, row: coord.row >= end ? coord.row - operation.count : coord.row };
  }

  if (operation.mode === "insert") {
    return { ...coord, column: coord.column >= operation.index ? coord.column + operation.count : coord.column };
  }
  if (coord.column >= operation.index && coord.column < end) {
    return null;
  }
  return { ...coord, column: coord.column >= end ? coord.column - operation.count : coord.column };
}

function shiftFormulaReferences(content: CellContent, operation: StructureOperation): CellContent {
  if (typeof content !== "string" || !content.startsWith("=")) {
    return content;
  }

  return content.replace(/(\$?)([A-Z]+)(\$?)(\d+)/g, (_match, columnLock: string, columnName: string, rowLock: string, rowName: string) => {
    const nextCoord = shiftReferenceCoord(
      { row: Number(rowName) - 1, column: columnNameToIndex(columnName) },
      operation
    );
    if (!nextCoord) {
      return "#REF!";
    }
    return `${columnLock}${columnIndexToName(nextCoord.column)}${rowLock}${nextCoord.row + 1}`;
  });
}

function shiftReferenceCoord(coord: { row: number; column: number }, operation: StructureOperation) {
  const end = operation.index + operation.count;
  if (operation.axis === "row") {
    if (operation.mode === "insert") {
      return { ...coord, row: coord.row >= operation.index ? coord.row + operation.count : coord.row };
    }
    if (coord.row >= operation.index && coord.row < end) {
      return null;
    }
    return { ...coord, row: coord.row >= end ? coord.row - operation.count : coord.row };
  }

  if (operation.mode === "insert") {
    return { ...coord, column: coord.column >= operation.index ? coord.column + operation.count : coord.column };
  }
  if (coord.column >= operation.index && coord.column < end) {
    return null;
  }
  return { ...coord, column: coord.column >= end ? coord.column - operation.count : coord.column };
}

function clampStructureIndex(value: number, min: number, max: number): number {
  if (!Number.isFinite(value)) {
    return min;
  }
  return Math.min(Math.max(Math.floor(value), min), max);
}

function setIndexFlags(
  flags: Record<string, boolean>,
  startIndex: number,
  endIndex: number,
  isFlagged: boolean,
  limit: number
): Record<string, boolean> {
  if (limit <= 0) {
    return {};
  }

  const minIndex = Math.min(startIndex, endIndex);
  const maxIndex = Math.max(startIndex, endIndex);
  const start = clampStructureIndex(minIndex, 0, limit - 1);
  const end = clampStructureIndex(maxIndex, 0, limit - 1);
  const nextFlags = compactIndexFlags(flags);

  for (let index = start; index <= end; index += 1) {
    const key = String(index);
    if (isFlagged) {
      nextFlags[key] = true;
    } else {
      delete nextFlags[key];
    }
  }

  return nextFlags;
}

function shiftIndexedNumberMap(
  values: Record<string, number>,
  operation: StructureOperation,
  axis: StructureOperation["axis"],
  clampValue: (value: number) => number
): Record<string, number> {
  const nextValues: Record<string, number> = {};
  for (const [key, value] of Object.entries(values)) {
    const index = Number(key);
    if (!Number.isFinite(index)) {
      continue;
    }
    const nextIndex = operation.axis === axis ? shiftLinearIndex(index, operation) : index;
    if (nextIndex !== null) {
      nextValues[String(nextIndex)] = clampValue(value);
    }
  }
  return nextValues;
}

function shiftIndexFlags(
  flags: Record<string, boolean>,
  operation: StructureOperation,
  axis: StructureOperation["axis"]
): Record<string, boolean> {
  const nextFlags: Record<string, boolean> = {};
  for (const [key, isFlagged] of Object.entries(flags)) {
    const index = Number(key);
    if (!isFlagged || !Number.isFinite(index)) {
      continue;
    }
    const nextIndex = operation.axis === axis ? shiftLinearIndex(index, operation) : index;
    if (nextIndex !== null) {
      nextFlags[String(nextIndex)] = true;
    }
  }
  return nextFlags;
}

function shiftLinearIndex(index: number, operation: StructureOperation): number | null {
  const end = operation.index + operation.count;
  if (operation.mode === "insert") {
    return index >= operation.index ? index + operation.count : index;
  }
  if (index >= operation.index && index < end) {
    return null;
  }
  return index >= end ? index - operation.count : index;
}

function compactIndexFlags(flags: Record<string, boolean>): Record<string, boolean> {
  return Object.fromEntries(Object.entries(flags).filter(([, isFlagged]) => isFlagged === true)) as Record<string, boolean>;
}

function indexFlagsEqual(left: Record<string, boolean>, right: Record<string, boolean>): boolean {
  const compactLeft = compactIndexFlags(left);
  const compactRight = compactIndexFlags(right);
  const leftKeys = Object.keys(compactLeft);
  const rightKeys = Object.keys(compactRight);
  return leftKeys.length === rightKeys.length && leftKeys.every((key) => compactRight[key] === true);
}

function cellFormatsEqual(left: CellFormat, right: CellFormat): boolean {
  return (
    left.bold === right.bold &&
    left.italic === right.italic &&
    left.textColor === right.textColor &&
    left.backgroundColor === right.backgroundColor &&
    left.numberFormat === right.numberFormat &&
    left.horizontalAlign === right.horizontalAlign &&
    left.verticalAlign === right.verticalAlign &&
    left.wrapText === right.wrapText &&
    cellBordersEqual(left.borders, right.borders)
  );
}

function cellBordersEqual(left: CellBorders | undefined, right: CellBorders | undefined): boolean {
  return (
    borderSidesEqual(left?.top, right?.top) &&
    borderSidesEqual(left?.right, right?.right) &&
    borderSidesEqual(left?.bottom, right?.bottom) &&
    borderSidesEqual(left?.left, right?.left)
  );
}

function borderSidesEqual(left: CellBorderSide | undefined, right: CellBorderSide | undefined): boolean {
  return left?.style === right?.style && left?.color === right?.color;
}

function withBorders(
  format: CellFormat,
  range: CellRange,
  coord: { row: number; column: number },
  preset: Exclude<BorderPreset, "none">,
  border: CellBorderSide
): CellFormat {
  const borders: CellBorders = { ...(format.borders ?? {}) };
  const allSides = preset === "all";

  if (allSides || preset === "top" || (preset === "outer" && coord.row === range.start.row)) {
    borders.top = { ...border };
  }
  if (allSides || preset === "right" || (preset === "outer" && coord.column === range.end.column)) {
    borders.right = { ...border };
  }
  if (allSides || preset === "bottom" || (preset === "outer" && coord.row === range.end.row)) {
    borders.bottom = { ...border };
  }
  if (allSides || preset === "left" || (preset === "outer" && coord.column === range.start.column)) {
    borders.left = { ...border };
  }

  return compactCellFormat({ ...format, borders });
}

function withoutBorders(format: CellFormat): CellFormat {
  return compactCellFormat({ ...format, borders: undefined });
}

function compactCellFormat(format: CellFormat): CellFormat {
  const nextFormat = cloneCellFormat(format);
  if (!nextFormat.borders || Object.values(nextFormat.borders).every((side) => side === undefined)) {
    delete nextFormat.borders;
  }
  return nextFormat;
}

function isEmptyFormat(format: CellFormat): boolean {
  return (
    format.bold === undefined &&
    format.italic === undefined &&
    format.textColor === undefined &&
    format.backgroundColor === undefined &&
    format.numberFormat === undefined &&
    format.horizontalAlign === undefined &&
    format.verticalAlign === undefined &&
    format.wrapText === undefined &&
    format.borders === undefined
  );
}

function cloneCellFormats(formats: Record<string, CellFormat>): Record<string, CellFormat> {
  return Object.fromEntries(Object.entries(formats).map(([address, format]) => [address, cloneCellFormat(format)]));
}

function cloneCellFormat(format: CellFormat): CellFormat {
  return {
    ...format,
    borders: format.borders ? cloneCellBorders(format.borders) : undefined
  };
}

function cloneCellBorders(borders: CellBorders): CellBorders {
  return {
    top: borders.top ? { ...borders.top } : undefined,
    right: borders.right ? { ...borders.right } : undefined,
    bottom: borders.bottom ? { ...borders.bottom } : undefined,
    left: borders.left ? { ...borders.left } : undefined
  };
}

function dataValidationsEqual(left: DataValidationRule | undefined, right: DataValidationRule): boolean {
  if (!left || left.type !== right.type || left.allowBlank !== right.allowBlank) {
    return false;
  }

  if (left.type === "list" && right.type === "list") {
    return left.values.length === right.values.length && left.values.every((value, index) => value === right.values[index]);
  }

  if (left.type === "number" && right.type === "number") {
    return left.min === right.min && left.max === right.max;
  }

  if (left.type === "textLength" && right.type === "textLength") {
    return left.min === right.min && left.max === right.max;
  }

  return false;
}

function validationRulesEqual(left: DataValidationRule | undefined, right: DataValidationRule | null): boolean {
  if (!left && !right) {
    return true;
  }
  if (!right) {
    return false;
  }
  return dataValidationsEqual(left, right);
}

function cloneValidationRule(rule: DataValidationRule): DataValidationRule {
  return rule.type === "list" ? { ...rule, values: [...rule.values] } : { ...rule };
}

function cloneValidations(validations: Record<string, DataValidationRule>): Record<string, DataValidationRule> {
  return Object.fromEntries(Object.entries(validations).map(([address, rule]) => [address, cloneValidationRule(rule)]));
}

function createDefaultProtection(): SheetProtection {
  return { isProtected: false, lockedCells: {}, unlockedCells: {} };
}

function getSheetProtection(sheet: SheetModel): SheetProtection {
  return cloneSheetProtection(sheet.protection ?? createDefaultProtection());
}

function cloneSheetProtection(protection: SheetProtection): SheetProtection {
  return {
    isProtected: Boolean(protection.isProtected),
    lockedCells: { ...(protection.lockedCells ?? {}) },
    unlockedCells: { ...(protection.unlockedCells ?? {}) }
  };
}

function shiftAddressFlags(flags: Record<string, boolean>, operation: StructureOperation): Record<string, boolean> {
  const nextFlags: Record<string, boolean> = {};
  for (const [address, isFlagged] of Object.entries(flags)) {
    if (!isFlagged) {
      continue;
    }
    const nextCoord = shiftCoord(parseCellAddress(address), operation);
    if (nextCoord) {
      nextFlags[formatCellAddress(nextCoord)] = true;
    }
  }
  return nextFlags;
}

function cloneConditionalFormatRule(rule: ConditionalFormatRule): ConditionalFormatRule {
  return {
    ...rule,
    range: {
      start: { ...rule.range.start },
      end: { ...rule.range.end }
    },
    condition: { ...rule.condition },
    format: { ...rule.format }
  };
}

function cloneSheetFilter(filter: SheetFilter): SheetFilter {
  return {
    ...filter,
    range: {
      start: { ...filter.range.start },
      end: { ...filter.range.end }
    }
  };
}

function cloneSheetChart(chart: SheetChart): SheetChart {
  return {
    ...chart,
    range: {
      start: { ...chart.range.start },
      end: { ...chart.range.end }
    },
    anchor: { ...chart.anchor }
  };
}

function cloneSheetMerge(merge: SheetMerge): SheetMerge {
  return {
    ...merge,
    range: cloneRange(merge.range)
  };
}

function cloneNamedRange(namedRange: NamedRange): NamedRange {
  return {
    ...namedRange,
    range: cloneRange(namedRange.range)
  };
}

function cloneRange(range: CellRange): CellRange {
  return {
    start: { ...range.start },
    end: { ...range.end }
  };
}

function nextConditionalFormatId(rules: readonly ConditionalFormatRule[]): string {
  let index = rules.length + 1;
  let id = `conditional-format-${index}`;
  const existingIds = new Set(rules.map((rule) => rule.id));
  while (existingIds.has(id)) {
    index += 1;
    id = `conditional-format-${index}`;
  }
  return id;
}

function nextSheetFilterId(filters: readonly SheetFilter[]): string {
  let index = filters.length + 1;
  let id = `filter-${index}`;
  const existingIds = new Set(filters.map((filter) => filter.id));
  while (existingIds.has(id)) {
    index += 1;
    id = `filter-${index}`;
  }
  return id;
}

function nextSheetChartId(charts: readonly SheetChart[]): string {
  let index = charts.length + 1;
  let id = `chart-${index}`;
  const existingIds = new Set(charts.map((chart) => chart.id));
  while (existingIds.has(id)) {
    index += 1;
    id = `chart-${index}`;
  }
  return id;
}

function nextSheetMergeId(merges: readonly SheetMerge[]): string {
  let index = merges.length + 1;
  let id = `merge-${index}`;
  const existingIds = new Set(merges.map((merge) => merge.id));
  while (existingIds.has(id)) {
    index += 1;
    id = `merge-${index}`;
  }
  return id;
}

function isCoordInRange(coord: { row: number; column: number }, range: CellRange): boolean {
  return coord.row >= range.start.row && coord.row <= range.end.row && coord.column >= range.start.column && coord.column <= range.end.column;
}

function rangesIntersect(left: CellRange, right: CellRange): boolean {
  return (
    left.start.row <= right.end.row &&
    left.end.row >= right.start.row &&
    left.start.column <= right.end.column &&
    left.end.column >= right.start.column
  );
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

function sortBodyRange(sheet: SheetModel, range: CellRange): CellRange {
  const normalized = normalizeRange(range);
  if (normalized.start.row === normalized.end.row || !rangeHasAutoFilterHeader(sheet, normalized)) {
    return normalized;
  }

  return {
    start: { row: normalized.start.row + 1, column: normalized.start.column },
    end: { ...normalized.end }
  };
}

function rangeHasAutoFilterHeader(sheet: SheetModel, range: CellRange): boolean {
  const matchingFilters = (sheet.filters ?? []).filter((filter) => rangesEqual(filter.range, range));
  if (matchingFilters.length > 0) {
    return matchingFilters.some((filter) => filter.hasHeader !== false);
  }

  return Boolean(sheet.autoFilterRange && rangesEqual(sheet.autoFilterRange, range));
}

function normalizeNamedRangeLookup(name: string): string {
  return name.trim().toLocaleLowerCase();
}

function collectRangeRows(sheet: SheetModel, range: CellRange): RangeRowSnapshot[] {
  const normalized = normalizeRange(range);
  const columnCount = normalized.end.column - normalized.start.column + 1;

  return Array.from({ length: normalized.end.row - normalized.start.row + 1 }, (_, rowOffset) => {
    const row = normalized.start.row + rowOffset;
    const cells = Array.from({ length: columnCount }, (_, columnOffset) => {
      const column = normalized.start.column + columnOffset;
      return sheet.cells[formatCellAddress({ row, column })] ?? null;
    });

    return {
      originalIndex: rowOffset,
      key: duplicateRowKey(cells),
      cells,
      formats: Array.from({ length: columnCount }, (_, columnOffset) => {
        const column = normalized.start.column + columnOffset;
        return cloneCellFormat((sheet.formats ?? {})[formatCellAddress({ row, column })] ?? {});
      }),
      validations: Array.from({ length: columnCount }, (_, columnOffset) => {
        const column = normalized.start.column + columnOffset;
        const rule = (sheet.validations ?? {})[formatCellAddress({ row, column })];
        return rule ? cloneValidationRule(rule) : null;
      }),
      comments: Array.from({ length: columnCount }, (_, columnOffset) => {
        const column = normalized.start.column + columnOffset;
        return (sheet.comments ?? {})[formatCellAddress({ row, column })] ?? null;
      }),
      hyperlinks: Array.from({ length: columnCount }, (_, columnOffset) => {
        const column = normalized.start.column + columnOffset;
        return (sheet.hyperlinks ?? {})[formatCellAddress({ row, column })] ?? null;
      })
    };
  });
}

function duplicateRowKey(cells: CellContent[]): string {
  return JSON.stringify(cells.map((cell) => [cell === null ? "blank" : typeof cell, cell]));
}

function compareSortRows(
  left: { key: CellContent; originalIndex: number },
  right: { key: CellContent; originalIndex: number },
  direction: "asc" | "desc"
): number {
  const leftBlank = left.key === null || left.key === "";
  const rightBlank = right.key === null || right.key === "";
  if (leftBlank && rightBlank) {
    return left.originalIndex - right.originalIndex;
  }
  if (leftBlank) {
    return 1;
  }
  if (rightBlank) {
    return -1;
  }

  const leftNumber = Number(left.key);
  const rightNumber = Number(right.key);
  const comparison =
    Number.isFinite(leftNumber) && Number.isFinite(rightNumber)
      ? leftNumber - rightNumber
      : String(left.key).localeCompare(String(right.key), undefined, { numeric: true, sensitivity: "base" });

  if (comparison === 0) {
    return left.originalIndex - right.originalIndex;
  }
  return direction === "asc" ? comparison : comparison * -1;
}

function normalizeAddress(address: string): string {
  return formatCellAddress(parseCellAddress(address));
}

function normalizeHexColor(color: string | undefined): string | undefined {
  const value = color?.trim();
  if (!value) {
    return undefined;
  }

  const normalized = value.startsWith("#") ? value : `#${value}`;
  return HEX_COLOR_PATTERN.test(normalized) ? normalized.toLowerCase() : undefined;
}

function uniqueFilterValues(values: readonly string[]): string[] {
  return [...new Set(values.map((value) => String(value)))];
}

function nextSheetId(workbook: WorkbookModel): string {
  let index = workbook.sheets.length + 1;
  let id = `sheet-${index}`;
  while (workbook.sheets.some((sheet) => sheet.id === id)) {
    index += 1;
    id = `sheet-${index}`;
  }
  return id;
}

function uniqueSheetName(workbook: WorkbookModel, requestedName: string, ignoreSheetId?: string): string {
  const existingNames = new Set(
    workbook.sheets.filter((sheet) => sheet.id !== ignoreSheetId).map((sheet) => sheet.name)
  );
  if (!existingNames.has(requestedName)) {
    return requestedName;
  }

  let index = 2;
  let name = `${requestedName} ${index}`;
  while (existingNames.has(name)) {
    index += 1;
    name = `${requestedName} ${index}`;
  }

  return name;
}
