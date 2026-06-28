import type { CellRange, SheetProtection, WorkbookModel } from "../types";
import { createBlankWorkbook } from "./workbook";

export const WORKBOOK_STORAGE_KEY = "javascript-spreadsheet-workbook";
const HEX_COLOR_PATTERN = /^#[0-9a-f]{6}$/i;

export function saveWorkbook(storage: Storage, workbook: WorkbookModel): void {
  storage.setItem(WORKBOOK_STORAGE_KEY, JSON.stringify(workbook));
}

export function loadWorkbook(storage: Storage): WorkbookModel {
  const serialized = storage.getItem(WORKBOOK_STORAGE_KEY);
  if (!serialized) {
    return createBlankWorkbook();
  }

  try {
    const parsed: unknown = JSON.parse(serialized);
    return isWorkbookModel(parsed) ? migrateWorkbook(parsed) : createBlankWorkbook();
  } catch {
    return createBlankWorkbook();
  }
}

function isWorkbookModel(value: unknown): value is WorkbookModel {
  if (!isRecord(value) || value.version !== 1 || typeof value.activeSheetId !== "string") {
    return false;
  }

  if (!Array.isArray(value.sheets) || value.sheets.length === 0) {
    return false;
  }

  return value.sheets.every(
    (sheet) =>
      isRecord(sheet) &&
      typeof sheet.id === "string" &&
      typeof sheet.name === "string" &&
      typeof sheet.rowCount === "number" &&
      typeof sheet.columnCount === "number" &&
      isRecord(sheet.cells) &&
      (sheet.formats === undefined || isRecord(sheet.formats)) &&
      (sheet.columnWidths === undefined || isRecord(sheet.columnWidths)) &&
      (sheet.rowHeights === undefined || isRecord(sheet.rowHeights)) &&
      (sheet.isHidden === undefined || typeof sheet.isHidden === "boolean") &&
      (sheet.tabColor === undefined || typeof sheet.tabColor === "string") &&
      (sheet.hiddenColumns === undefined || isRecord(sheet.hiddenColumns)) &&
      (sheet.hiddenRows === undefined || isRecord(sheet.hiddenRows)) &&
      (sheet.freezeTopRow === undefined || typeof sheet.freezeTopRow === "boolean") &&
      (sheet.freezeFirstColumn === undefined || typeof sheet.freezeFirstColumn === "boolean") &&
      (sheet.autoFilterRange === undefined || isCellRange(sheet.autoFilterRange)) &&
      (sheet.comments === undefined || isRecord(sheet.comments)) &&
      (sheet.hyperlinks === undefined || isRecord(sheet.hyperlinks)) &&
      (sheet.validations === undefined || isRecord(sheet.validations)) &&
      (sheet.conditionalFormats === undefined || Array.isArray(sheet.conditionalFormats)) &&
      (sheet.filters === undefined || Array.isArray(sheet.filters)) &&
      (sheet.charts === undefined || Array.isArray(sheet.charts)) &&
      (sheet.merges === undefined || Array.isArray(sheet.merges)) &&
      (sheet.protection === undefined || isRecord(sheet.protection))
  ) && (value.namedRanges === undefined || Array.isArray(value.namedRanges));
}

function migrateWorkbook(workbook: WorkbookModel): WorkbookModel {
  const migratedSheets = workbook.sheets.map((sheet) => ({
    ...sheet,
    isHidden: sheet.isHidden === true,
    tabColor: normalizeHexColor(sheet.tabColor),
    formats: sheet.formats ?? {},
    columnWidths: sheet.columnWidths ?? {},
    rowHeights: sheet.rowHeights ?? {},
    hiddenColumns: isRecord(sheet.hiddenColumns) ? booleanFlagRecord(sheet.hiddenColumns) : {},
    hiddenRows: isRecord(sheet.hiddenRows) ? booleanFlagRecord(sheet.hiddenRows) : {},
    freezeTopRow: sheet.freezeTopRow === true,
    freezeFirstColumn: sheet.freezeFirstColumn === true,
    comments: sheet.comments ?? {},
    hyperlinks: sheet.hyperlinks ?? {},
    validations: sheet.validations ?? {},
    conditionalFormats: sheet.conditionalFormats ?? [],
    autoFilterRange: migratedAutoFilterRange(sheet),
    filters: sheet.filters ?? [],
    charts: sheet.charts ?? [],
    merges: sheet.merges ?? [],
    protection: migrateProtection(sheet.protection)
  }));
  const sheets = migratedSheets.some((sheet) => sheet.isHidden !== true)
    ? migratedSheets
    : migratedSheets.map((sheet, index) => (index === 0 ? { ...sheet, isHidden: false } : sheet));
  const activeSheet = sheets.find((sheet) => sheet.id === workbook.activeSheetId && sheet.isHidden !== true) ?? sheets.find((sheet) => sheet.isHidden !== true) ?? sheets[0];

  return {
    ...workbook,
    activeSheetId: activeSheet.id,
    namedRanges: workbook.namedRanges ?? [],
    sheets
  };
}

function migratedAutoFilterRange(sheet: WorkbookModel["sheets"][number]): CellRange | undefined {
  if (isCellRange(sheet.autoFilterRange)) {
    return cloneRange(sheet.autoFilterRange);
  }

  if (!Array.isArray(sheet.filters)) {
    return undefined;
  }

  const filterWithRange = sheet.filters.find((filter) => isRecord(filter) && isCellRange(filter.range));
  return filterWithRange && isRecord(filterWithRange) && isCellRange(filterWithRange.range)
    ? cloneRange(filterWithRange.range)
    : undefined;
}

function migrateProtection(value: unknown): SheetProtection {
  if (!isRecord(value)) {
    return { isProtected: false, lockedCells: {}, unlockedCells: {} };
  }

  return {
    isProtected: value.isProtected === true,
    lockedCells: isRecord(value.lockedCells) ? booleanFlagRecord(value.lockedCells) : {},
    unlockedCells: isRecord(value.unlockedCells) ? booleanFlagRecord(value.unlockedCells) : {}
  };
}

function isCellRange(value: unknown): value is CellRange {
  return (
    isRecord(value) &&
    isCellCoord(value.start) &&
    isCellCoord(value.end)
  );
}

function isCellCoord(value: unknown): value is CellRange["start"] {
  return isRecord(value) && Number.isInteger(value.row) && Number.isInteger(value.column);
}

function cloneRange(range: CellRange): CellRange {
  return {
    start: { ...range.start },
    end: { ...range.end }
  };
}

function booleanFlagRecord(value: Record<string, unknown>): Record<string, boolean> {
  return Object.fromEntries(Object.entries(value).filter(([, isFlagged]) => isFlagged === true)) as Record<string, boolean>;
}

function normalizeHexColor(color: string | undefined): string | undefined {
  const value = color?.trim();
  if (!value) {
    return undefined;
  }

  const normalized = value.startsWith("#") ? value : `#${value}`;
  return HEX_COLOR_PATTERN.test(normalized) ? normalized.toLowerCase() : undefined;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}
