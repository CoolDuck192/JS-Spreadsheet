/**
 * Library entry point so the spreadsheet can be embedded in a host app:
 *
 *   import Spreadsheet, { createBlankWorkbook, importWorkbookFromXlsx } from "javascript-spreadsheet-clone";
 *
 * The default export is the full spreadsheet UI; the named exports are the
 * headless workbook/engine/io functions for programmatic use.
 */

export { default } from "./App";
export { default as Spreadsheet } from "./App";

export type {
  CellContent,
  CellCoord,
  CellRange,
  CellFormat,
  ConditionalFormatRule,
  DataValidationRule,
  NamedRange,
  SheetChart,
  SheetFilter,
  SheetMerge,
  SheetModel,
  WorkbookModel
} from "./types";

export { createBlankWorkbook, getActiveSheet, getCellContent, setCellContent } from "./lib/workbook";
export { createFormulaEngine, type FormulaEngine } from "./lib/formulaEngine";
export { loadWorkbook, saveWorkbook, WORKBOOK_STORAGE_KEY } from "./lib/persistence";
export { exportWorkbookToXlsx, importWorkbookFromXlsx } from "./lib/xlsx";
export { createPivotTable } from "./lib/pivot";
export {
  importWorkbookFromGoogleSheets,
  parseSpreadsheetId,
  type GoogleSheetsImportResult
} from "./lib/googleSheets";
export {
  createBrowserTokenProvider,
  SHEETS_READONLY_SCOPE,
  SHEETS_READWRITE_SCOPE,
  type TokenProvider
} from "./lib/googleAuth";
