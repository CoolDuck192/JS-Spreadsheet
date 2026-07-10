/**
 * Library entry point so the spreadsheet can be embedded in a host app:
 *
 *   import Spreadsheet, { createBlankWorkbook, importWorkbookFromXlsx } from "js-spreadsheet";
 *
 * The default export is the full spreadsheet UI; the named exports are the
 * headless workbook/engine/io functions for programmatic use.
 */

export { default, Spreadsheet } from "./react/Spreadsheet";
export type {
  SpreadsheetCommonProps,
  SpreadsheetErrorEvent,
  SpreadsheetProps,
  WorkbookChangeEvent,
  WorkbookFeatureConfiguration,
  WorkbookStorage,
  WorkbookThemeToken
} from "./react/Spreadsheet";
export { useWorkbookSession } from "./react/useWorkbookSession";
export type {
  UseWorkbookSessionCommonOptions,
  UseWorkbookSessionOptions
} from "./react/useWorkbookSession";
export {
  WorkbookSessionProvider,
  useWorkbookSessionContext
} from "./react/WorkbookSessionContext";
export { DataTable } from "./react/DataTable";
export { useTableSession } from "./react/useTableSession";
export { useTableSnapshot } from "./react/useTableSnapshot";
export {
  TableSessionProvider,
  useTableSessionContext
} from "./react/TableSessionContext";
export type {
  CellEditorProps,
  CellRenderContext,
  ColumnDef,
  DataTableHandle,
  DataTablePresentationProps,
  DataTableProps,
  HeaderRenderContext
} from "./react/tableTypes";
export {
  RecordTableSession,
  createLocalRecordTableSession
} from "./table/local/RecordTableSession";
export type {
  LocalRecordSource,
  LocalRecordTableSessionOptions,
  RecordFormulaService
} from "./table/local/RecordTableSession";
export type {
  ExportArtifact,
  ExportOptions,
  TableIntent,
  TableSelection,
  TableSession,
  TableViewSnapshot,
  TableViewState
} from "./table/core/types";
export {
  createWorkbookSession,
  type CreateWorkbookSessionOptions,
  type WorkbookCommandResult,
  type WorkbookDiagnosticEvent,
  type WorkbookSession,
  type WorkbookSnapshot
} from "./core/workbook/WorkbookSession";
export type {
  SpreadsheetServices,
  TokenProvider as WorkbookTokenProvider,
  WorkbookExportArtifact,
  WorkbookExporter,
  WorkbookImporter,
  WorkbookImportPayload,
  WorkbookServiceOptions
} from "./core/workbook/services";

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
