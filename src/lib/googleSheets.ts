/**
 * Google Sheets connector: link a user's Google Sheet and import it as a
 * WorkbookModel, preserving formulas. Read path uses the Sheets API v4
 * spreadsheets.get + values.batchGet with valueRenderOption=FORMULA so
 * formulas arrive as "=..." strings and flow through the same engine as
 * typed input, mirroring the xlsx import path.
 */

import type { CellContent, SheetModel, WorkbookModel } from "../types";
import { SHEETS_READONLY_SCOPE, type TokenProvider } from "./googleAuth";

const SHEETS_API_BASE = "https://sheets.googleapis.com/v4/spreadsheets";
const DEFAULT_ROW_COUNT = 100;
const DEFAULT_COLUMN_COUNT = 26;

export type GoogleSheetsImportResult = {
  workbook: WorkbookModel;
  spreadsheetTitle: string;
};

/** Extract the spreadsheet id from a full URL or accept a bare id. */
export function parseSpreadsheetId(input: string): string | null {
  const trimmed = input.trim();
  if (!trimmed) {
    return null;
  }

  const urlMatch = trimmed.match(/\/spreadsheets\/d\/([A-Za-z0-9_-]+)/);
  if (urlMatch) {
    return urlMatch[1];
  }

  return /^[A-Za-z0-9_-]{20,}$/.test(trimmed) ? trimmed : null;
}

export async function importWorkbookFromGoogleSheets(
  spreadsheetIdOrUrl: string,
  tokenProvider: TokenProvider,
  fetchImpl: typeof fetch = fetch
): Promise<GoogleSheetsImportResult> {
  const spreadsheetId = parseSpreadsheetId(spreadsheetIdOrUrl);
  if (!spreadsheetId) {
    throw new Error("Not a Google Sheets URL or spreadsheet id");
  }

  const token = await tokenProvider.getAccessToken([SHEETS_READONLY_SCOPE]);
  const authHeaders = { Authorization: `Bearer ${token}` };

  const metadata = await fetchJson<{
    properties?: { title?: string };
    sheets?: Array<{ properties?: { title?: string; hidden?: boolean } }>;
  }>(`${SHEETS_API_BASE}/${spreadsheetId}?fields=properties.title,sheets.properties(title,hidden)`, authHeaders, fetchImpl);

  const sheetTitles = (metadata.sheets ?? [])
    .map((sheet) => sheet.properties?.title)
    .filter((title): title is string => Boolean(title));
  if (sheetTitles.length === 0) {
    throw new Error("The Google Sheet has no sheets");
  }

  const rangesQuery = sheetTitles.map((title) => `ranges=${encodeURIComponent(quoteA1SheetName(title))}`).join("&");
  const valuesResponse = await fetchJson<{
    valueRanges?: Array<{ values?: unknown[][] }>;
  }>(
    `${SHEETS_API_BASE}/${spreadsheetId}/values:batchGet?${rangesQuery}&valueRenderOption=FORMULA&dateTimeRenderOption=FORMATTED_STRING`,
    authHeaders,
    fetchImpl
  );

  const hiddenByTitle = new Map(
    (metadata.sheets ?? []).map((sheet) => [sheet.properties?.title ?? "", sheet.properties?.hidden === true])
  );

  const sheets: SheetModel[] = sheetTitles.map((title, index) => {
    const values = valuesResponse.valueRanges?.[index]?.values ?? [];
    return valuesToSheetModel(`gsheet-${index + 1}`, title, values, hiddenByTitle.get(title) === true);
  });

  const activeSheet = sheets.find((sheet) => sheet.isHidden !== true) ?? sheets[0];

  return {
    spreadsheetTitle: metadata.properties?.title ?? "Google Sheet",
    workbook: {
      version: 2,
      activeSheetId: activeSheet.id,
      sheets,
      namedRanges: [],
      tables: []
    }
  };
}

async function fetchJson<T>(url: string, headers: Record<string, string>, fetchImpl: typeof fetch): Promise<T> {
  const response = await fetchImpl(url, { headers });
  if (!response.ok) {
    if (response.status === 401 || response.status === 403) {
      throw new Error("Google Sheets access denied — sign in again or check the sheet's sharing settings");
    }
    if (response.status === 404) {
      throw new Error("Google Sheet not found — check the URL");
    }
    if (response.status === 429) {
      throw new Error("Google Sheets rate limit reached — try again in a minute");
    }
    throw new Error(`Google Sheets request failed (${response.status})`);
  }
  return (await response.json()) as T;
}

function valuesToSheetModel(id: string, name: string, values: unknown[][], isHidden: boolean): SheetModel {
  const cells: Record<string, CellContent> = {};
  let maxColumn = 0;

  values.forEach((row, rowIndex) => {
    row.forEach((value, columnIndex) => {
      const content = toCellContent(value);
      if (content !== null) {
        cells[`${columnLabel(columnIndex)}${rowIndex + 1}`] = content;
        maxColumn = Math.max(maxColumn, columnIndex + 1);
      }
    });
  });

  return {
    id,
    name,
    rowCount: Math.max(DEFAULT_ROW_COUNT, values.length),
    columnCount: Math.max(DEFAULT_COLUMN_COUNT, maxColumn),
    cells,
    formats: {},
    columnWidths: {},
    rowHeights: {},
    isHidden,
    hiddenColumns: {},
    hiddenRows: {},
    freezeTopRow: false,
    freezeFirstColumn: false,
    comments: {},
    hyperlinks: {},
    validations: {},
    conditionalFormats: [],
    filters: [],
    charts: [],
    merges: [],
    protection: { isProtected: false, lockedCells: {}, unlockedCells: {} }
  };
}

function toCellContent(value: unknown): CellContent {
  if (value === null || value === undefined) {
    return null;
  }
  if (typeof value === "number" || typeof value === "boolean") {
    return value;
  }
  const text = String(value);
  return text === "" ? null : text;
}

function columnLabel(index: number): string {
  let remaining = index + 1;
  let label = "";
  while (remaining > 0) {
    const remainder = (remaining - 1) % 26;
    label = String.fromCharCode(65 + remainder) + label;
    remaining = Math.floor((remaining - 1) / 26);
  }
  return label;
}

function quoteA1SheetName(title: string): string {
  return `'${title.replace(/'/g, "''")}'`;
}
