/**
 * Google Sheets connector: import a user's Google Sheet one time as a
 * WorkbookModel, preserving formulas. The read path uses the Sheets API v4
 * spreadsheets.get + values.batchGet with valueRenderOption=FORMULA so
 * formulas arrive as "=..." strings and flow through the same engine as
 * typed input, mirroring the xlsx import path.
 */

import type { CellContent, SheetModel, WorkbookModel } from "../types";
import { SHEETS_READONLY_SCOPE, type TokenProvider } from "./googleAuth";
import { GoogleSheetsError, toGoogleSheetsError } from "./googleErrors";

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
    throw new GoogleSheetsError(
      "invalid_sheet_url",
      "Enter a valid Google Sheets URL or spreadsheet ID.",
      true
    );
  }

  let token: string;
  try {
    token = await tokenProvider.getAccessToken([SHEETS_READONLY_SCOPE]);
  } catch (error) {
    throw toGoogleSheetsError(error);
  }
  const authHeaders = { Authorization: `Bearer ${token}` };

  const metadata = await fetchJson<{
    properties?: { title?: string };
    sheets?: Array<{ properties?: { title?: string; hidden?: boolean } }>;
  }>(`${SHEETS_API_BASE}/${spreadsheetId}?fields=properties.title,sheets.properties(title,hidden)`, authHeaders, fetchImpl);

  const sheetTitles = (metadata.sheets ?? [])
    .map((sheet) => sheet.properties?.title)
    .filter((title): title is string => Boolean(title));
  if (sheetTitles.length === 0) {
    throw new GoogleSheetsError(
      "sheet_not_found",
      "The Google Sheet does not contain an importable worksheet.",
      true
    );
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
  let response: Response;
  try {
    response = await fetchImpl(url, { headers });
  } catch {
    throw new GoogleSheetsError(
      "network_failed",
      "Google Sheets could not be reached. Check your connection and try again.",
      true
    );
  }
  if (!response.ok) {
    if (response.status === 403 && (await isSheetsApiDisabled(response))) {
      throw new GoogleSheetsError(
        "api_not_enabled",
        "The Google Sheets API is not enabled for this Google OAuth project.",
        true
      );
    }
    if (response.status === 401 || response.status === 403) {
      throw new GoogleSheetsError(
        "access_denied",
        "Google Sheets access was denied. Sign in again or check the sheet sharing settings.",
        true
      );
    }
    if (response.status === 404) {
      throw new GoogleSheetsError(
        "sheet_not_found",
        "Google Sheet not found. Check the URL and sharing settings.",
        true
      );
    }
    if (response.status === 429) {
      throw new GoogleSheetsError(
        "rate_limited",
        "Google Sheets is receiving too many requests. Wait a moment and try again.",
        true
      );
    }
    throw new GoogleSheetsError(
      "unknown",
      "Google Sheets import failed. Check your connection and try again.",
      true
    );
  }
  try {
    return (await response.json()) as T;
  } catch {
    throw new GoogleSheetsError(
      "network_failed",
      "Google Sheets returned an unreadable response. Try again.",
      true
    );
  }
}

const API_DISABLED_STATUSES = new Set(["SERVICE_DISABLED", "API_NOT_ENABLED"]);
const API_DISABLED_REASONS = new Set([
  "accessNotConfigured",
  "apiNotEnabled",
  "serviceDisabled"
]);

async function isSheetsApiDisabled(response: Response): Promise<boolean> {
  let body: unknown;
  try {
    body = await response.json();
  } catch {
    return false;
  }
  if (!isRecord(body) || !isRecord(body.error)) {
    return false;
  }

  const status = body.error.status;
  if (typeof status === "string" && API_DISABLED_STATUSES.has(status)) {
    return true;
  }
  const details = body.error.details;
  if (Array.isArray(details) && details.some(
    (entry) =>
      isRecord(entry) &&
      typeof entry.reason === "string" &&
      API_DISABLED_STATUSES.has(entry.reason)
  )) {
    return true;
  }
  const errors = body.error.errors;
  if (!Array.isArray(errors)) {
    return false;
  }
  return errors.some(
    (entry) =>
      isRecord(entry) &&
      typeof entry.reason === "string" &&
      API_DISABLED_REASONS.has(entry.reason)
  );
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
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
