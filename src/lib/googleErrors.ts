export type GoogleSheetsErrorCode =
  | "configuration_missing"
  | "incompatible_origin"
  | "origin_mismatch"
  | "invalid_client"
  | "invalid_sheet_url"
  | "gis_load_failed"
  | "popup_blocked"
  | "popup_closed"
  | "access_denied"
  | "api_not_enabled"
  | "sheet_not_found"
  | "rate_limited"
  | "network_failed"
  | "unknown";

const SAFE_GOOGLE_SHEETS_ERROR_MESSAGES: Readonly<Record<GoogleSheetsErrorCode, string>> = {
  configuration_missing:
    "Google Sheets import is not configured. Add a public Google OAuth client ID or provide a token provider.",
  incompatible_origin:
    "Google browser OAuth requires an eligible, authorized HTTPS origin.",
  origin_mismatch:
    "This site origin is not authorized for the Google OAuth client.",
  invalid_client:
    "Enter a valid Google OAuth Web application client ID.",
  invalid_sheet_url:
    "Enter a valid Google Sheets URL or spreadsheet ID.",
  gis_load_failed:
    "Google sign-in could not be loaded. Check your connection and try again.",
  popup_blocked:
    "Google sign-in was blocked by the browser. Allow pop-ups and try again.",
  popup_closed:
    "Google sign-in was closed before access was granted.",
  access_denied:
    "Google Sheets access was denied. Sign in again or check the sheet sharing settings.",
  api_not_enabled:
    "The Google Sheets API is not enabled for this Google Cloud project.",
  sheet_not_found:
    "The Google Sheet could not be found or is not shared with this account.",
  rate_limited:
    "Google Sheets is temporarily limiting requests. Wait a moment and try again.",
  network_failed:
    "Google Sheets could not be reached. Check your connection and try again.",
  unknown:
    "Google Sheets import failed. Check your connection and try again."
};

export class GoogleSheetsError extends Error {
  constructor(
    readonly code: GoogleSheetsErrorCode,
    message: string,
    readonly recoverable: boolean,
    options?: ErrorOptions
  ) {
    super(message, options);
    this.name = "GoogleSheetsError";
  }
}

export function toGoogleSheetsError(error: unknown): GoogleSheetsError {
  if (error instanceof GoogleSheetsError) {
    return new GoogleSheetsError(
      error.code,
      SAFE_GOOGLE_SHEETS_ERROR_MESSAGES[error.code],
      error.recoverable
    );
  }
  return new GoogleSheetsError(
    "unknown",
    SAFE_GOOGLE_SHEETS_ERROR_MESSAGES.unknown,
    true
  );
}
