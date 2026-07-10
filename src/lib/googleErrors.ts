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

const UNKNOWN_GOOGLE_SHEETS_ERROR_MESSAGE =
  "Google Sheets import failed. Check your connection and try again.";

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
    return error;
  }
  return new GoogleSheetsError("unknown", UNKNOWN_GOOGLE_SHEETS_ERROR_MESSAGE, true);
}
