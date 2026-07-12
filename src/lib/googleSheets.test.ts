import { beforeEach, describe, expect, it, vi } from "vitest";
import { GoogleSheetsError, toGoogleSheetsError } from "./googleErrors";
import { importWorkbookFromGoogleSheets, parseSpreadsheetId } from "./googleSheets";
import { SHEETS_READONLY_SCOPE, type TokenProvider } from "./googleAuth";

const SPREADSHEET_ID = "1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms";

const getAccessToken = vi.fn<(scopes: readonly string[]) => Promise<string>>();
const invalidateAccessToken = vi.fn<(scopes: readonly string[]) => void | Promise<void>>();
const stubTokenProvider = {
  getAccessToken,
  invalidateAccessToken
} as TokenProvider & {
  invalidateAccessToken(scopes: readonly string[]): void | Promise<void>;
};

function jsonResponse(body: unknown, status = 200): Response {
  return {
    ok: status >= 200 && status < 300,
    status,
    json: async () => body
  } as Response;
}

describe("parseSpreadsheetId", () => {
  it("extracts the id from a share URL", () => {
    expect(
      parseSpreadsheetId(`https://docs.google.com/spreadsheets/d/${SPREADSHEET_ID}/edit#gid=0`)
    ).toBe(SPREADSHEET_ID);
  });

  it("accepts a bare id", () => {
    expect(parseSpreadsheetId(SPREADSHEET_ID)).toBe(SPREADSHEET_ID);
  });

  it("rejects unrelated text", () => {
    expect(parseSpreadsheetId("not a sheet")).toBeNull();
    expect(parseSpreadsheetId("")).toBeNull();
  });
});

describe("importWorkbookFromGoogleSheets", () => {
  beforeEach(() => {
    getAccessToken.mockReset();
    getAccessToken.mockResolvedValue("test-token");
    invalidateAccessToken.mockReset();
  });

  it("imports sheets, values, and formulas into a workbook model", async () => {
    const fetchImpl = vi.fn(async (url: RequestInfo | URL, init?: RequestInit) => {
      void init;
      const href = String(url);
      if (href.includes("/values:batchGet")) {
        expect(href).toContain("valueRenderOption=FORMULA");
        return jsonResponse({
          valueRanges: [
            { values: [["Project", 42], ["=LEN(A1)", true]] },
            { values: [] }
          ]
        });
      }
      return jsonResponse({
        properties: { title: "Budget 2026" },
        sheets: [
          { properties: { title: "Data", hidden: false } },
          { properties: { title: "Archive", hidden: true } }
        ]
      });
    });

    const result = await importWorkbookFromGoogleSheets(SPREADSHEET_ID, stubTokenProvider, fetchImpl as typeof fetch);

    expect(result.spreadsheetTitle).toBe("Budget 2026");
    expect(result.workbook.sheets).toHaveLength(2);
    const [data, archive] = result.workbook.sheets;
    expect(data.name).toBe("Data");
    expect(data.cells).toEqual({ A1: "Project", B1: 42, A2: "=LEN(A1)", B2: true });
    expect(archive.isHidden).toBe(true);
    expect(result.workbook.activeSheetId).toBe(data.id);

    const authHeader = (fetchImpl.mock.calls[0][1]?.headers ?? {}) as Record<string, string>;
    expect(authHeader.Authorization).toBe("Bearer test-token");
  });

  it("rejects invalid input before any network call", async () => {
    const fetchImpl = vi.fn();
    const error = await captureGoogleError(
      importWorkbookFromGoogleSheets("nope", stubTokenProvider, fetchImpl as unknown as typeof fetch)
    );

    expect(error.code).toBe("invalid_sheet_url");
    expect(getAccessToken).not.toHaveBeenCalled();
    expect(fetchImpl).not.toHaveBeenCalled();
  });

  it.each([401, 403])("maps HTTP %s to access_denied", async (status) => {
    const fetchImpl = vi.fn(async () => jsonResponse({ error: { message: "private response body" } }, status));

    const error = await captureGoogleError(
      importWorkbookFromGoogleSheets(SPREADSHEET_ID, stubTokenProvider, fetchImpl as typeof fetch)
    );

    expect(error.code).toBe("access_denied");
    expect(error.message).not.toContain("private response body");
    expect(error.message).not.toContain("test-token");
    expect(invalidateAccessToken).toHaveBeenCalledWith([SHEETS_READONLY_SCOPE]);
  });

  it("keeps the access-denied error when token invalidation fails", async () => {
    invalidateAccessToken.mockRejectedValueOnce(new Error("private invalidation failure"));
    const fetchImpl = vi.fn(async () => jsonResponse({}, 401));

    const error = await captureGoogleError(
      importWorkbookFromGoogleSheets(SPREADSHEET_ID, stubTokenProvider, fetchImpl as typeof fetch)
    );

    expect(invalidateAccessToken).toHaveBeenCalledWith([SHEETS_READONLY_SCOPE]);
    expect(error).toMatchObject({ code: "access_denied", recoverable: true });
    expect(error.message).not.toContain("private invalidation failure");
  });

  it("distinguishes a disabled Sheets API from ordinary HTTP 403", async () => {
    const fetchImpl = vi.fn(async () =>
      jsonResponse(
        {
          error: {
            status: "PERMISSION_DENIED",
            message: "private API response body",
            errors: [{ reason: "accessNotConfigured" }]
          }
        },
        403
      )
    );

    const error = await captureGoogleError(
      importWorkbookFromGoogleSheets(SPREADSHEET_ID, stubTokenProvider, fetchImpl as typeof fetch)
    );

    expect(error.code).toBe("api_not_enabled");
    expect(error.message).not.toContain("private API response body");
    expect(error.message).not.toContain("test-token");
    expect(invalidateAccessToken).not.toHaveBeenCalled();
  });

  it("recognizes modern SERVICE_DISABLED details without exposing metadata or help URLs", async () => {
    const fetchImpl = vi.fn(async () =>
      jsonResponse(
        {
          error: {
            status: "PERMISSION_DENIED",
            message: "private response body",
            details: [{
              "@type": "type.googleapis.com/google.rpc.ErrorInfo",
              reason: "SERVICE_DISABLED",
              metadata: {
                consumer: "projects/private-project",
                service: "sheets.googleapis.com"
              }
            }, {
              "@type": "type.googleapis.com/google.rpc.Help",
              links: [{ description: "private", url: "https://private.example/enable?token=secret" }]
            }]
          }
        },
        403
      )
    );

    const error = await captureGoogleError(
      importWorkbookFromGoogleSheets(SPREADSHEET_ID, stubTokenProvider, fetchImpl as typeof fetch)
    );

    expect(error).toMatchObject({ code: "api_not_enabled", recoverable: true });
    expect(error.message).toBe("The Google Sheets API is not enabled for this Google OAuth project.");
    expect(error.message).not.toMatch(/private-project|private\.example|token=secret|sheets\.googleapis\.com/i);
  });

  it.each([
    [404, "sheet_not_found"],
    [429, "rate_limited"],
    [500, "unknown"]
  ] as const)("maps HTTP %s to %s", async (status, code) => {
    const fetchImpl = vi.fn(async () => jsonResponse({ error: { message: "private response body" } }, status));

    const error = await captureGoogleError(
      importWorkbookFromGoogleSheets(SPREADSHEET_ID, stubTokenProvider, fetchImpl as typeof fetch)
    );

    expect(error.code).toBe(code);
    expect(error.message).not.toContain("private response body");
    expect(error.message).not.toContain("test-token");
  });

  it("maps a rejected network request without surfacing its details", async () => {
    const fetchImpl = vi.fn(async () => {
      throw new Error("Bearer test-token: private network response body");
    });

    const error = await captureGoogleError(
      importWorkbookFromGoogleSheets(SPREADSHEET_ID, stubTokenProvider, fetchImpl as unknown as typeof fetch)
    );

    expect(error.code).toBe("network_failed");
    expect(error.message).not.toContain("test-token");
    expect(error.message).not.toContain("private network response body");
  });

  it("sanitizes untyped token provider failures", async () => {
    getAccessToken.mockRejectedValueOnce(new Error("Bearer private-token: private OAuth response body"));

    const error = await captureGoogleError(
      importWorkbookFromGoogleSheets(SPREADSHEET_ID, stubTokenProvider, vi.fn() as unknown as typeof fetch)
    );

    expect(error.code).toBe("unknown");
    expect(error.message).not.toContain("private-token");
    expect(error.message).not.toContain("private OAuth response body");
  });

  it("remaps typed provider failures to a library-owned safe message", async () => {
    getAccessToken.mockRejectedValueOnce(
      new GoogleSheetsError(
        "access_denied",
        "Bearer private-token from https://private.example/response-body",
        false
      )
    );

    const error = await captureGoogleError(
      importWorkbookFromGoogleSheets(SPREADSHEET_ID, stubTokenProvider, vi.fn() as unknown as typeof fetch)
    );

    expect(error).toMatchObject({ code: "access_denied", recoverable: false });
    expect(error.message).toBe("Google Sheets access was denied. Sign in again or check the sheet sharing settings.");
    expect(error.message).not.toMatch(/private-token|private\.example|response-body/i);
  });
});

describe("toGoogleSheetsError", () => {
  it("remaps typed and arbitrary failures to library-owned safe messages", () => {
    const typed = new GoogleSheetsError(
      "popup_closed",
      "Bearer private-token from https://private.example/response-body",
      false
    );
    const sanitizedTyped = toGoogleSheetsError(typed);
    expect(sanitizedTyped).not.toBe(typed);
    expect(sanitizedTyped).toMatchObject({ code: "popup_closed", recoverable: false });
    expect(sanitizedTyped.message).toBe("Google sign-in was closed before access was granted.");
    expect(sanitizedTyped.message).not.toMatch(/private-token|private\.example|response-body/i);

    const sanitized = toGoogleSheetsError(new Error("Bearer private-token: private response body"));
    expect(sanitized).toMatchObject({ code: "unknown", recoverable: true });
    expect(sanitized.message).not.toContain("private-token");
    expect(sanitized.message).not.toContain("private response body");
  });
});

async function captureGoogleError(promise: Promise<unknown>): Promise<GoogleSheetsError> {
  try {
    await promise;
  } catch (error) {
    expect(error).toBeInstanceOf(GoogleSheetsError);
    return error as GoogleSheetsError;
  }
  throw new Error("Expected Google Sheets operation to reject");
}
