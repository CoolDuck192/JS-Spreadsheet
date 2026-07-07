import { describe, expect, it, vi } from "vitest";
import { importWorkbookFromGoogleSheets, parseSpreadsheetId } from "./googleSheets";
import type { TokenProvider } from "./googleAuth";

const SPREADSHEET_ID = "1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms";

const stubTokenProvider: TokenProvider = {
  getAccessToken: vi.fn().mockResolvedValue("test-token")
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

  it("maps permission errors to a friendly message", async () => {
    const fetchImpl = vi.fn(async () => jsonResponse({}, 403));
    await expect(
      importWorkbookFromGoogleSheets(SPREADSHEET_ID, stubTokenProvider, fetchImpl as typeof fetch)
    ).rejects.toThrow(/access denied/i);
  });

  it("rejects invalid input before any network call", async () => {
    const fetchImpl = vi.fn();
    await expect(
      importWorkbookFromGoogleSheets("nope", stubTokenProvider, fetchImpl as unknown as typeof fetch)
    ).rejects.toThrow(/Not a Google Sheets URL/);
    expect(fetchImpl).not.toHaveBeenCalled();
  });
});
