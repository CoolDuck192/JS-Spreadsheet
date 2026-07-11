// @vitest-environment-options {"url":"http://192.168.6.232:4173"}

import { act, render, screen, waitFor, within } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import App, { Spreadsheet } from "./App";
import { createWorkbookSession } from "./core/workbook/WorkbookSession";
import { GoogleSheetsError } from "./lib/googleErrors";
import { createBlankWorkbook, setCellContent } from "./lib/workbook";

const SHEET_ID = "12345678901234567890";

function deferred<T>() {
  let resolve!: (value: T) => void;
  let reject!: (reason?: unknown) => void;
  const promise = new Promise<T>((resolvePromise, rejectPromise) => {
    resolve = resolvePromise;
    reject = rejectPromise;
  });
  return { promise, resolve, reject };
}

function workbookWith(value: string) {
  const workbook = createBlankWorkbook();
  return setCellContent(workbook, workbook.activeSheetId, "A1", value);
}

async function openGoogleImport(user: ReturnType<typeof userEvent.setup>) {
  const tabs = within(screen.getByRole("tablist", { name: "Ribbon tabs" }));
  await user.click(tabs.getByRole("tab", { name: "File" }));
  await user.click(screen.getByRole("button", { name: "Import Google Sheet" }));
}

describe("Task 8 Google Sheets App integration", () => {
  beforeEach(() => {
    localStorage.clear();
    vi.stubEnv("VITE_GOOGLE_CLIENT_ID", "");
  });

  afterEach(() => {
    vi.restoreAllMocks();
    vi.unstubAllGlobals();
    vi.unstubAllEnvs();
  });

  it("Task 8 fresh standalone LAN state explains built-in origin blocking", async () => {
    expect(window.location.origin).toBe("http://192.168.6.232:4173");
    const user = userEvent.setup();
    render(<App />);

    await openGoogleImport(user);

    const dialog = screen.getByRole("dialog", { name: "Import Google Sheet" });
    expect(screen.getByLabelText("OAuth origin")).toHaveValue("http://192.168.6.232:4173");
    expect(dialog).toHaveTextContent(/HTTPS DNS (name|origin)/i);
    expect(dialog).toHaveTextContent(/host.*token provider/i);
    expect(dialog).not.toHaveTextContent(/authentication is provided by this app's host/i);
    expect(dialog).not.toHaveTextContent(/localhost/i);
    expect(screen.queryByLabelText("Google OAuth client ID")).not.toBeInTheDocument();
    await waitFor(() => expect(screen.getByRole("button", { name: "Close" })).toHaveFocus());
  });

  it("Task 8 failed import keeps workbook and history unchanged and sanitizes host output", async () => {
    const initial = workbookWith("before");
    const session = createWorkbookSession({ workbook: initial });
    const onError = vi.fn();
    const tokenProvider = {
      prepare: vi.fn().mockResolvedValue(undefined),
      getAccessToken: vi.fn().mockRejectedValue(
        new GoogleSheetsError(
          "access_denied",
          "Bearer private-token from https://private.example/response-body",
          false
        )
      )
    };
    const user = userEvent.setup();
    render(
      <Spreadsheet
        session={session}
        services={{ googleSheets: { tokenProvider } }}
        onError={onError}
      />
    );

    await openGoogleImport(user);
    await user.type(await screen.findByLabelText("Google Sheet URL or spreadsheet ID"), SHEET_ID);
    await user.click(screen.getByRole("button", { name: "Import and replace workbook" }));

    const alert = await screen.findByRole("alert");
    expect(alert).toHaveTextContent(
      "Google Sheets access was denied. Sign in again or check the sheet sharing settings."
    );
    expect(document.body).not.toHaveTextContent(/private-token|private\.example|response-body/i);
    expect(onError).toHaveBeenCalledWith({
      code: "service.google.access_denied",
      message: "Google Sheets access was denied. Sign in again or check the sheet sharing settings.",
      recoverable: false
    });
    expect(session.getSnapshot()).toMatchObject({
      workbook: initial,
      revision: "0",
      canUndo: false,
      canRedo: false
    });
    session.destroy();
  });

  it("Task 8 canceled import ignores the late result and preserves history", async () => {
    const initial = workbookWith("before");
    const session = createWorkbookSession({ workbook: initial });
    const token = deferred<string>();
    const tokenProvider = {
      prepare: vi.fn().mockResolvedValue(undefined),
      getAccessToken: vi.fn(() => token.promise)
    };
    const fetchMock = vi.fn()
      .mockResolvedValueOnce(new Response(JSON.stringify({
        properties: { title: "Late workbook" },
        sheets: [{ properties: { title: "Late sheet" } }]
      }), { status: 200, headers: { "Content-Type": "application/json" } }))
      .mockResolvedValueOnce(new Response(JSON.stringify({
        valueRanges: [{ values: [["late"]] }]
      }), { status: 200, headers: { "Content-Type": "application/json" } }));
    vi.stubGlobal("fetch", fetchMock);
    const user = userEvent.setup();
    render(<Spreadsheet session={session} services={{ googleSheets: { tokenProvider } }} />);

    await openGoogleImport(user);
    await user.type(await screen.findByLabelText("Google Sheet URL or spreadsheet ID"), SHEET_ID);
    await user.click(screen.getByRole("button", { name: "Import and replace workbook" }));
    expect(
      within(screen.getByRole("dialog", { name: "Import Google Sheet" })).getByRole("status")
    ).toHaveTextContent("Waiting for Google sign-in");

    await user.click(screen.getByRole("button", { name: "Cancel" }));
    expect(document.querySelector("dialog")).not.toHaveAttribute("open");

    await act(async () => token.resolve("ephemeral-token"));
    await waitFor(() => expect(fetchMock).toHaveBeenCalledTimes(2));
    await act(async () => Promise.resolve());

    expect(session.getSnapshot()).toMatchObject({
      workbook: initial,
      revision: "0",
      canUndo: false,
      canRedo: false
    });
    session.destroy();
  });
});
