import { fireEvent, render, screen, waitFor } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import { createRef, useMemo, useRef, useState, type RefObject } from "react";
import { afterEach, describe, expect, it, vi } from "vitest";
import { GoogleSheetsError } from "../lib/googleErrors";
import type { GoogleSheetsImportController } from "../react/useGoogleSheetsImport";
import { GoogleSheetsImportDialog } from "./GoogleSheetsImportDialog";

const CLIENT_ID = "123-abc.apps.googleusercontent.com";

function setupController(
  overrides: Partial<GoogleSheetsImportController> = {}
): GoogleSheetsImportController {
  return {
    open: true,
    phase: "setup",
    origin: "https://sheets.example.com",
    originAssessment: {
      status: "eligible",
      origin: "https://sheets.example.com",
      registration: "unverified"
    },
    clientIdSource: "missing",
    clientIdDraft: "",
    sheetDraft: "",
    setClientIdDraft: vi.fn(),
    setSheetDraft: vi.fn(),
    openDialog: vi.fn(),
    closeDialog: vi.fn(),
    saveClientId: vi.fn().mockResolvedValue(undefined),
    forgetClientId: vi.fn().mockResolvedValue(undefined),
    changeClientId: vi.fn(),
    importSheet: vi.fn(),
    ...overrides
  };
}

function dialogOpener(): RefObject<HTMLButtonElement | null> {
  const opener = document.createElement("button");
  opener.textContent = "Open import";
  opener.dataset.externalDialogOpener = "true";
  document.body.appendChild(opener);
  const ref = createRef<HTMLButtonElement>();
  ref.current = opener;
  return ref;
}

describe("GoogleSheetsImportDialog", () => {
  afterEach(() => {
    document.body
      .querySelectorAll("button[data-external-dialog-opener='true']")
      .forEach((button) => button.remove());
    vi.restoreAllMocks();
  });

  it("opens as a labelled modal, focuses setup, and saves the normalized public ID", async () => {
    const user = userEvent.setup();
    const saveClientId = vi.fn().mockResolvedValue(undefined);
    const setClientIdDraft = vi.fn();
    render(
      <GoogleSheetsImportDialog
        controller={setupController({
          clientIdDraft: ` ${CLIENT_ID} `,
          saveClientId,
          setClientIdDraft
        })}
        opener={dialogOpener()}
      />
    );

    const dialog = screen.getByRole("dialog", { name: "Import Google Sheet" });
    expect(dialog).toHaveAttribute("aria-modal", "true");
    await waitFor(() => expect(screen.getByLabelText("Google OAuth client ID")).toHaveFocus());

    await user.click(screen.getByRole("button", { name: "Save and continue" }));

    expect(setClientIdDraft).toHaveBeenCalledWith(CLIENT_ID);
    expect(saveClientId).toHaveBeenCalledWith(CLIENT_ID);
  });

  it("shows and copies the exact eligible-but-unverified origin", async () => {
    const user = userEvent.setup();
    const writeText = vi.fn().mockResolvedValue(undefined);
    Object.defineProperty(navigator, "clipboard", {
      configurable: true,
      value: { writeText }
    });
    render(
      <GoogleSheetsImportDialog
        controller={setupController()}
        opener={dialogOpener()}
      />
    );

    expect(screen.getByText("https://sheets.example.com", { selector: "code" })).toBeInTheDocument();
    expect(screen.getByText(/registration cannot be verified until Google sign-in/i)).toBeInTheDocument();

    await user.click(screen.getByRole("button", { name: "Copy" }));

    expect(writeText).toHaveBeenCalledWith("https://sheets.example.com");
  });

  it("gives actionable HTTPS DNS guidance for a LAN IP without recommending localhost", async () => {
    const closeDialog = vi.fn();
    render(
      <GoogleSheetsImportDialog
        controller={setupController({
          phase: "blocked",
          origin: "http://192.168.6.232:4173",
          originAssessment: {
            status: "blocked",
            origin: "http://192.168.6.232:4173",
            reason: "ip_literal"
          },
          error: new GoogleSheetsError(
            "incompatible_origin",
            "Google browser OAuth cannot use this origin.",
            true
          ),
          closeDialog
        })}
        opener={dialogOpener()}
      />
    );

    const dialog = screen.getByRole("dialog", { name: "Import Google Sheet" });
    expect(dialog).toHaveTextContent("http://192.168.6.232:4173");
    expect(dialog).toHaveTextContent(/HTTPS DNS (name|origin)/i);
    expect(dialog).not.toHaveTextContent(/localhost/i);
    await waitFor(() => expect(screen.getByRole("button", { name: "Close" })).toHaveFocus());
  });

  it("focuses the sheet field when ready and makes replacement explicit", async () => {
    render(
      <GoogleSheetsImportDialog
        controller={setupController({
          phase: "ready",
          clientIdSource: "managed",
          clientIdDraft: CLIENT_ID
        })}
        opener={dialogOpener()}
      />
    );

    await waitFor(() => expect(screen.getByLabelText("Google Sheet URL or spreadsheet ID")).toHaveFocus());
    expect(screen.getByText(/replace the current workbook/i)).toBeInTheDocument();
    expect(screen.getByRole("button", { name: "Import and replace workbook" })).toBeEnabled();
  });

  it("offers Change and Forget only for a stored client ID", async () => {
    const user = userEvent.setup();
    const changeClientId = vi.fn();
    const forgetClientId = vi.fn().mockResolvedValue(undefined);
    const { rerender } = render(
      <GoogleSheetsImportDialog
        controller={setupController({
          phase: "ready",
          clientIdSource: "stored",
          clientIdDraft: CLIENT_ID,
          changeClientId,
          forgetClientId
        })}
        opener={dialogOpener()}
      />
    );

    await user.click(screen.getByRole("button", { name: "Change client ID" }));
    await user.click(screen.getByRole("button", { name: "Forget client ID" }));
    expect(changeClientId).toHaveBeenCalledTimes(1);
    expect(forgetClientId).toHaveBeenCalledTimes(1);

    rerender(
      <GoogleSheetsImportDialog
        controller={setupController({
          phase: "ready",
          clientIdSource: "managed",
          clientIdDraft: CLIENT_ID
        })}
        opener={dialogOpener()}
      />
    );
    expect(screen.queryByRole("button", { name: "Change client ID" })).not.toBeInTheDocument();
    expect(screen.queryByRole("button", { name: "Forget client ID" })).not.toBeInTheDocument();
  });

  it("focuses the first invalid setup field and reports an inline alert", async () => {
    const user = userEvent.setup();
    render(
      <GoogleSheetsImportDialog
        controller={setupController({ clientIdDraft: "not-a-client-id" })}
        opener={dialogOpener()}
      />
    );

    await user.click(screen.getByRole("button", { name: "Save and continue" }));

    expect(screen.getByRole("alert")).toHaveTextContent(/Web application client ID/i);
    expect(screen.getByLabelText("Google OAuth client ID")).toHaveFocus();
  });

  it("focuses an invalid sheet field and keeps the typed error inline", async () => {
    render(
      <GoogleSheetsImportDialog
        controller={setupController({
          phase: "error",
          clientIdSource: "managed",
          clientIdDraft: CLIENT_ID,
          sheetDraft: "bad sheet",
          error: new GoogleSheetsError(
            "invalid_sheet_url",
            "Enter a valid Google Sheets URL or spreadsheet ID.",
            true
          )
        })}
        opener={dialogOpener()}
      />
    );

    expect(screen.getByRole("alert")).toHaveTextContent(/valid Google Sheets URL/i);
    await waitFor(() => expect(screen.getByLabelText("Google Sheet URL or spreadsheet ID")).toHaveFocus());
    expect(screen.getByRole("button", { name: "Retry" })).toBeInTheDocument();
  });

  it.each([
    ["preparing", "Preparing Google sign-in"],
    ["authorizing", "Waiting for Google sign-in"],
    ["importing", "Importing Google Sheet"]
  ] as const)("marks %s work busy while leaving Cancel available", (phase, progress) => {
    render(
      <GoogleSheetsImportDialog
        controller={setupController({
          phase,
          clientIdSource: "managed",
          clientIdDraft: CLIENT_ID,
          sheetDraft: phase === "preparing" ? "" : "12345678901234567890"
        })}
        opener={dialogOpener()}
      />
    );

    const dialog = screen.getByRole("dialog", { name: "Import Google Sheet" });
    expect(dialog).toHaveAttribute("aria-busy", "true");
    expect(screen.getByRole("status")).toHaveTextContent(progress);
    expect(screen.getByRole("button", { name: "Cancel" })).toBeEnabled();
    expect(screen.queryByRole("button", { name: "Import and replace workbook" })).not.toBeEnabled();
  });

  it("handles native Escape/cancel and restores focus to the opener", async () => {
    function Harness() {
      const opener = useRef<HTMLButtonElement>(null);
      const [open, setOpen] = useState(true);
      const controller = useMemo(
        () => setupController({
          open,
          phase: open ? "ready" : "closed",
          clientIdSource: "managed",
          clientIdDraft: CLIENT_ID,
          closeDialog: () => setOpen(false)
        }),
        [open]
      );
      return (
        <>
          <button ref={opener} type="button">Open import</button>
          <GoogleSheetsImportDialog controller={controller} opener={opener} />
        </>
      );
    }

    render(<Harness />);
    const dialog = screen.getByRole("dialog", { name: "Import Google Sheet" });
    await waitFor(() => expect(screen.getByLabelText("Google Sheet URL or spreadsheet ID")).toHaveFocus());

    fireEvent(dialog, new Event("cancel", { bubbles: false, cancelable: true }));

    await waitFor(() => expect(screen.getByRole("button", { name: "Open import" })).toHaveFocus());
    expect(dialog).not.toHaveAttribute("open");
  });
});
