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
  const phase = overrides.phase ?? "setup";
  const clientIdSource = overrides.clientIdSource ?? "missing";
  const storageBusy = overrides.storageBusy ?? false;
  const clientIdEditable = phase === "setup" && ["missing", "session"].includes(clientIdSource);
  const showSheetInput = ["ready", "authorizing", "importing"].includes(phase) || phase === "error";
  return {
    open: true,
    phase,
    origin: "https://sheets.example.com",
    originAssessment: {
      status: "eligible",
      origin: "https://sheets.example.com",
      registration: "unverified"
    },
    clientIdSource,
    clientIdDraft: "",
    sheetDraft: "",
    hostAuthentication: phase !== "blocked" && clientIdSource === "missing",
    clientIdEditable,
    showSheetInput,
    storageBusy,
    canSaveClientId: clientIdEditable && !storageBusy,
    canChangeClientId: ["stored", "session"].includes(clientIdSource) && !storageBusy,
    canForgetClientId: ["stored", "session"].includes(clientIdSource) && !storageBusy,
    canImport: phase === "ready" && !storageBusy,
    canRetry: phase === "error" && !storageBusy,
    setClientIdDraft: vi.fn(),
    setSheetDraft: vi.fn(),
    openDialog: vi.fn(),
    closeDialog: vi.fn(),
    saveClientId: vi.fn().mockResolvedValue(undefined),
    forgetClientId: vi.fn().mockResolvedValue(undefined),
    changeClientId: vi.fn(),
    importSheet: vi.fn(),
    retry: vi.fn(),
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

  it("opens as a labelled modal, focuses setup, and uses the typed no-arg save action", async () => {
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

    expect(saveClientId).toHaveBeenCalledWith();
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

    expect(screen.getByLabelText("OAuth origin")).toHaveValue("https://sheets.example.com");
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
    expect(dialog).toHaveTextContent(/host.*token provider/i);
    expect(dialog).not.toHaveTextContent(/authentication is provided by this app's host/i);
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

  it("offers Change and Forget for stored and session client IDs, but not managed IDs", async () => {
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
          clientIdSource: "session",
          clientIdDraft: CLIENT_ID,
          changeClientId,
          forgetClientId
        })}
        opener={dialogOpener()}
      />
    );
    await user.click(screen.getByRole("button", { name: "Change client ID" }));
    await user.click(screen.getByRole("button", { name: "Forget client ID" }));
    expect(changeClientId).toHaveBeenCalledTimes(2);
    expect(forgetClientId).toHaveBeenCalledTimes(2);

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
    const retry = vi.fn();
    const importSheet = vi.fn();
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
          ),
          retry,
          importSheet
        })}
        opener={dialogOpener()}
      />
    );

    expect(screen.getByRole("alert")).toHaveTextContent(/valid Google Sheets URL/i);
    await waitFor(() => expect(screen.getByLabelText("Google Sheet URL or spreadsheet ID")).toHaveFocus());
    await userEvent.setup().click(screen.getByRole("button", { name: "Retry" }));
    expect(retry).toHaveBeenCalledTimes(1);
    expect(importSheet).not.toHaveBeenCalled();
  });

  it("disables conflicting configuration actions while storage is busy", () => {
    render(
      <GoogleSheetsImportDialog
        controller={setupController({
          phase: "loading",
          clientIdSource: "stored",
          clientIdDraft: CLIENT_ID,
          storageBusy: true,
          storageAction: "forget",
          showSheetInput: true,
          canChangeClientId: false,
          canForgetClientId: false,
          canImport: false
        })}
        opener={dialogOpener()}
      />
    );

    expect(screen.getByRole("dialog", { name: "Import Google Sheet" })).toHaveAttribute("aria-busy", "true");
    expect(screen.getByRole("button", { name: "Change client ID" })).toBeDisabled();
    expect(screen.getByRole("button", { name: "Forget client ID" })).toBeDisabled();
    expect(screen.getByRole("button", { name: "Import and replace workbook" })).toBeDisabled();
    expect(screen.getByRole("status")).toHaveTextContent(/forgetting saved Google OAuth client ID/i);
    expect(screen.getByRole("button", { name: "Cancel" })).toBeEnabled();
  });

  it.each(["preparing", "authorizing", "importing"] as const)(
    "honors explicit stored-ID action legality while %s",
    (phase) => {
      render(
        <GoogleSheetsImportDialog
          controller={setupController({
            phase,
            clientIdSource: "stored",
            clientIdDraft: CLIENT_ID,
            sheetDraft: phase === "preparing" ? "" : "12345678901234567890",
            canChangeClientId: false,
            canForgetClientId: false,
            canImport: false
          })}
          opener={dialogOpener()}
        />
      );

      expect(screen.getByRole("button", { name: "Change client ID" })).toBeDisabled();
      expect(screen.getByRole("button", { name: "Forget client ID" })).toBeDisabled();
    }
  );

  it("shows a nonfatal storage warning without disabling a ready import", () => {
    render(
      <GoogleSheetsImportDialog
        controller={setupController({
          phase: "ready",
          clientIdSource: "session",
          clientIdDraft: CLIENT_ID,
          warning: new GoogleSheetsError(
            "unknown",
            "The Google OAuth client ID could not be saved. You can continue for this session.",
            true
          ),
          showSheetInput: true,
          canImport: true
        })}
        opener={dialogOpener()}
      />
    );

    expect(screen.getByRole("alert")).toHaveTextContent(/could not be saved/i);
    expect(screen.getByRole("button", { name: "Import and replace workbook" })).toBeEnabled();
  });

  it.each(["missing", "rejected"] as const)(
    "selects the exact origin for manual copy when clipboard is %s",
    async (clipboardState) => {
      const user = userEvent.setup();
      const writeText = clipboardState === "rejected"
        ? vi.fn().mockRejectedValue(new Error("private clipboard detail"))
        : undefined;
      Object.defineProperty(navigator, "clipboard", {
        configurable: true,
        value: writeText ? { writeText } : undefined
      });
      const origin = "http://192.168.6.232:4173";
      render(
        <GoogleSheetsImportDialog
          controller={setupController({
            phase: "blocked",
            origin,
            originAssessment: { status: "blocked", origin, reason: "ip_literal" },
            hostAuthentication: false,
            clientIdEditable: false,
            showSheetInput: false,
            canRetry: false,
            error: new GoogleSheetsError("incompatible_origin", "Safe origin guidance.", true)
          })}
          opener={dialogOpener()}
        />
      );

      await user.click(screen.getByRole("button", { name: "Copy" }));

      const originField = screen.getByLabelText("OAuth origin") as HTMLInputElement;
      expect(originField).toHaveFocus();
      expect(originField.selectionStart).toBe(0);
      expect(originField.selectionEnd).toBe(origin.length);
      expect(screen.getByRole("status")).toHaveTextContent(/selected.*(Ctrl\+C|Command\+C)/i);
      expect(document.body).not.toHaveTextContent(/localhost|private clipboard detail/i);
    }
  );

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
