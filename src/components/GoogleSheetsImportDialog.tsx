import {
  useEffect,
  useId,
  useRef,
  useState,
  type FormEvent,
  type RefObject
} from "react";
import { validateGoogleClientId } from "../lib/googleConfiguration";
import type { GoogleSheetsImportController } from "../react/useGoogleSheetsImport";

type GoogleSheetsImportDialogProps = Readonly<{
  controller: GoogleSheetsImportController;
  opener: RefObject<HTMLElement | null>;
}>;

const PENDING_PHASES = new Set(["preparing", "authorizing", "importing"]);

export function GoogleSheetsImportDialog({
  controller,
  opener
}: GoogleSheetsImportDialogProps) {
  const dialogRef = useRef<HTMLDialogElement>(null);
  const clientIdRef = useRef<HTMLInputElement>(null);
  const sheetRef = useRef<HTMLInputElement>(null);
  const closeRef = useRef<HTMLButtonElement>(null);
  const titleId = useId();
  const descriptionId = useId();
  const clientErrorId = useId();
  const [clientFieldError, setClientFieldError] = useState<string | null>(null);
  const [copyStatus, setCopyStatus] = useState<string | null>(null);
  const busy = PENDING_PHASES.has(controller.phase);
  const setupVisible =
    controller.phase === "setup" &&
    (controller.clientIdSource === "missing" || controller.clientIdSource === "session");
  const sheetVisible = showsSheetInput(controller);

  useEffect(() => {
    const dialog = dialogRef.current;
    if (!dialog) {
      return;
    }
    if (controller.open && !dialog.open) {
      dialog.showModal();
    } else if (!controller.open && dialog.open) {
      dialog.close();
    }
  }, [controller.open]);

  useEffect(() => {
    if (!controller.open) {
      return;
    }
    queueMicrotask(() => {
      if (controller.phase === "blocked") {
        closeRef.current?.focus();
        return;
      }
      if (setupVisible) {
        clientIdRef.current?.focus();
        return;
      }
      if (
        sheetVisible &&
        (controller.phase === "ready" ||
          controller.error?.code === "invalid_sheet_url")
      ) {
        sheetRef.current?.focus();
      }
    });
  }, [
    controller.error?.code,
    controller.open,
    controller.phase,
    setupVisible,
    sheetVisible
  ]);

  function restoreOpenerFocus() {
    opener.current?.focus({ preventScroll: true });
  }

  function submit(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    if (setupVisible) {
      const validation = validateGoogleClientId(controller.clientIdDraft);
      if (!validation.valid) {
        setClientFieldError(validation.message);
        clientIdRef.current?.focus();
        return;
      }
      setClientFieldError(null);
      controller.setClientIdDraft(validation.value);
      void (controller.saveClientId as (clientId?: string) => Promise<void>)(
        validation.value
      );
      return;
    }
    controller.importSheet();
  }

  function close() {
    controller.closeDialog();
  }

  function copyOrigin() {
    setCopyStatus(null);
    void navigator.clipboard?.writeText(controller.origin).then(
      () => setCopyStatus("Origin copied."),
      () => setCopyStatus("Origin could not be copied. Select and copy it manually.")
    );
  }

  return (
    <dialog
      ref={dialogRef}
      className="google-sheets-import-dialog"
      aria-modal="true"
      aria-labelledby={titleId}
      aria-describedby={descriptionId}
      aria-busy={busy}
      onCancel={(event) => {
        event.preventDefault();
        close();
      }}
      onClose={restoreOpenerFocus}
    >
      <form className="google-sheets-import-form" onSubmit={submit} noValidate>
        <header className="google-sheets-import-header">
          <div>
            <h2 id={titleId}>Import Google Sheet</h2>
            <p id={descriptionId}>
              Importing replaces the current workbook. You can undo the replacement from
              workbook history.
            </p>
          </div>
        </header>

        <div className="google-sheets-import-body">
          <section className="google-sheets-import-section" aria-labelledby={`${titleId}-origin`}>
            <div className="google-sheets-import-section-heading">
              <h3 id={`${titleId}-origin`}>OAuth origin</h3>
              <button className="google-sheets-import-button secondary" type="button" onClick={copyOrigin}>
                Copy
              </button>
            </div>
            <code className="google-sheets-import-origin">{controller.origin}</code>
            <OriginStatus controller={controller} />
            {copyStatus ? <p className="google-sheets-import-copy-status" role="status">{copyStatus}</p> : null}
          </section>

          {setupVisible ? (
            <section className="google-sheets-import-section" aria-labelledby={`${titleId}-setup`}>
              <h3 id={`${titleId}-setup`}>Set up Google access</h3>
              <ol className="google-sheets-import-checklist">
                <li>Create or select a Google Cloud project.</li>
                <li>Enable the Google Sheets API.</li>
                <li>Configure the OAuth consent screen and add any required test users.</li>
                <li>Create an OAuth 2.0 client ID for a Web application.</li>
                <li>
                  Add the exact origin shown above to Authorized JavaScript origins.
                </li>
              </ol>
              <p className="google-sheets-import-help">
                Paste only the public client ID. Never enter a client secret.
              </p>
              <label className="google-sheets-import-field">
                <span>Google OAuth client ID</span>
                <input
                  ref={clientIdRef}
                  type="text"
                  value={controller.clientIdDraft}
                  onChange={(event) => {
                    setClientFieldError(null);
                    controller.setClientIdDraft(event.currentTarget.value);
                  }}
                  aria-invalid={clientFieldError ? "true" : undefined}
                  aria-describedby={clientFieldError ? clientErrorId : undefined}
                  autoComplete="off"
                  spellCheck={false}
                  disabled={busy}
                />
              </label>
              {clientFieldError ? (
                <p id={clientErrorId} className="google-sheets-import-alert" role="alert">
                  {clientFieldError}
                </p>
              ) : null}
            </section>
          ) : (
            <ConfigurationSummary controller={controller} />
          )}

          {sheetVisible ? (
            <section className="google-sheets-import-section" aria-labelledby={`${titleId}-sheet`}>
              <h3 id={`${titleId}-sheet`}>Choose a Google Sheet</h3>
              <label className="google-sheets-import-field">
                <span>Google Sheet URL or spreadsheet ID</span>
                <input
                  ref={sheetRef}
                  type="text"
                  value={controller.sheetDraft}
                  onChange={(event) => controller.setSheetDraft(event.currentTarget.value)}
                  aria-invalid={controller.error?.code === "invalid_sheet_url" ? "true" : undefined}
                  autoComplete="off"
                  spellCheck={false}
                  disabled={busy}
                />
              </label>
              <p className="google-sheets-import-warning">
                This one-time import will replace the current workbook, including every sheet.
              </p>
            </section>
          ) : null}

          {controller.error && !clientFieldError ? (
            <p className="google-sheets-import-alert" role="alert">
              {controller.error.message}
            </p>
          ) : null}

          {busy ? <ImportProgress phase={controller.phase} /> : null}
        </div>

        <footer className="google-sheets-import-actions">
          {controller.phase === "blocked" ? (
            <button ref={closeRef} className="google-sheets-import-button primary" type="button" onClick={close}>
              Close
            </button>
          ) : (
            <>
              {setupVisible ? (
                <button className="google-sheets-import-button primary" type="submit" disabled={busy}>
                  Save and continue
                </button>
              ) : controller.phase === "error" ? (
                <button className="google-sheets-import-button primary" type="submit">
                  Retry
                </button>
              ) : (
                <button
                  className="google-sheets-import-button primary"
                  type="submit"
                  disabled={controller.phase !== "ready"}
                >
                  Import and replace workbook
                </button>
              )}
              <button className="google-sheets-import-button secondary" type="button" onClick={close}>
                Cancel
              </button>
            </>
          )}
        </footer>
      </form>
    </dialog>
  );
}

function OriginStatus({ controller }: { controller: GoogleSheetsImportController }) {
  if (controller.error?.code === "origin_mismatch") {
    return (
      <p className="google-sheets-import-origin-status blocked">
        Google rejected this exact origin. Add it to Authorized JavaScript origins and try again.
      </p>
    );
  }
  if (controller.originAssessment.status === "eligible") {
    return (
      <p className="google-sheets-import-origin-status eligible">
        This origin is eligible for browser OAuth. Its registration cannot be verified until Google sign-in.
      </p>
    );
  }
  if (controller.phase !== "blocked") {
    return (
      <p className="google-sheets-import-origin-status eligible">
        Authentication is provided by this app&apos;s host, so the built-in browser OAuth origin restriction
        does not block this import.
      </p>
    );
  }
  if (controller.originAssessment.reason === "ip_literal") {
    return (
      <p className="google-sheets-import-origin-status blocked">
        Google browser OAuth cannot use a raw LAN IP origin. Keep the app available at its current
        address, expose it through an HTTPS DNS name, and add that exact HTTPS DNS origin to
        Authorized JavaScript origins.
      </p>
    );
  }
  return (
    <p className="google-sheets-import-origin-status blocked">
      Google browser OAuth requires an authorized HTTPS DNS origin for this address.
    </p>
  );
}

function ConfigurationSummary({ controller }: { controller: GoogleSheetsImportController }) {
  if (controller.clientIdSource === "missing") {
    if (["loading", "blocked"].includes(controller.phase)) {
      return null;
    }
    return (
      <section className="google-sheets-import-section">
        <h3>Authentication</h3>
        <p className="google-sheets-import-help">Google authentication is provided by this app&apos;s host.</p>
      </section>
    );
  }

  return (
    <section className="google-sheets-import-section" aria-label="Google OAuth configuration">
      <h3>Google OAuth configuration</h3>
      <label className="google-sheets-import-field">
        <span>Google OAuth client ID</span>
        <input type="text" value={controller.clientIdDraft} readOnly />
      </label>
      <p className="google-sheets-import-help">
        {controller.clientIdSource === "managed"
          ? "This client ID is managed by the app host."
          : controller.clientIdSource === "stored"
            ? "This public client ID is saved in this browser."
            : "This public client ID is available for this session."}
      </p>
      {controller.clientIdSource === "stored" ? (
        <div className="google-sheets-import-inline-actions">
          <button className="google-sheets-import-button secondary" type="button" onClick={controller.changeClientId}>
            Change client ID
          </button>
          <button
            className="google-sheets-import-button secondary danger"
            type="button"
            onClick={() => void controller.forgetClientId()}
          >
            Forget client ID
          </button>
        </div>
      ) : null}
    </section>
  );
}

function ImportProgress({ phase }: { phase: GoogleSheetsImportController["phase"] }) {
  const message = phase === "preparing"
    ? "Preparing Google sign-in…"
    : phase === "authorizing"
      ? "Waiting for Google sign-in…"
      : "Importing Google Sheet…";
  return <p className="google-sheets-import-progress" role="status" aria-live="polite">{message}</p>;
}

function showsSheetInput(controller: GoogleSheetsImportController): boolean {
  if (["ready", "authorizing", "importing"].includes(controller.phase)) {
    return true;
  }
  if (controller.phase !== "error") {
    return false;
  }
  return ![
    "configuration_missing",
    "incompatible_origin",
    "invalid_client",
    "gis_load_failed"
  ].includes(controller.error?.code ?? "unknown");
}
