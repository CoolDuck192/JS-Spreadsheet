import { useCallback, useEffect, useLayoutEffect, useMemo, useRef, useState } from "react";
import type {
  GoogleSheetsServiceConfiguration,
  TokenProvider
} from "../core/workbook/services";
import { createBrowserTokenProvider } from "../lib/googleAuth";
import {
  assessGoogleOAuthOrigin,
  resolveGoogleAuthSource,
  validateGoogleClientId,
  type GoogleOAuthOriginAssessment
} from "../lib/googleConfiguration";
import { GoogleSheetsError, toGoogleSheetsError } from "../lib/googleErrors";
import {
  importWorkbookFromGoogleSheets,
  parseSpreadsheetId,
  type GoogleSheetsImportResult
} from "../lib/googleSheets";
import {
  createGoogleTokenProviderRuntime,
  type GoogleTokenProviderPreparation
} from "./googleTokenProviderRuntime";
import { useGoogleClientIdConfiguration } from "./useGoogleClientIdConfiguration";

export type GoogleSheetsImportPhase =
  | "closed"
  | "loading"
  | "setup"
  | "blocked"
  | "preparing"
  | "ready"
  | "authorizing"
  | "importing"
  | "error";

export type GoogleSheetsImportController = Readonly<{
  open: boolean;
  phase: GoogleSheetsImportPhase;
  origin: string;
  originAssessment: GoogleOAuthOriginAssessment;
  clientIdSource: "managed" | "stored" | "session" | "missing";
  clientIdDraft: string;
  sheetDraft: string;
  hostAuthentication: boolean;
  clientIdEditable: boolean;
  showSheetInput: boolean;
  storageBusy: boolean;
  storageAction?: "save" | "forget";
  canSaveClientId: boolean;
  canChangeClientId: boolean;
  canForgetClientId: boolean;
  canImport: boolean;
  canRetry: boolean;
  error?: GoogleSheetsError;
  warning?: GoogleSheetsError;
  setClientIdDraft(value: string): void;
  setSheetDraft(value: string): void;
  openDialog(): void;
  closeDialog(): void;
  saveClientId(): Promise<void>;
  forgetClientId(): Promise<void>;
  changeClientId(): void;
  importSheet(): void;
  retry(): void;
}>;

type ReadinessPhase = "loading" | "setup" | "blocked" | "preparing" | "ready" | "error";
type ImportOperation = Readonly<{
  phase: "authorizing" | "importing" | "error";
  error?: GoogleSheetsError;
}>;

export function useGoogleSheetsImport(options: Readonly<{
  configuration?: GoogleSheetsServiceConfiguration;
  deprecatedTokenProviderFactory?: (clientId: string) => TokenProvider;
  origin: string;
  onImported(result: GoogleSheetsImportResult): void | Promise<void>;
  onError?(error: GoogleSheetsError): void;
}>): GoogleSheetsImportController {
  const { configuration, deprecatedTokenProviderFactory, origin } = options;
  const [open, setOpen] = useState(false);
  const [sheetDraft, setSheetDraftState] = useState("");
  const [readiness, setReadinessState] = useState<ReadinessPhase>("loading");
  const [readinessError, setReadinessError] = useState<GoogleSheetsError | undefined>();
  const [readinessRetryable, setReadinessRetryableState] = useState(false);
  const [operation, setOperation] = useState<ImportOperation | null>(null);
  const [preparationAttempt, setPreparationAttempt] = useState(0);

  const readinessRef = useRef<ReadinessPhase>("loading");
  const readinessRetryableRef = useRef(false);
  const providerRef = useRef<TokenProvider | null>(null);
  const attemptGenerationRef = useRef(0);
  const importPendingRef = useRef(false);
  const onErrorRef = useRef(options.onError);
  onErrorRef.current = options.onError;

  const managedClientId = configuration?.clientId?.trim() ?? "";

  const updateReadiness = useCallback((next: ReadinessPhase) => {
    readinessRef.current = next;
    setReadinessState(next);
  }, []);

  const updateReadinessRetryable = useCallback((next: boolean) => {
    readinessRetryableRef.current = next;
    setReadinessRetryableState(next);
  }, []);

  const notifyError = useCallback((error: GoogleSheetsError) => {
    try {
      onErrorRef.current?.(error);
    } catch {
      // Host observability must never replace the dialog's safe inline error.
    }
  }, []);

  const clientIdConfiguration = useGoogleClientIdConfiguration({
    storage: configuration?.clientIdStorage,
    ignoreLoadFailure: Boolean(configuration?.tokenProvider || managedClientId),
    onError: notifyError
  });
  const [providerRuntime] = useState(createGoogleTokenProviderRuntime);
  const originAssessment = useMemo(() => assessGoogleOAuthOrigin(origin), [origin]);
  const authSource = resolveGoogleAuthSource(
    configuration,
    clientIdConfiguration.editableClientId,
    deprecatedTokenProviderFactory,
    createBrowserTokenProvider
  );
  const hostAuthentication = Boolean(
    configuration?.tokenProvider ||
    configuration?.tokenProviderFactory ||
    deprecatedTokenProviderFactory
  );
  const usesBuiltInGoogleAuth = !hostAuthentication;
  const clientIdSource = authSource.kind === "client" ? authSource.source : "missing";
  const displayedClientId = clientIdSource === "managed"
    ? managedClientId
    : clientIdConfiguration.clientIdDraft;

  useEffect(() => {
    let active = true;

    if (usesBuiltInGoogleAuth && originAssessment.status === "blocked") {
      providerRef.current = null;
      const error = incompatibleOriginError(originAssessment);
      setReadinessError(error);
      updateReadinessRetryable(false);
      updateReadiness("blocked");
      return () => {
        active = false;
      };
    }

    if (authSource.kind === "missing") {
      providerRef.current = null;
      setReadinessError(undefined);
      updateReadinessRetryable(false);
      updateReadiness(
        clientIdConfiguration.storageStatus === "loading" ? "loading" : "setup"
      );
      return () => {
        active = false;
      };
    }

    if (
      authSource.kind === "client" &&
      clientIdConfiguration.editingClientId &&
      authSource.source !== "managed"
    ) {
      providerRef.current = null;
      setReadinessError(undefined);
      updateReadinessRetryable(false);
      updateReadiness("setup");
      return () => {
        active = false;
      };
    }

    if (authSource.kind === "client") {
      const validation = validateGoogleClientId(authSource.clientId);
      if (!validation.valid) {
        providerRef.current = null;
        const error = new GoogleSheetsError("invalid_client", validation.message, true);
        setReadinessError(error);
        updateReadinessRetryable(false);
        updateReadiness(authSource.source === "managed" ? "error" : "setup");
        return () => {
          active = false;
        };
      }
    }

    let preparation: GoogleTokenProviderPreparation;
    try {
      preparation = providerRuntime.resolve(authSource);
    } catch (caught) {
      const error = toGoogleSheetsError(caught);
      providerRef.current = null;
      setReadinessError(error);
      updateReadinessRetryable(true);
      updateReadiness("error");
      notifyError(error);
      return () => {
        active = false;
      };
    }

    providerRef.current = preparation.provider;
    setReadinessError(undefined);
    updateReadinessRetryable(false);
    updateReadiness("preparing");
    preparation.prepare().then(
      () => {
        if (!active) {
          return;
        }
        setReadinessError(undefined);
        updateReadinessRetryable(false);
        updateReadiness("ready");
      },
      (caught) => {
        if (!active) {
          return;
        }
        const error = toGoogleSheetsError(caught);
        setReadinessError(error);
        updateReadinessRetryable(true);
        updateReadiness("error");
        notifyError(error);
      }
    );

    return () => {
      active = false;
    };
  }, [
    authSource.kind,
    authSource.kind === "provider" ? authSource.provider : undefined,
    authSource.kind === "client" ? authSource.clientId : undefined,
    authSource.kind === "client" ? authSource.source : undefined,
    authSource.kind === "client" ? authSource.factory : undefined,
    clientIdConfiguration.editingClientId,
    clientIdConfiguration.storageStatus,
    notifyError,
    originAssessment,
    preparationAttempt,
    providerRuntime,
    updateReadiness,
    updateReadinessRetryable,
    usesBuiltInGoogleAuth
  ]);

  const setClientIdDraft = useCallback((value: string) => {
    clientIdConfiguration.setClientIdDraft(value);
    setReadinessError((current) => current?.code === "invalid_client" ? undefined : current);
  }, [clientIdConfiguration.setClientIdDraft]);

  const setSheetDraft = useCallback((value: string) => {
    setSheetDraftState(value);
    setOperation((current) => current?.error?.code === "invalid_sheet_url" ? null : current);
  }, []);

  const openDialog = useCallback(() => {
    setOpen(true);
  }, []);

  const invalidateImportAttempt = useCallback(() => {
    attemptGenerationRef.current += 1;
    importPendingRef.current = false;
    setOperation(null);
  }, []);

  // Invalidate during the commit itself so a settled import microtask cannot
  // reach a stale target between a target/auth change and passive cleanup.
  useLayoutEffect(() => {
    invalidateImportAttempt();
    return () => {
      attemptGenerationRef.current += 1;
      importPendingRef.current = false;
    };
  }, [
    authSource.kind,
    authSource.kind === "provider" ? authSource.provider : undefined,
    authSource.kind === "client" ? authSource.clientId : undefined,
    authSource.kind === "client" ? authSource.source : undefined,
    authSource.kind === "client" ? authSource.factory : undefined,
    configuration?.clientIdStorage,
    invalidateImportAttempt,
    options.onImported,
    origin
  ]);

  const closeDialog = useCallback(() => {
    invalidateImportAttempt();
    setOpen(false);
  }, [invalidateImportAttempt]);

  const saveClientId = useCallback(async () => {
    const validation = validateGoogleClientId(clientIdConfiguration.clientIdDraft);
    if (!validation.valid) {
      const error = new GoogleSheetsError("invalid_client", validation.message, true);
      setReadinessError(error);
      updateReadinessRetryable(false);
      setOperation(null);
      updateReadiness("setup");
      return;
    }

    const persistence = clientIdConfiguration.saveClientId(validation.value);
    setReadinessError(undefined);
    updateReadinessRetryable(false);
    setOperation(null);
    await persistence;
  }, [
    clientIdConfiguration.clientIdDraft,
    clientIdConfiguration.saveClientId,
    updateReadiness,
    updateReadinessRetryable
  ]);

  const forgetClientId = useCallback(async () => {
    invalidateImportAttempt();
    const outcome = await clientIdConfiguration.forgetClientId();
    if (outcome.status === "cleared") {
      providerRuntime.clear();
      providerRef.current = null;
      setReadinessError(undefined);
      updateReadinessRetryable(false);
      updateReadiness("setup");
    }
  }, [
    clientIdConfiguration.forgetClientId,
    invalidateImportAttempt,
    providerRuntime,
    updateReadiness,
    updateReadinessRetryable
  ]);

  const changeClientId = useCallback(() => {
    if (clientIdConfiguration.isStorageBusy()) {
      return;
    }
    invalidateImportAttempt();
    providerRuntime.clear();
    providerRef.current = null;
    clientIdConfiguration.changeClientId();
    setReadinessError(undefined);
    updateReadinessRetryable(false);
    updateReadiness("setup");
  }, [
    clientIdConfiguration.changeClientId,
    clientIdConfiguration.isStorageBusy,
    invalidateImportAttempt,
    providerRuntime,
    updateReadiness,
    updateReadinessRetryable
  ]);

  const importSheet = useCallback(() => {
    if (importPendingRef.current || clientIdConfiguration.isStorageBusy()) {
      return;
    }

    const provider = providerRef.current;
    if (readinessRef.current !== "ready" || !provider) {
      return;
    }

    if (!parseSpreadsheetId(sheetDraft)) {
      const error = new GoogleSheetsError(
        "invalid_sheet_url",
        "Enter a valid Google Sheets URL or spreadsheet ID.",
        true
      );
      setOperation({ phase: "error", error });
      notifyError(error);
      return;
    }

    const generation = ++attemptGenerationRef.current;
    const importTarget = options.onImported;
    importPendingRef.current = true;
    setOperation({ phase: "authorizing" });

    const trackedProvider: TokenProvider = {
      getAccessToken(scopes) {
        let pendingToken: Promise<string>;
        try {
          pendingToken = provider.getAccessToken(scopes);
        } catch (caught) {
          return Promise.reject(caught);
        }
        return Promise.resolve(pendingToken).then((token) => {
          if (
            generation === attemptGenerationRef.current &&
            importPendingRef.current
          ) {
            setOperation({ phase: "importing" });
          }
          return token;
        });
      }
    };

    let pendingImport: Promise<GoogleSheetsImportResult>;
    try {
      // This call intentionally occurs before any await so OAuth can consume the
      // button's user activation synchronously.
      pendingImport = importWorkbookFromGoogleSheets(sheetDraft, trackedProvider);
    } catch (caught) {
      finishImportFailure(caught, generation);
      return;
    }

    pendingImport.then(
      async (imported) => {
        if (generation !== attemptGenerationRef.current || !importPendingRef.current) {
          return;
        }
        try {
          await importTarget(imported);
        } catch (caught) {
          finishImportFailure(caught, generation);
          return;
        }
        if (generation !== attemptGenerationRef.current || !importPendingRef.current) {
          return;
        }
        importPendingRef.current = false;
        attemptGenerationRef.current += 1;
        setSheetDraftState("");
        setReadinessError(undefined);
        setOperation(null);
        setOpen(false);
      },
      (caught) => finishImportFailure(caught, generation)
    );

    function finishImportFailure(caught: unknown, failedGeneration: number) {
      if (
        failedGeneration !== attemptGenerationRef.current ||
        !importPendingRef.current
      ) {
        return;
      }
      importPendingRef.current = false;
      const error = toGoogleSheetsError(caught);
      setOperation({ phase: "error", error });
      notifyError(error);
    }
  }, [clientIdConfiguration.isStorageBusy, notifyError, options.onImported, sheetDraft]);

  const retry = useCallback(() => {
    if (clientIdConfiguration.isStorageBusy()) {
      return;
    }
    if (operation?.phase === "error" && operation.error?.recoverable !== false) {
      importSheet();
      return;
    }
    if (readinessRef.current !== "error" || !readinessRetryableRef.current) {
      return;
    }
    attemptGenerationRef.current += 1;
    setReadinessError(undefined);
    updateReadinessRetryable(false);
    setOperation(null);
    updateReadiness("preparing");
    setPreparationAttempt((current) => current + 1);
  }, [
    clientIdConfiguration.isStorageBusy,
    importSheet,
    operation,
    updateReadiness,
    updateReadinessRetryable
  ]);

  const storageBusy = clientIdConfiguration.storageBusy;
  const phase: GoogleSheetsImportPhase = !open
    ? "closed"
    : storageBusy
      ? "loading"
      : operation?.phase ?? readiness;
  const error = operation?.error ?? readinessError ?? clientIdConfiguration.clientIdError;
  const clientIdEditable =
    readiness === "setup" &&
    authSource.kind !== "provider" &&
    clientIdSource !== "managed";
  const showSheetInput = readiness === "ready" || operation !== null;
  const configurationActionPending =
    readiness === "preparing" ||
    operation?.phase === "authorizing" ||
    operation?.phase === "importing";
  const canRetry =
    !storageBusy &&
    ((operation?.phase === "error" && operation.error?.recoverable !== false) ||
      (readiness === "error" && readinessRetryable));

  return {
    open,
    phase,
    origin,
    originAssessment,
    clientIdSource,
    clientIdDraft: displayedClientId,
    sheetDraft,
    hostAuthentication,
    clientIdEditable,
    showSheetInput,
    storageBusy,
    ...(clientIdConfiguration.storageAction
      ? { storageAction: clientIdConfiguration.storageAction }
      : {}),
    canSaveClientId: clientIdEditable && !storageBusy,
    canChangeClientId:
      clientIdSource === "stored" && !storageBusy && !configurationActionPending,
    canForgetClientId:
      clientIdSource === "stored" && !storageBusy && !configurationActionPending,
    canImport:
      readiness === "ready" && operation === null && !storageBusy && !importPendingRef.current,
    canRetry,
    ...(error ? { error } : {}),
    ...(clientIdConfiguration.warning
      ? { warning: clientIdConfiguration.warning }
      : {}),
    setClientIdDraft,
    setSheetDraft,
    openDialog,
    closeDialog,
    saveClientId,
    forgetClientId,
    changeClientId,
    importSheet,
    retry
  };
}

function incompatibleOriginError(
  assessment: Extract<GoogleOAuthOriginAssessment, { status: "blocked" }>
): GoogleSheetsError {
  const message = assessment.reason === "ip_literal"
    ? "Google browser OAuth requires an authorized HTTPS DNS origin and cannot use a raw LAN IP."
    : "Google browser OAuth requires an eligible, authorized HTTPS origin.";
  return new GoogleSheetsError("incompatible_origin", message, true);
}
