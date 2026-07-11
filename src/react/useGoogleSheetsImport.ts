import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import type {
  GoogleClientIdStorage,
  GoogleSheetsServiceConfiguration,
  TokenProvider
} from "../core/workbook/services";
import { createBrowserTokenProvider } from "../lib/googleAuth";
import {
  assessGoogleOAuthOrigin,
  resolveGoogleAuthSource,
  validateGoogleClientId,
  type EditableGoogleClientId,
  type GoogleOAuthOriginAssessment
} from "../lib/googleConfiguration";
import { GoogleSheetsError, toGoogleSheetsError } from "../lib/googleErrors";
import {
  importWorkbookFromGoogleSheets,
  parseSpreadsheetId,
  type GoogleSheetsImportResult
} from "../lib/googleSheets";
import { createSerialOperationQueue } from "./serialOperationQueue";

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

type ProviderRecord = {
  kind: "direct" | "factory";
  identity: TokenProvider | ((clientId: string) => TokenProvider);
  clientId: string;
  provider: TokenProvider;
  prepared: boolean;
  preparation?: Promise<void>;
};

type StorageLoad = {
  storage: GoogleClientIdStorage;
  result: string | null | Promise<string | null>;
};

const STORAGE_LOAD_ERROR_MESSAGE =
  "Saved Google OAuth setup could not be loaded. Enter a client ID to continue for this session.";
const STORAGE_SAVE_ERROR_MESSAGE =
  "The Google OAuth client ID could not be saved. You can continue for this session.";
const STORAGE_CLEAR_ERROR_MESSAGE =
  "The saved Google OAuth client ID could not be cleared. The existing ID remains available.";

export function useGoogleSheetsImport(options: Readonly<{
  configuration?: GoogleSheetsServiceConfiguration;
  deprecatedTokenProviderFactory?: (clientId: string) => TokenProvider;
  origin: string;
  onImported(result: GoogleSheetsImportResult): void | Promise<void>;
  onError?(error: GoogleSheetsError): void;
}>): GoogleSheetsImportController {
  const { configuration, deprecatedTokenProviderFactory, origin } = options;
  const [open, setOpen] = useState(false);
  const [storageStatus, setStorageStatus] = useState<"loading" | "loaded">(
    isClientIdStorage(configuration?.clientIdStorage) ? "loading" : "loaded"
  );
  const [editableClientId, setEditableClientIdState] = useState<EditableGoogleClientId | null>(null);
  const [clientIdDraft, setClientIdDraftState] = useState("");
  const [sheetDraft, setSheetDraftState] = useState("");
  const [editingClientId, setEditingClientId] = useState(false);
  const [readiness, setReadinessState] = useState<ReadinessPhase>("loading");
  const [readinessError, setReadinessError] = useState<GoogleSheetsError | undefined>();
  const [readinessRetryable, setReadinessRetryableState] = useState(false);
  const [warning, setWarning] = useState<GoogleSheetsError | undefined>();
  const [operation, setOperation] = useState<ImportOperation | null>(null);
  const [preparationAttempt, setPreparationAttempt] = useState(0);
  const [storagePendingCount, setStoragePendingCount] = useState(0);
  const [storageAction, setStorageAction] = useState<"save" | "forget" | undefined>();

  const readinessRef = useRef<ReadinessPhase>("loading");
  const readinessRetryableRef = useRef(false);
  const editableClientIdRef = useRef<EditableGoogleClientId | null>(null);
  const providerRef = useRef<TokenProvider | null>(null);
  const providerRecordsRef = useRef<ProviderRecord[]>([]);
  const storageLoadRef = useRef<StorageLoad | null>(null);
  const storageQueueRef = useRef(createSerialOperationQueue());
  const storageVersionRef = useRef(0);
  const storagePendingRef = useRef(0);
  const clientIdTouchedRef = useRef(false);
  const attemptGenerationRef = useRef(0);
  const importPendingRef = useRef(false);
  const onImportedRef = useRef(options.onImported);
  const onErrorRef = useRef(options.onError);
  onImportedRef.current = options.onImported;
  onErrorRef.current = options.onError;

  const originAssessment = useMemo(() => assessGoogleOAuthOrigin(origin), [origin]);
  const managedClientId = configuration?.clientId?.trim() ?? "";
  const authSource = resolveGoogleAuthSource(
    configuration,
    editableClientId,
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
  const displayedClientId = clientIdSource === "managed" ? managedClientId : clientIdDraft;

  const updateReadiness = useCallback((next: ReadinessPhase) => {
    readinessRef.current = next;
    setReadinessState(next);
  }, []);

  const updateReadinessRetryable = useCallback((next: boolean) => {
    readinessRetryableRef.current = next;
    setReadinessRetryableState(next);
  }, []);

  const updateEditableClientId = useCallback((next: EditableGoogleClientId | null) => {
    editableClientIdRef.current = next;
    setEditableClientIdState(next);
  }, []);

  const notifyError = useCallback((error: GoogleSheetsError) => {
    try {
      onErrorRef.current?.(error);
    } catch {
      // Host observability must never replace the dialog's safe inline error.
    }
  }, []);

  const beginStorageOperation = useCallback((action: "save" | "forget") => {
    storagePendingRef.current += 1;
    setStoragePendingCount(storagePendingRef.current);
    setStorageAction(action);
  }, []);

  const finishStorageOperation = useCallback(() => {
    storagePendingRef.current = Math.max(0, storagePendingRef.current - 1);
    setStoragePendingCount(storagePendingRef.current);
    if (storagePendingRef.current === 0) {
      setStorageAction(undefined);
    }
  }, []);

  useEffect(() => {
    const storage = configuration?.clientIdStorage;
    if (!isClientIdStorage(storage)) {
      setStorageStatus("loaded");
      return undefined;
    }

    setStorageStatus("loading");
    if (storageLoadRef.current?.storage !== storage) {
      try {
        storageLoadRef.current = { storage, result: storage.load() };
      } catch (error) {
        storageLoadRef.current = { storage, result: Promise.reject(error) };
      }
    }

    const applyStoredValue = (storedValue: string | null) => {
      setStorageStatus("loaded");
      if (clientIdTouchedRef.current || storedValue === null) {
        return;
      }
      const validation = validateGoogleClientId(storedValue);
      if (!validation.valid) {
        if (!configuration?.tokenProvider && !managedClientId) {
          setReadinessError(
            new GoogleSheetsError("invalid_client", validation.message, true)
          );
        }
        return;
      }
      updateEditableClientId({ value: validation.value, source: "stored" });
      setClientIdDraftState(validation.value);
    };
    const applyStorageFailure = () => {
      setStorageStatus("loaded");
      if (configuration?.tokenProvider || managedClientId) {
        return;
      }
      const error = storageError(STORAGE_LOAD_ERROR_MESSAGE);
      setWarning(error);
      updateReadiness("setup");
      notifyError(error);
    };
    const loadResult = storageLoadRef.current.result;
    if (!isPromiseLike(loadResult)) {
      applyStoredValue(loadResult);
      return undefined;
    }

    let active = true;
    loadResult.then(
      (storedValue) => {
        if (active) {
          applyStoredValue(storedValue);
        }
      },
      () => {
        if (active) {
          applyStorageFailure();
        }
      }
    );

    return () => {
      active = false;
    };
  }, [
    configuration?.clientIdStorage,
    configuration?.tokenProvider,
    managedClientId,
    notifyError,
    updateEditableClientId,
    updateReadiness
  ]);

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
      updateReadiness(storageStatus === "loading" ? "loading" : "setup");
      return () => {
        active = false;
      };
    }

    if (authSource.kind === "client" && editingClientId && authSource.source !== "managed") {
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

    let record: ProviderRecord;
    try {
      record = getProviderRecord(providerRecordsRef.current, authSource);
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

    providerRef.current = record.provider;
    setReadinessError(undefined);
    updateReadinessRetryable(false);
    updateReadiness("preparing");
    prepareProvider(record).then(
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
    editingClientId,
    notifyError,
    originAssessment,
    preparationAttempt,
    storageStatus,
    updateReadiness,
    updateReadinessRetryable,
    usesBuiltInGoogleAuth
  ]);

  const setClientIdDraft = useCallback((value: string) => {
    clientIdTouchedRef.current = true;
    setClientIdDraftState(value);
    setReadinessError((current) => current?.code === "invalid_client" ? undefined : current);
  }, []);

  const setSheetDraft = useCallback((value: string) => {
    setSheetDraftState(value);
    setOperation((current) => current?.error?.code === "invalid_sheet_url" ? null : current);
  }, []);

  const openDialog = useCallback(() => {
    setOpen(true);
  }, []);

  const closeDialog = useCallback(() => {
    attemptGenerationRef.current += 1;
    importPendingRef.current = false;
    setOperation(null);
    setOpen(false);
  }, []);

  const saveClientId = useCallback(async () => {
    const validation = validateGoogleClientId(clientIdDraft);
    if (!validation.valid) {
      const error = new GoogleSheetsError("invalid_client", validation.message, true);
      setReadinessError(error);
      updateReadinessRetryable(false);
      setOperation(null);
      setEditingClientId(true);
      updateReadiness("setup");
      return;
    }

    clientIdTouchedRef.current = true;
    setClientIdDraftState(validation.value);
    updateEditableClientId({ value: validation.value, source: "session" });
    setEditingClientId(false);
    setReadinessError(undefined);
    updateReadinessRetryable(false);
    setOperation(null);

    const storage = configuration?.clientIdStorage;
    if (!isClientIdStorage(storage)) {
      return;
    }

    const version = ++storageVersionRef.current;
    beginStorageOperation("save");
    try {
      await storageQueueRef.current.enqueue(() => storage.save(validation.value));
      if (version === storageVersionRef.current) {
        updateEditableClientId({ value: validation.value, source: "stored" });
        setWarning(undefined);
      }
    } catch {
      if (version === storageVersionRef.current) {
        updateEditableClientId({ value: validation.value, source: "session" });
        const error = storageError(STORAGE_SAVE_ERROR_MESSAGE);
        setWarning(error);
        notifyError(error);
      }
    } finally {
      finishStorageOperation();
    }
  }, [
    beginStorageOperation,
    clientIdDraft,
    configuration?.clientIdStorage,
    finishStorageOperation,
    notifyError,
    updateEditableClientId,
    updateReadiness,
    updateReadinessRetryable
  ]);

  const forgetClientId = useCallback(async () => {
    clientIdTouchedRef.current = true;
    const storage = configuration?.clientIdStorage;
    if (!isClientIdStorage(storage)) {
      clearEditableClientId();
      return;
    }

    const version = ++storageVersionRef.current;
    beginStorageOperation("forget");
    try {
      await storageQueueRef.current.enqueue(() => storage.clear());
      if (version === storageVersionRef.current) {
        clearEditableClientId();
        setWarning(undefined);
      }
    } catch {
      if (version === storageVersionRef.current) {
        const error = storageError(STORAGE_CLEAR_ERROR_MESSAGE);
        setWarning(error);
        notifyError(error);
      }
    } finally {
      finishStorageOperation();
    }

    function clearEditableClientId() {
      providerRecordsRef.current = [];
      providerRef.current = null;
      updateEditableClientId(null);
      setClientIdDraftState("");
      setEditingClientId(true);
      setReadinessError(undefined);
      updateReadinessRetryable(false);
      setOperation(null);
      updateReadiness("setup");
    }
  }, [
    beginStorageOperation,
    configuration?.clientIdStorage,
    finishStorageOperation,
    notifyError,
    updateEditableClientId,
    updateReadiness,
    updateReadinessRetryable
  ]);

  const changeClientId = useCallback(() => {
    if (storagePendingRef.current > 0) {
      return;
    }
    attemptGenerationRef.current += 1;
    importPendingRef.current = false;
    clientIdTouchedRef.current = true;
    providerRecordsRef.current = [];
    providerRef.current = null;
    updateEditableClientId(editableClientIdRef.current
      ? { value: editableClientIdRef.current.value, source: "session" }
      : null);
    setEditingClientId(true);
    setReadinessError(undefined);
    updateReadinessRetryable(false);
    setWarning(undefined);
    setOperation(null);
    updateReadiness("setup");
  }, [updateEditableClientId, updateReadiness, updateReadinessRetryable]);

  const importSheet = useCallback(() => {
    if (importPendingRef.current || storagePendingRef.current > 0) {
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
          await onImportedRef.current(imported);
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
  }, [notifyError, sheetDraft]);

  const retry = useCallback(() => {
    if (storagePendingRef.current > 0) {
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
  }, [importSheet, operation, updateReadiness, updateReadinessRetryable]);

  const storageBusy = storagePendingCount > 0;
  const phase: GoogleSheetsImportPhase = !open
    ? "closed"
    : storageBusy
      ? "loading"
      : operation?.phase ?? readiness;
  const error = operation?.error ?? readinessError;
  const clientIdEditable =
    readiness === "setup" &&
    authSource.kind !== "provider" &&
    clientIdSource !== "managed";
  const showSheetInput = readiness === "ready" || operation !== null;
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
    ...(storageAction ? { storageAction } : {}),
    canSaveClientId: clientIdEditable && !storageBusy,
    canChangeClientId: clientIdSource === "stored" && !storageBusy,
    canForgetClientId: clientIdSource === "stored" && !storageBusy,
    canImport:
      readiness === "ready" && operation === null && !storageBusy && !importPendingRef.current,
    canRetry,
    ...(error ? { error } : {}),
    ...(warning ? { warning } : {}),
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

function getProviderRecord(
  records: ProviderRecord[],
  source: Exclude<ReturnType<typeof resolveGoogleAuthSource>, { kind: "missing" }>
): ProviderRecord {
  const kind = source.kind === "provider" ? "direct" : "factory";
  const identity = source.kind === "provider" ? source.provider : source.factory;
  const clientId = source.kind === "provider" ? "" : source.clientId;
  const cached = records.find(
    (record) =>
      record.kind === kind &&
      record.identity === identity &&
      record.clientId === clientId
  );
  if (cached) {
    return cached;
  }

  const provider = source.kind === "provider"
    ? source.provider
    : source.factory(source.clientId);
  if (!provider || typeof provider.getAccessToken !== "function") {
    throw new Error("Invalid Google token provider.");
  }
  const record: ProviderRecord = {
    kind,
    identity,
    clientId,
    provider,
    prepared: false
  };
  records.push(record);
  return record;
}

function prepareProvider(record: ProviderRecord): Promise<void> {
  if (record.prepared) {
    return Promise.resolve();
  }
  if (record.preparation) {
    return record.preparation;
  }
  if (!record.provider.prepare) {
    record.prepared = true;
    return Promise.resolve();
  }

  try {
    record.preparation = Promise.resolve(record.provider.prepare()).then(
      () => {
        record.prepared = true;
        record.preparation = undefined;
      },
      (error: unknown) => {
        record.preparation = undefined;
        throw error;
      }
    );
  } catch (error) {
    return Promise.reject(error);
  }
  return record.preparation;
}

function isClientIdStorage(
  storage: GoogleClientIdStorage | false | undefined
): storage is GoogleClientIdStorage {
  return Boolean(storage);
}

function isPromiseLike<T>(value: T | Promise<T>): value is Promise<T> {
  return typeof (value as Promise<T>)?.then === "function";
}

function storageError(message: string): GoogleSheetsError {
  return new GoogleSheetsError("unknown", message, true);
}

function incompatibleOriginError(
  assessment: Extract<GoogleOAuthOriginAssessment, { status: "blocked" }>
): GoogleSheetsError {
  const message = assessment.reason === "ip_literal"
    ? "Google browser OAuth requires an authorized HTTPS DNS origin and cannot use a raw LAN IP."
    : "Google browser OAuth requires an eligible, authorized HTTPS origin.";
  return new GoogleSheetsError("incompatible_origin", message, true);
}
