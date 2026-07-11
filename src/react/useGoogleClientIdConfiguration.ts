import { useCallback, useEffect, useRef, useState } from "react";
import type { GoogleClientIdStorage } from "../core/workbook/services";
import {
  validateGoogleClientId,
  type EditableGoogleClientId
} from "../lib/googleConfiguration";
import { GoogleSheetsError } from "../lib/googleErrors";
import { createSerialOperationQueue } from "./serialOperationQueue";

export type GoogleClientIdForgetOutcome = Readonly<{
  status: "cleared" | "retained";
}>;

export type GoogleClientIdConfiguration = Readonly<{
  storageStatus: "loading" | "loaded";
  editableClientId: EditableGoogleClientId | null;
  clientIdDraft: string;
  editingClientId: boolean;
  storageBusy: boolean;
  storageAction?: "save" | "forget";
  clientIdError?: GoogleSheetsError;
  warning?: GoogleSheetsError;
  setClientIdDraft(value: string): void;
  saveClientId(normalizedClientId: string): Promise<void>;
  forgetClientId(): Promise<GoogleClientIdForgetOutcome>;
  changeClientId(): void;
  isStorageBusy(): boolean;
}>;

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

export function useGoogleClientIdConfiguration(options: Readonly<{
  storage: GoogleClientIdStorage | false | undefined;
  ignoreLoadFailure: boolean;
  onError(error: GoogleSheetsError): void;
}>): GoogleClientIdConfiguration {
  const { storage, ignoreLoadFailure } = options;
  const [storageStatus, setStorageStatus] = useState<"loading" | "loaded">(
    isClientIdStorage(storage) ? "loading" : "loaded"
  );
  const [editableClientId, setEditableClientIdState] = useState<EditableGoogleClientId | null>(null);
  const [clientIdDraft, setClientIdDraftState] = useState("");
  const [editingClientId, setEditingClientId] = useState(false);
  const [storagePendingCount, setStoragePendingCount] = useState(0);
  const [storageAction, setStorageAction] = useState<"save" | "forget" | undefined>();
  const [clientIdError, setClientIdError] = useState<GoogleSheetsError | undefined>();
  const [warning, setWarning] = useState<GoogleSheetsError | undefined>();
  const [storageQueue] = useState(createSerialOperationQueue);

  const editableClientIdRef = useRef<EditableGoogleClientId | null>(null);
  const storageLoadRef = useRef<StorageLoad | null>(null);
  const storageVersionRef = useRef(0);
  const storagePendingRef = useRef(0);
  const clientIdTouchedRef = useRef(false);
  const onErrorRef = useRef(options.onError);
  onErrorRef.current = options.onError;

  const updateEditableClientId = useCallback((next: EditableGoogleClientId | null) => {
    editableClientIdRef.current = next;
    setEditableClientIdState(next);
  }, []);

  const notifyError = useCallback((error: GoogleSheetsError) => {
    onErrorRef.current(error);
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
        if (!ignoreLoadFailure) {
          setClientIdError(
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
      if (ignoreLoadFailure) {
        return;
      }
      const error = storageError(STORAGE_LOAD_ERROR_MESSAGE);
      setWarning(error);
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
  }, [ignoreLoadFailure, notifyError, storage, updateEditableClientId]);

  const setClientIdDraft = useCallback((value: string) => {
    clientIdTouchedRef.current = true;
    setClientIdDraftState(value);
    setClientIdError(undefined);
  }, []);

  const saveClientId = useCallback(async (normalizedClientId: string) => {
    clientIdTouchedRef.current = true;
    setClientIdDraftState(normalizedClientId);
    updateEditableClientId({ value: normalizedClientId, source: "session" });
    setEditingClientId(false);
    setClientIdError(undefined);

    if (!isClientIdStorage(storage)) {
      return;
    }

    const version = ++storageVersionRef.current;
    beginStorageOperation("save");
    try {
      await storageQueue.enqueue(() => storage.save(normalizedClientId));
      if (version === storageVersionRef.current) {
        updateEditableClientId({ value: normalizedClientId, source: "stored" });
        setWarning(undefined);
      }
    } catch {
      if (version === storageVersionRef.current) {
        updateEditableClientId({ value: normalizedClientId, source: "session" });
        const error = storageError(STORAGE_SAVE_ERROR_MESSAGE);
        setWarning(error);
        notifyError(error);
      }
    } finally {
      finishStorageOperation();
    }
  }, [
    beginStorageOperation,
    finishStorageOperation,
    notifyError,
    storage,
    storageQueue,
    updateEditableClientId
  ]);

  const forgetClientId = useCallback(async (): Promise<GoogleClientIdForgetOutcome> => {
    clientIdTouchedRef.current = true;
    if (!isClientIdStorage(storage)) {
      clearEditableClientId();
      return { status: "cleared" };
    }

    const version = ++storageVersionRef.current;
    beginStorageOperation("forget");
    try {
      await storageQueue.enqueue(() => storage.clear());
      if (version !== storageVersionRef.current) {
        return { status: "retained" };
      }
      clearEditableClientId();
      setWarning(undefined);
      return { status: "cleared" };
    } catch {
      if (version === storageVersionRef.current) {
        const error = storageError(STORAGE_CLEAR_ERROR_MESSAGE);
        setWarning(error);
        notifyError(error);
      }
      return { status: "retained" };
    } finally {
      finishStorageOperation();
    }

    function clearEditableClientId() {
      updateEditableClientId(null);
      setClientIdDraftState("");
      setEditingClientId(true);
      setClientIdError(undefined);
    }
  }, [
    beginStorageOperation,
    finishStorageOperation,
    notifyError,
    storage,
    storageQueue,
    updateEditableClientId
  ]);

  const changeClientId = useCallback(() => {
    updateEditableClientId(editableClientIdRef.current
      ? { value: editableClientIdRef.current.value, source: "session" }
      : null);
    setEditingClientId(true);
    setClientIdError(undefined);
    setWarning(undefined);
  }, [updateEditableClientId]);

  const isStorageBusy = useCallback(() => storagePendingRef.current > 0, []);
  const storageBusy = storagePendingCount > 0;

  return {
    storageStatus,
    editableClientId,
    clientIdDraft,
    editingClientId,
    storageBusy,
    ...(storageAction ? { storageAction } : {}),
    ...(clientIdError ? { clientIdError } : {}),
    ...(warning ? { warning } : {}),
    setClientIdDraft,
    saveClientId,
    forgetClientId,
    changeClientId,
    isStorageBusy
  };
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
