import type { GoogleClientIdStorage } from "../core/workbook/services";
import { validateGoogleClientId } from "../lib/googleConfiguration";
import { GoogleSheetsError } from "../lib/googleErrors";

export const GOOGLE_CLIENT_ID_STORAGE_KEY =
  "javascript-spreadsheet.google-client-id.v1";

type ClientIdStorage = Pick<Storage, "getItem" | "setItem" | "removeItem">;

export function createBrowserGoogleClientIdStorage(
  storage: ClientIdStorage
): GoogleClientIdStorage {
  return {
    async load() {
      const stored = storage.getItem(GOOGLE_CLIENT_ID_STORAGE_KEY);
      if (stored === null) {
        return null;
      }
      const validation = validateGoogleClientId(stored);
      return validation.valid ? validation.value : null;
    },

    async save(clientId) {
      const validation = validateGoogleClientId(clientId);
      if (!validation.valid) {
        throw new GoogleSheetsError("invalid_client", validation.message, true);
      }
      storage.setItem(GOOGLE_CLIENT_ID_STORAGE_KEY, validation.value);
    },

    async clear() {
      storage.removeItem(GOOGLE_CLIENT_ID_STORAGE_KEY);
    }
  };
}

let defaultStorage: GoogleClientIdStorage | null = null;

/** Standalone-only adapter. Browser storage is resolved lazily on each operation. */
export function getDefaultBrowserGoogleClientIdStorage(): GoogleClientIdStorage {
  defaultStorage ??= {
    load() {
      const stored = getLocalStorage().getItem(GOOGLE_CLIENT_ID_STORAGE_KEY);
      if (stored === null) {
        return null;
      }
      const validation = validateGoogleClientId(stored);
      return validation.valid ? validation.value : null;
    },
    save(clientId) {
      const validation = validateGoogleClientId(clientId);
      if (!validation.valid) {
        throw new GoogleSheetsError("invalid_client", validation.message, true);
      }
      getLocalStorage().setItem(GOOGLE_CLIENT_ID_STORAGE_KEY, validation.value);
    },
    clear() {
      getLocalStorage().removeItem(GOOGLE_CLIENT_ID_STORAGE_KEY);
    }
  };
  return defaultStorage;
}

function getLocalStorage(): Storage {
  if (typeof window === "undefined") {
    throw new Error("Browser storage is unavailable.");
  }
  return window.localStorage;
}
