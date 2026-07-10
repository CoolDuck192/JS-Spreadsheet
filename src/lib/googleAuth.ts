/**
 * Pluggable auth for Google APIs.
 *
 * Hosts can inject their own token source. Standalone browser clients can use
 * Google Identity Services without a client secret; access tokens remain in
 * memory and are requested on demand.
 */

import type { TokenProvider } from "../core/workbook/services";
import { validateGoogleClientId } from "./googleConfiguration";
import { GoogleSheetsError } from "./googleErrors";

export type { TokenProvider };

export type BrowserTokenProvider = Omit<TokenProvider, "prepare"> & {
  prepare(): Promise<void>;
};

export const SHEETS_READONLY_SCOPE =
  "https://www.googleapis.com/auth/spreadsheets.readonly";
export const SHEETS_READWRITE_SCOPE = "https://www.googleapis.com/auth/spreadsheets";

type GoogleTokenClient = {
  requestAccessToken: (overrides?: { prompt?: string }) => void;
};

type GoogleIdentityServices = {
  accounts: {
    oauth2: {
      initTokenClient: (config: {
        client_id: string;
        scope: string;
        callback: (response: { access_token?: string; error?: string }) => void;
        error_callback?: (error: { type?: string; message?: string }) => void;
      }) => GoogleTokenClient;
    };
  };
};

const GIS_SCRIPT_URL = "https://accounts.google.com/gsi/client";
const TOKEN_EXPIRY_SAFETY_MS = 5 * 60 * 1000;
const ASSUMED_TOKEN_LIFETIME_MS = 55 * 60 * 1000;

declare global {
  interface Window {
    google?: GoogleIdentityServices;
  }
}

/**
 * Standalone-browser token provider backed by Google Identity Services.
 * Call prepare before a user gesture so the first token request can open its
 * popup synchronously from that gesture.
 */
export function createBrowserTokenProvider(clientId: string): BrowserTokenProvider {
  let cached: { token: string; scopeKey: string; expiresAt: number } | null = null;
  let preparedGoogle: GoogleIdentityServices | null = null;
  let normalizedClientId: string | null = null;
  let preparation: Promise<void> | null = null;

  const provider: BrowserTokenProvider = {
    prepare() {
      if (preparedGoogle) {
        return Promise.resolve();
      }
      const validation = validateGoogleClientId(clientId);
      if (!validation.valid) {
        return Promise.reject(
          new GoogleSheetsError("invalid_client", validation.message, true)
        );
      }
      normalizedClientId = validation.value;
      preparation ??= loadGoogleIdentityServices()
        .then((google) => {
          preparedGoogle = google;
        })
        .catch((error: unknown) => {
          preparation = null;
          if (error instanceof GoogleSheetsError) {
            throw error;
          }
          throw googleIdentityLoadError();
        });
      return preparation;
    },

    getAccessToken(scopes) {
      const scopeKey = [...scopes].sort().join(" ");
      if (
        cached &&
        cached.scopeKey === scopeKey &&
        cached.expiresAt > Date.now() + TOKEN_EXPIRY_SAFETY_MS
      ) {
        return Promise.resolve(cached.token);
      }

      if (preparedGoogle && normalizedClientId) {
        return requestToken(preparedGoogle, normalizedClientId, scopeKey).then((token) => {
          cached = {
            token,
            scopeKey,
            expiresAt: Date.now() + ASSUMED_TOKEN_LIFETIME_MS
          };
          return token;
        });
      }

      return provider.prepare().then(() => provider.getAccessToken(scopes));
    }
  };

  return provider;
}

function requestToken(
  google: GoogleIdentityServices,
  clientId: string,
  scopeKey: string
): Promise<string> {
  return new Promise<string>((resolve, reject) => {
    try {
      const client = google.accounts.oauth2.initTokenClient({
        client_id: clientId,
        scope: scopeKey,
        callback: (response) => {
          if (typeof response.access_token === "string" && response.access_token) {
            resolve(response.access_token);
            return;
          }
          reject(oauthResponseError(response.error));
        },
        error_callback: (error) => {
          reject(oauthPopupError(error.type));
        }
      });
      if (!client || typeof client.requestAccessToken !== "function") {
        reject(googleIdentityLoadError());
        return;
      }
      client.requestAccessToken();
    } catch {
      reject(googleIdentityLoadError());
    }
  });
}

function oauthResponseError(error: string | undefined): GoogleSheetsError {
  if (error === "access_denied") {
    return new GoogleSheetsError(
      "access_denied",
      "Google Sheets access was not granted. Try again and approve access.",
      true
    );
  }
  if (error === "invalid_client") {
    return new GoogleSheetsError(
      "invalid_client",
      "The Google OAuth client ID is invalid. Use a Web application client ID.",
      true
    );
  }
  if (error === "origin_mismatch") {
    return new GoogleSheetsError(
      "origin_mismatch",
      "This site origin is not authorized for the Google OAuth client.",
      true
    );
  }
  return new GoogleSheetsError(
    "unknown",
    "Google sign-in failed. Check the configuration and try again.",
    true
  );
}

function oauthPopupError(type: string | undefined): GoogleSheetsError {
  if (type === "popup_failed_to_open") {
    return new GoogleSheetsError(
      "popup_blocked",
      "Google sign-in could not open. Allow pop-ups and try again.",
      true
    );
  }
  if (type === "popup_closed") {
    return new GoogleSheetsError(
      "popup_closed",
      "Google sign-in was closed before access was granted.",
      true
    );
  }
  return new GoogleSheetsError(
    "unknown",
    "Google sign-in failed. Check the configuration and try again.",
    true
  );
}

let gisLoadPromise: Promise<GoogleIdentityServices> | null = null;

function loadGoogleIdentityServices(): Promise<GoogleIdentityServices> {
  const available = getAvailableGoogleIdentityServices();
  if (available) {
    return Promise.resolve(available);
  }
  if (typeof document === "undefined") {
    return Promise.reject(googleIdentityLoadError());
  }

  gisLoadPromise ??= new Promise((resolve, reject) => {
    const script = document.createElement("script");
    script.src = GIS_SCRIPT_URL;
    script.async = true;
    script.onload = () => {
      const loaded = getAvailableGoogleIdentityServices();
      if (loaded) {
        resolve(loaded);
        return;
      }
      script.remove();
      gisLoadPromise = null;
      reject(googleIdentityLoadError());
    };
    script.onerror = () => {
      script.remove();
      gisLoadPromise = null;
      reject(googleIdentityLoadError());
    };
    document.head.appendChild(script);
  });

  return gisLoadPromise;
}

function getAvailableGoogleIdentityServices(): GoogleIdentityServices | null {
  if (
    typeof window !== "undefined" &&
    typeof window.google?.accounts?.oauth2?.initTokenClient === "function"
  ) {
    return window.google;
  }
  return null;
}

function googleIdentityLoadError(): GoogleSheetsError {
  return new GoogleSheetsError(
    "gis_load_failed",
    "Google sign-in could not be prepared. Check your connection and try again.",
    true
  );
}
