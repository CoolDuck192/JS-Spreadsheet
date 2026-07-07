/**
 * Pluggable auth for Google APIs.
 *
 * The spreadsheet is designed to be embedded in a host app, so auth is an
 * interface: the host can supply its own token source (its existing Google
 * OAuth session, a backend-minted token, etc.). For standalone use,
 * `createBrowserTokenProvider` implements the interface with Google Identity
 * Services' token client — the correct flow for a static SPA with no backend
 * (no client secret, ~1h access tokens requested on demand).
 */

export type TokenProvider = {
  /** Resolve an OAuth2 access token bearing the given scopes. */
  getAccessToken: (scopes: readonly string[]) => Promise<string>;
};

export const SHEETS_READONLY_SCOPE = "https://www.googleapis.com/auth/spreadsheets.readonly";
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
 * Standalone-browser TokenProvider backed by Google Identity Services.
 * Requires a Google Cloud OAuth client ID (Web application type) whose
 * authorized JavaScript origins include the page's origin.
 */
export function createBrowserTokenProvider(clientId: string): TokenProvider {
  let cached: { token: string; scopeKey: string; expiresAt: number } | null = null;

  return {
    async getAccessToken(scopes) {
      const scopeKey = [...scopes].sort().join(" ");
      if (cached && cached.scopeKey === scopeKey && cached.expiresAt > Date.now() + TOKEN_EXPIRY_SAFETY_MS) {
        return cached.token;
      }

      const google = await loadGoogleIdentityServices();
      const token = await new Promise<string>((resolve, reject) => {
        const client = google.accounts.oauth2.initTokenClient({
          client_id: clientId,
          scope: scopeKey,
          callback: (response) => {
            if (response.access_token) {
              resolve(response.access_token);
            } else {
              reject(new Error(response.error ?? "Google sign-in was cancelled"));
            }
          },
          // GIS reports non-OAuth failures (popup closed, popup blocked) here, not
          // in callback — without it a dismissed popup leaves the promise pending.
          error_callback: (error) => {
            reject(new Error(error.message ?? error.type ?? "Google sign-in was cancelled"));
          }
        });
        client.requestAccessToken();
      });

      cached = { token, scopeKey, expiresAt: Date.now() + ASSUMED_TOKEN_LIFETIME_MS };
      return token;
    }
  };
}

let gisLoadPromise: Promise<GoogleIdentityServices> | null = null;

function loadGoogleIdentityServices(): Promise<GoogleIdentityServices> {
  if (window.google?.accounts?.oauth2) {
    return Promise.resolve(window.google);
  }

  gisLoadPromise ??= new Promise((resolve, reject) => {
    const script = document.createElement("script");
    script.src = GIS_SCRIPT_URL;
    script.async = true;
    script.onload = () => {
      if (window.google?.accounts?.oauth2) {
        resolve(window.google);
      } else {
        // Clear the cached promise so a later call can retry, matching onerror.
        gisLoadPromise = null;
        reject(new Error("Google Identity Services failed to initialize"));
      }
    };
    script.onerror = () => {
      gisLoadPromise = null;
      reject(new Error("Could not load Google Identity Services"));
    };
    document.head.appendChild(script);
  });

  return gisLoadPromise;
}
