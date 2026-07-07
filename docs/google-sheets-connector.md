# Google Sheets Connector

Link any user's Google Sheet into the spreadsheet: **Data → Link Google Sheet**, paste the sheet's URL (or bare spreadsheet id), sign in with Google, and the workbook imports with formulas intact (`valueRenderOption=FORMULA`), including hidden sheets.

## One-time setup (standalone app)

The app is a static SPA — there is no backend and no client secret. Auth uses Google Identity Services' token flow, which needs only an OAuth client id:

1. In [Google Cloud Console](https://console.cloud.google.com/): create (or pick) a project and enable the **Google Sheets API**.
2. Create **Credentials → OAuth client ID → Web application**, and add your app's origin(s) to **Authorized JavaScript origins** — e.g. `http://127.0.0.1:5173` for local dev and your deployed origin for production.
3. Put the client id in `.env.local` at the repo root:

   ```bash
   VITE_GOOGLE_CLIENT_ID=1234567890-abc.apps.googleusercontent.com
   ```

4. Restart the dev server. The **Link Google Sheet** button now opens Google sign-in; each user authorizes with their own account and can only access sheets they can read.

Without the env var the button explains what's missing in the status bar instead of failing silently.

## How it works

- `src/lib/googleAuth.ts` — a `TokenProvider` interface plus `createBrowserTokenProvider(clientId)`, the Google Identity Services implementation (in-browser token client, ~1h access tokens, cached and re-requested on demand; popup-closed/blocked surfaces as a rejected promise via `error_callback`).
- `src/lib/googleSheets.ts` — `parseSpreadsheetId` (URL or bare id) and `importWorkbookFromGoogleSheets(idOrUrl, tokenProvider)`: fetches spreadsheet metadata, batch-gets all sheets' values with formulas, and builds a `WorkbookModel`. Friendly errors for 401/403 (access denied), 404 (not found), and 429 (rate limit).

## Embedding note

Host applications should not depend on the built-in browser auth: pass your own `TokenProvider` (backed by your existing Google OAuth session or a backend token mint) to `importWorkbookFromGoogleSheets`. Both the function and the interface are exported from the library entry point (`src/index.ts`).

## Current scope and roadmap

Read-only import today. Planned (see the roadmap in the audit report): write-back via `values.batchUpdate`, incremental refresh, and a link-management panel showing sync state.
