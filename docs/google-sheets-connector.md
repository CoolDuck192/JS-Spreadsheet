# Google Sheets Import

Choose **File → Import Google Sheet**, enter a Google Sheets URL or spreadsheet
ID, and authorize read-only access. The import fetches formulas and hidden
sheets, then replaces the current workbook through normal workbook history.
It is a one-time import: there is no continuous link, refresh job, write-back,
or synchronization state.

## Standalone setup

The standalone app is a static browser application. It uses Google Identity
Services and never needs a client secret or backend.

1. In [Google Cloud Console](https://console.cloud.google.com/), create or select
   a project and enable the **Google Sheets API**.
2. Configure the OAuth consent screen and any required test users.
3. Create an **OAuth client ID** for a **Web application**.
4. Add the app's exact eligible HTTPS DNS origin to **Authorized JavaScript
   origins**.
5. In the spreadsheet, choose **File → Import Google Sheet** and paste the
   public client ID into the setup dialog.

A fresh clone opens this setup dialog when configuration is absent. The dialog
always shows the exact current origin. Google browser OAuth requires an
eligible, exactly registered origin; a raw LAN IP is intentionally diagnosed as
incompatible. Expose the app through an HTTPS DNS name and register that exact
origin, or have the application host provide a token provider.

The standalone dialog stores only the public client ID in browser local storage
under `javascript-spreadsheet.google-client-id.v1`. OAuth access tokens remain
in memory. The app never stores a client secret. **Change client ID** and
**Forget client ID** affect only this Google configuration, not workbook
storage.

## Embedded applications

Configure the public `Spreadsheet` through `services.googleSheets`:

```tsx
import { Spreadsheet } from "js-spreadsheet/react";
import type {
  GoogleClientIdStorage,
  TokenProvider
} from "js-spreadsheet/core";

const tokenProvider: TokenProvider = {
  async getAccessToken(scopes) {
    return hostGoogleSession.getAccessToken(scopes);
  }
};

const clientIdStorage: GoogleClientIdStorage = {
  load: () => hostSettings.load("google-client-id"),
  save: (clientId) => hostSettings.save("google-client-id", clientId),
  clear: () => hostSettings.remove("google-client-id")
};

export function Workbook() {
  return (
    <Spreadsheet
      services={{
        googleSheets: {
          tokenProvider,
          // Or use clientId plus tokenProviderFactory for host-created providers.
          // clientId: "1234567890-abc.apps.googleusercontent.com",
          // tokenProviderFactory: (clientId) => createHostProvider(clientId),
          // clientIdStorage
        }
      }}
    />
  );
}
```

A supplied `tokenProvider` takes precedence and can bypass built-in browser
OAuth origin setup. Alternatively, pass a managed `clientId`; optionally pair
it with `tokenProviderFactory`. `clientIdStorage` is opt-in for embedded
instances, so the reusable `Spreadsheet` persists no Google client ID unless
the host explicitly supplies that adapter. Set
`features={{ googleSheets: false }}` when the host does not offer Google import.

## Browser policy and iframes

Hosts using the built-in Google browser provider must allow these origins in
their Content Security Policy:

- `script-src https://accounts.google.com` for Google Identity Services.
- `connect-src https://accounts.google.com https://sheets.googleapis.com` for
  authentication and read-only Sheets API requests.
- `frame-src https://accounts.google.com` for Google sign-in UI.

If the spreadsheet is inside a sandboxed iframe, the iframe must permit popup
sign-in (for example, `sandbox="allow-scripts allow-same-origin allow-popups"`).
The top-level host's popup policy must also allow Google sign-in. A host-supplied
token provider can use the host's existing authentication instead.

## Security and failure behavior

The connector requests
`https://www.googleapis.com/auth/spreadsheets.readonly`. It reports blocked or
closed popups, origin mismatch, denied access, disabled API, missing sheets,
rate limiting, and network failures inside the import dialog. Canceling an
in-flight import invalidates late results, and a failed or canceled operation
does not replace the workbook. Error output is sanitized so bearer tokens and
raw provider responses are not rendered.

The implementation is exported from `js-spreadsheet/connectors/google`:
`createBrowserTokenProvider`, `parseSpreadsheetId`, and
`importWorkbookFromGoogleSheets` are available for custom integrations.
