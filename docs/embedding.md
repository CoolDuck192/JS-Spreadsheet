# Embedding JS-Spreadsheet in Your App

`src/index.ts` is the library entry point. The default export is the full spreadsheet UI; named exports are the headless functions.

```tsx
import Spreadsheet from "javascript-spreadsheet-clone";        // full UI component
import {
  createBlankWorkbook,
  setCellContent,
  getCellContent,
  createFormulaEngine,
  importWorkbookFromXlsx,
  exportWorkbookToXlsx,
  createPivotTable,
  importWorkbookFromGoogleSheets,
  type WorkbookModel,
  type TokenProvider
} from "javascript-spreadsheet-clone";
```

## Headless usage (no UI)

The workbook model is a plain immutable value — every mutation helper returns a new snapshot:

```ts
let workbook = createBlankWorkbook();
workbook = setCellContent(workbook, workbook.activeSheetId, "A1", 2);
workbook = setCellContent(workbook, workbook.activeSheetId, "A2", "=A1*21");

const engine = createFormulaEngine(workbook);
engine.getDisplayValue(workbook.activeSheetId, "A2"); // "42"
engine.update(nextWorkbook); // cheap cell-level diff against the previous snapshot
engine.destroy();            // release HyperFormula resources
```

`importWorkbookFromXlsx(arrayBuffer)` / `exportWorkbookToXlsx(workbook)` round-trip .xlsx files (values, formulas incl. shared formulas, formats, merges, validation, conditional formats, protection, multiple sheets). `importWorkbookFromGoogleSheets(urlOrId, tokenProvider)` pulls a Google Sheet (see [google-sheets-connector.md](google-sheets-connector.md)).

## Embedding the UI

The `<Spreadsheet />` component currently renders the full application (toolbar, formula bar, grid, sheet tabs, status bar) and persists to `localStorage`. Import `src/App.css` alongside it for styles.

Current constraints to plan around (these are the active workstream — see the roadmap):

- No controlled-props API yet (`workbook`/`onWorkbookChange` are planned); state lives inside the component.
- Persistence is hardwired to `localStorage` (`saveWorkbook`/`loadWorkbook` accept any `Storage`, but the component does not yet take one as a prop).
- Styles are a global stylesheet, not CSS modules — scope collisions are possible in a host app.
- No packaged build yet; consume the source via your bundler (Vite/webpack resolve it fine) rather than from npm.

For host-app Google auth, implement `TokenProvider` (`{ getAccessToken(scopes): Promise<string> }`) and pass it to the connector instead of the built-in Google Identity Services flow.
