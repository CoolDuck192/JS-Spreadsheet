# Embedding JS-Spreadsheet in Your App

There is no npm package yet — add the cloned repo as a source dependency and let your bundler compile it (the package's `main` points at `src/index.ts`, so the bundler must handle TypeScript in dependencies; Vite and modern webpack setups do):

```bash
pnpm add file:../JS-Spreadsheet     # or: pnpm add github:CoolDuck192/JS-Spreadsheet
```

`src/index.ts` is the library entry point. The default export is the full spreadsheet UI; named exports are the headless functions.

```tsx
import Spreadsheet from "js-spreadsheet";        // full UI component
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
} from "js-spreadsheet";
```

**Licensing note:** the formula engine is [HyperFormula](https://hyperformula.handsontable.com/), which this app initializes with its GPLv3 license key. If your host application cannot comply with GPLv3, you need a commercial HyperFormula license.

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

The `<Spreadsheet />` component currently renders the full application (toolbar, formula bar, grid, sheet tabs, status bar) and persists to `localStorage`. Import the stylesheet alongside it:

```ts
import "js-spreadsheet/src/App.css";
```

Current constraints to plan around (these are the active workstream — see the roadmap):

- No controlled-props API yet (`workbook`/`onWorkbookChange` are planned); state lives inside the component.
- Persistence is hardwired to `localStorage` (`saveWorkbook`/`loadWorkbook` accept any `Storage`, but the component does not yet take one as a prop).
- Styles are a global stylesheet, not CSS modules — scope collisions are possible in a host app.
- No packaged build yet; consume the source via your bundler (Vite/webpack resolve it fine) rather than from npm.

For host-app Google auth, implement `TokenProvider` (`{ getAccessToken(scopes): Promise<string> }`) and pass it to the connector instead of the built-in Google Identity Services flow.
