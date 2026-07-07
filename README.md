# JavaScript Spreadsheet Clone

A standalone browser spreadsheet built with Vite, React, TypeScript, and HyperFormula.

## Features

- Excel-like grid with row and column headers.
- Cell editing, formula bar, keyboard navigation, keyboard shortcuts for common formatting/link/filter actions, range selection, row/column/sheet header selection, merged cells, protected sheets, read-only cells, plain external paste plus internal cut/copy/paste with formulas, formatting, validation, comments, hyperlinks, paste values, paste formats, and transpose paste, fill-down, fill-right, AutoFill handle number/date/month/weekday series and formula extension with copied formatting/validation, undo, redo, clear all/contents/formats/conditional formats/hyperlinks/validation, toolbar and context-menu row/column insert/delete, manual and auto-fit row/column sizing, row/column hide and unhide, sheet tab hide and restore, freeze panes, reset view, worksheet gridline/header/formula-bar/formula-text/sheet-tab visibility, worksheet zoom controls, and print-ready worksheet output.
- Data workflow tools for AutoSum and quick Sum/Average/Count/Min/Max formula insertion, formula auditing for same-sheet precedents/dependents with jump navigation, Go To navigation for references and named ranges, named ranges from the name box with a manager for selecting/deleting names, cell comments with context-menu clearing, cell hyperlinks, find/replace, remove duplicates, selected-range filtering, header AutoFilter menus with unique-value filtering and sort actions, header-aware AutoFilter table sorting A-Z or Z-A, and live selection summaries for count, sum, average, min, and max.
- HyperFormula-powered formulas, including supported Excel-style functions such as `SUM`, `AVERAGE`, `IF`, references, ranges, and formula errors.
- Searchable function library plus formula autocomplete backed by HyperFormula's registered function list when entering formulas from the formula bar or directly inside a cell, with click and keyboard selection.
- Pivot table builder for summarizing selected tabular ranges by primary and detail row fields, column field, value field, and aggregation, with formatted-number parsing, naturally sorted labels, and fast single-pass grouping; embedded bar, line, and pie charts for selected label/value ranges.
- Large-sheet handling with bulk paste/import writes and viewport row rendering to keep bigger datasets responsive.
- Cell formatting for bold, italic, wrapped text, text color, fill color, Format Painter, borders, number/currency/percent/date display, horizontal and vertical alignment, data validation with in-cell list pickers, number bounds, text-length rules, and visible rule management, plus conditional-format highlights, data bars, and color scales including blank/not-blank, duplicate/unique, and top/bottom ranked-value rules with visible rule management.
- Sheet tabs with add, rename, duplicate, delete, hide, restore, reorder, color, and switch actions.
- CSV and XLSX import/export, including common cell styling, data validation rules, conditional formats, protected sheets with unlocked cells, freeze panes, named ranges, AutoFilter ranges, active sheet tabs, hidden sheet tabs, blank sheets, and workbook structure such as dimensions, merges, comments, hyperlinks, and hidden rows/columns. Shared formulas are translated per cell on import, and date cells import as real dates.
- Drag-and-drop import: drop an `.xlsx` or `.csv` file anywhere on the app to load it.
- Pivot drill-down: double-click any pivot value or total cell to open a `Details` sheet with the source rows behind that number, like Excel's Show Details.
- Google Sheets connector: link a Google Sheet from Data → Link Google Sheet (see setup below).
- Browser local-storage persistence with startup recovery for stale saved sheet state, debounced autosave, and graceful handling of full browser storage.
- Excel-sized grid limits (1,048,576 rows × 16,384 columns) with an incremental formula engine: edits apply as cell-level diffs instead of engine rebuilds, so large sheets stay responsive.

## Install And Run

Prerequisites:

- Git
- **Node.js 22.13 or newer** (the pinned pnpm 11 requires it — on Node 20 `pnpm` fails with `ERR_UNKNOWN_BUILTIN_MODULE: node:sqlite`)
- pnpm 11 via corepack (no separate install needed)

Clone the repository:

```bash
git clone https://github.com/CoolDuck192/JS-Spreadsheet.git
cd JS-Spreadsheet
```

Install dependencies:

```bash
corepack enable
pnpm install    # corepack downloads the pinned pnpm 11.7.0 automatically
```

Start the local dev server:

```bash
pnpm run dev
```

The dev server defaults to:

```text
http://127.0.0.1:5173
```

To reach it from another machine on your network, bind all interfaces:

```bash
pnpm exec vite --host 0.0.0.0
```

Build and preview a production bundle:

```bash
pnpm run build
pnpm exec vite preview --host 0.0.0.0   # serves the build on port 4173
```

> **Testing with large datasets?** Use the production preview. The dev server
> carries React development-build overhead that makes 100k-row sheets feel far
> slower than they are in production (measured: single edits ~1.2s in the
> production build vs ~11s under the dev server on a 100k-row sheet).

Run the test suite:

```bash
pnpm test
```

Run browser end-to-end tests:

```bash
pnpm exec playwright install
pnpm run test:e2e
```

## Project Docs

- [docs/getting-started.md](docs/getting-started.md) — prerequisites, setup, dev vs production preview, tests, first five minutes.
- [docs/features.md](docs/features.md) — full feature tour (grid, formulas, pivots with drill-down, import/export, data tools).
- [docs/google-sheets-connector.md](docs/google-sheets-connector.md) — one-time OAuth setup and how the connector works.
- [docs/embedding.md](docs/embedding.md) — using the library entry point headlessly or embedding the UI in your own app.
- [docs/performance.md](docs/performance.md) — measured 100k-row numbers, how the app stays fast, memory guards.
- [docs/architecture.md](docs/architecture.md) — source layout, the immutable-snapshot + incremental-engine design.
- [docs/audit-2026-07-07.md](docs/audit-2026-07-07.md) — known-issues backlog with failure scenarios and fix sketches.

## Formula Engine

Formula support is provided by HyperFormula 3.3.0 using the GPLv3 license key string. HyperFormula supports a broad Excel-compatible function set; exact function coverage follows the installed engine version.

The engine is configured for Excel parity: 1,048,576 × 16,384 grid limits, ISO (`2026-01-15`) and US (`01/15/2026`) date entry parsed as date serials (typing a date auto-applies the date format, and date arithmetic like `=B1-A1` works), the 1900 leap-year compatibility behavior, and `TRUE`/`FALSE` boolean display.

## Google Sheets Connector

Linking a Google Sheet needs a Google OAuth client id (the app is a static SPA, so no server or client secret is involved):

1. In [Google Cloud Console](https://console.cloud.google.com/), create a project, enable the **Google Sheets API**, and create an **OAuth client ID** of type *Web application* with your app's origin (e.g. `http://127.0.0.1:5173`) in *Authorized JavaScript origins*.
2. Put the client id in `.env.local`:

   ```bash
   VITE_GOOGLE_CLIENT_ID=1234567890-abc.apps.googleusercontent.com
   ```

3. Restart the dev server, then use **Data → Link Google Sheet** and paste a sheet URL. The user signs in with their own Google account; formulas are imported as formulas (`valueRenderOption=FORMULA`).

Embedders can supply their own token source instead by implementing the `TokenProvider` interface in `src/lib/googleAuth.ts`.

## Embedding

`src/index.ts` exports the app as a library: the default export is the `<Spreadsheet />` component, alongside headless workbook, engine, xlsx, pivot, and Google Sheets functions. A dedicated Vite library build target is on the roadmap.
