# JS-Spreadsheet

A standalone browser spreadsheet built with Vite, React, TypeScript, and HyperFormula.

## Features

- Excel-like grid with row and column headers.
- Embeddable React `DataTable` for local records, host-authoritative remote APIs,
  and live structured workbook tables, with typed columns, virtualization,
  editing, query tools, conflict handling, undo, and CSV/XLSX export.
- Cell editing, formula bar, keyboard navigation, keyboard shortcuts for common formatting/link/filter actions, range selection, row/column/sheet header selection, merged cells, protected sheets, read-only cells, plain external paste plus internal cut/copy/paste with formulas, formatting, validation, comments, hyperlinks, paste values, paste formats, and transpose paste, fill-down, fill-right, AutoFill handle number/date/month/weekday series and formula extension with copied formatting/validation, undo, redo, clear all/contents/formats/conditional formats/hyperlinks/validation, toolbar and context-menu row/column insert/delete, manual and auto-fit row/column sizing, row/column hide and unhide, sheet tab hide and restore, freeze panes, reset view, worksheet gridline/header/formula-bar/formula-text/sheet-tab visibility, worksheet zoom controls, and print-ready worksheet output.
- Data workflow tools for AutoSum and quick Sum/Average/Count/Min/Max formula insertion, formula auditing for same-sheet precedents/dependents with jump navigation, Go To navigation for references and named ranges, named ranges from the name box with a manager for selecting/deleting names, cell comments with context-menu clearing, cell hyperlinks, find/replace, remove duplicates, selected-range filtering, header AutoFilter menus with unique-value filtering and sort actions, header-aware AutoFilter table sorting A-Z or Z-A, and live selection summaries for count, sum, average, min, and max.
- HyperFormula-powered formulas, including supported Excel-style functions such as `SUM`, `AVERAGE`, `IF`, references, ranges, and formula errors.
- Searchable function library plus formula autocomplete backed by HyperFormula's registered function list when entering formulas from the formula bar or directly inside a cell, with click and keyboard selection.
- Pivot table builder for summarizing selected tabular ranges by primary and detail row fields, column field, value field, and aggregation, with formatted-number parsing, naturally sorted labels, and fast single-pass grouping; embedded bar, line, and pie charts for selected label/value ranges.
- Large-sheet handling with bulk paste/import writes and viewport row rendering to keep bigger datasets responsive.
- Cell formatting for bold, italic, wrapped text, font family and size (stored in points, as in Excel, and round-tripped through XLSX), text color, fill color, Format Painter, borders, number/currency/percent/date display, horizontal and vertical alignment, data validation with in-cell list pickers, number bounds, text-length rules, and visible rule management, plus conditional-format highlights, data bars, and color scales including blank/not-blank, duplicate/unique, and top/bottom ranked-value rules with visible rule management.
- Excel-style ribbon toolbar with selection-aware formatting state: toggles and selects reflect the whole selection and show a mixed indicator when the range disagrees (toggling from mixed applies to all, like Excel); split buttons for Paste, AutoSum, and Borders variants; quick $ / % / , number-format buttons; panel-launcher buttons that show open state; and shortcut tooltips that are platform-aware (⌘ on Apple platforms, Ctrl elsewhere — both bindings work everywhere).
- Sheet tabs with add, rename, duplicate, delete, hide, restore, reorder, color, and switch actions.
- CSV and XLSX import/export, including common cell styling, data validation rules, conditional formats, protected sheets with unlocked cells, freeze panes, named ranges, AutoFilter ranges, active sheet tabs, hidden sheet tabs, blank sheets, and workbook structure such as dimensions, merges, comments, hyperlinks, and hidden rows/columns. Shared formulas are translated per cell on import, and date cells import as real dates.
- Drag-and-drop import: drop an `.xlsx` or `.csv` file anywhere on the app to load it.
- Pivot drill-down: double-click any pivot value or total cell to open a `Details` sheet with the source rows behind that number, like Excel's Show Details.
- Google Sheets connector: one-time, read-only workbook import from File → Import Google Sheet (see setup below).
- Browser local-storage persistence with startup recovery for stale saved sheet state, debounced autosave, and graceful handling of full browser storage.
- Excel-sized grid limits (1,048,576 rows × 16,384 columns) with an incremental formula engine: edits apply as cell-level diffs instead of engine rebuilds, so large sheets stay responsive.

## Install And Run

Prerequisites:

- Git
- **Node.js 22.13 or newer to develop or build this repository** (the pinned pnpm 11 requires it — on Node 20 `pnpm` fails with `ERR_UNKNOWN_BUILTIN_MODULE: node:sqlite`)
- pnpm 11 via corepack (no separate install needed)

This Node version is a repository-tooling requirement, not a runtime engine
requirement for applications consuming the built browser package.

Clone the repository:

```bash
git clone https://github.com/CoolDuck192/JS-Spreadsheet.git
cd JS-Spreadsheet
```

Activate the pinned Node.js version and install dependencies. If `nvm` is not
available, the fallback downloads a portable Node binary and puts it first on
`PATH` for pnpm, Vite, and Playwright:

```bash
if command -v nvm >/dev/null 2>&1; then
  nvm install
  nvm use
else
  NODE22_BIN="$(npm exec --yes --package=node@22.13.1 -- node -p 'process.execPath')"
  export PATH="$(dirname "$NODE22_BIN"):$PATH"
  hash -r
fi
corepack --version
corepack pnpm install --frozen-lockfile
```

Start the local dev server:

```bash
corepack pnpm run dev
```

The dev server binds all interfaces. Open the `Network` URL that Vite reports,
or use the machine IP and port directly:

```text
http://<machine-ip>:5173
```

When handing off a manually started server, report the machine IP, port, and
serving process so another reviewer can open the same instance.

Build and preview a production bundle:

```bash
corepack pnpm run build
corepack pnpm exec vite preview --host 0.0.0.0 --port 4173 --strictPort
```

> **Testing with large datasets?** Use the production preview. The dev server
> carries React development-build overhead that makes 100k-row sheets feel far
> slower than they are in production (measured: single edits ~1.2s in the
> production build vs ~11s under the dev server on a 100k-row sheet).

Run the test suite:

```bash
corepack pnpm test
```

Run browser end-to-end tests. With the production preview above still running,
the external base URL makes Playwright reuse that one server instead of starting
its development server:

```bash
corepack pnpm exec playwright install
E2E_BASE_URL=http://192.168.6.232:4173 corepack pnpm run test:e2e
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

> **Using this in your own project?** The GPLv3 HyperFormula key means your use must be GPLv3-compatible; otherwise you need a [commercial HyperFormula license](https://hyperformula.handsontable.com/guide/license-key.html).

The engine is configured for Excel parity: 1,048,576 × 16,384 grid limits, ISO (`2026-01-15`) and US (`01/15/2026`) date entry parsed as date serials (typing a date auto-applies the date format, and date arithmetic like `=B1-A1` works), the 1900 leap-year compatibility behavior, and `TRUE`/`FALSE` boolean display.

## Google Sheets Connector

Choose **File → Import Google Sheet**. This is a one-time, read-only import that
replaces the current workbook through normal workbook history; it does not
create a live link, refresh job, write-back path, or synchronization state.

A fresh standalone clone opens an in-app setup dialog when no client ID is
configured. Create a Web application OAuth client in Google Cloud, enable the
Google Sheets API, and register the app's exact eligible HTTPS DNS origin under
Authorized JavaScript origins. A raw LAN IP is intentionally reported as
incompatible with built-in browser OAuth; expose the app through an HTTPS DNS
name or have the host provide a token provider. Only the public client ID is
saved, under `javascript-spreadsheet.google-client-id.v1`; client secrets and
access tokens are never stored by this workflow.

Embedded applications configure `services.googleSheets` with a host
`tokenProvider`, or with `clientId` plus an optional `tokenProviderFactory`.
They may explicitly provide `clientIdStorage`; otherwise embedded client-ID
storage is disabled. Set `features={{ googleSheets: false }}` to remove the
command. See [docs/google-sheets-connector.md](docs/google-sheets-connector.md)
and [docs/embedding.md](docs/embedding.md) for CSP, iframe popup, and service
configuration details.

## Embedding

Build the typed library with `corepack pnpm run build:lib`. React hosts use
`js-spreadsheet/react` (or the root alias), framework-independent code uses
`js-spreadsheet/core`, and both UI surfaces use `js-spreadsheet/styles.css`.
See [docs/embedding.md](docs/embedding.md) for local rows, remote APIs, workbook
tables, package entrypoints, lifecycle rules, and XLSX identity behavior.

## License

[GPL-3.0-or-later](LICENSE). This matches the project's use of HyperFormula under its GPLv3 license key — if you need to use this code in a non-GPL application, you would also need a commercial HyperFormula license.
