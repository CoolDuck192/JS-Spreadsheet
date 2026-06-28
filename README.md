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
- CSV and XLSX import/export, including common cell styling, data validation rules, conditional formats, protected sheets with unlocked cells, freeze panes, named ranges, AutoFilter ranges, active sheet tabs, hidden sheet tabs, blank sheets, and workbook structure such as dimensions, merges, comments, hyperlinks, and hidden rows/columns.
- Browser local-storage persistence with startup recovery for stale saved sheet state.

## Install And Run

Prerequisites:

- Git
- Node.js 20 or newer
- pnpm 11 or newer

Clone the repository:

```bash
git clone https://github.com/CoolDuck192/JS-Spreadsheet.git
cd JS-Spreadsheet
```

Install dependencies:

```bash
corepack enable
corepack prepare pnpm@11.7.0 --activate
pnpm install
```

Start the local dev server:

```bash
pnpm run dev
```

The dev server defaults to:

```text
http://127.0.0.1:5173
```

Build a production bundle:

```bash
pnpm run build
```

Run the test suite:

```bash
pnpm test
```

Run browser end-to-end tests:

```bash
pnpm exec playwright install
pnpm run test:e2e
```

## Formula Engine

Formula support is provided by HyperFormula 3.3.0 using the GPLv3 license key string. HyperFormula supports a broad Excel-compatible function set; exact function coverage follows the installed engine version.
