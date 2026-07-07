# Features Tour

## Grid and editing

Excel-like grid with Excel-sized limits (1,048,576 rows × 16,384 columns), row-window virtualization, frozen panes, merged cells, protected sheets/read-only cells, zoom, print view, and manual/auto-fit row and column sizing. Cell editing commits with Enter; Escape cancels; formula autocomplete follows Excel semantics (Tab accepts the highlighted suggestion, ↑/↓ navigate, ←/→ move the caret, Enter always commits what you typed).

## Formulas

HyperFormula 3.3 engine (400+ Excel-compatible functions), configured for Excel parity: ISO (`2026-01-15`) and US (`01/15/2026`) date entry parses to date serials matching Excel's numbering exactly, typing a date auto-applies the date format, date arithmetic works (`=B1-A1`), and booleans display as `TRUE`/`FALSE`. Formula auditing shows precedents/dependents with jump navigation; a searchable function library and AutoSum helpers round it out.

## Pivot tables with drill-down

Select a data range → Insert → Pivot Table. Configure row fields (primary + detail), an optional column field, value field, and aggregator (SUM/COUNT/AVERAGE/MIN/MAX). The pivot lands on a new sheet with totals — and **double-clicking any value, subtotal, or grand-total cell opens a Details sheet listing exactly the source rows behind that number**, like Excel's Show Details. Drill-down metadata is session-scoped and guards against stale sheets.

## Import and export

- **Drag-and-drop**: drop an `.xlsx` or `.csv` anywhere on the app.
- **XLSX fidelity**: values, formulas (shared formulas translated per cell), date cells as real dates, number/currency/percent/date formats, merges, comments, hyperlinks, validation rules, conditional formats, protection with unlocked cells, freeze panes, named ranges, AutoFilter ranges, hidden rows/columns/sheets.
- **CSV** import/export, plus **Google Sheets** import (see [google-sheets-connector.md](google-sheets-connector.md)).

## Data tools

Find/replace, remove duplicates, sorting, header AutoFilter menus with unique-value filtering, custom filters, live selection statistics (count/sum/avg/min/max) in the status bar, data validation (lists with in-cell pickers, number bounds, text length), conditional formatting (value rules, blanks/duplicates/top-bottom, data bars, color scales), embedded bar/line/pie charts, comments, hyperlinks, and named ranges.

## Persistence

Debounced autosave to browser localStorage with startup recovery, quota-safe failure handling, and an explicit size cutoff (~100k cells) beyond which the status bar directs you to XLSX export. Undo history holds the last 100 edits.
