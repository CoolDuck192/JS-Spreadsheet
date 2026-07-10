# Performance Notes

## Workbook-session foundation gate (2026-07-10)

The reproducible foundation benchmark runs on Node `22.13.1` and pnpm `11.7.0`.
The current reference host is Linux x86-64 with an Intel Xeon W-2133 (6 cores / 12
threads) and 62 GiB RAM. `src/core/workbook/WorkbookSession.perf.test.ts` builds a
100,000-row workbook with a live aggregate formula, edits and recalculates one
cell, commits a 100-command transaction, verifies projection-before-publication,
and exercises weight-bounded undo retention. The initial correctness-only
calibration completed the test body in 17.3 seconds; the milestone gate enforces
the individual build, edit, burst, and history budgets unless
`SKIP_PERF_ASSERT=1` is set explicitly.

Typed input is parsed once at the command boundary: numbers and booleans remain
typed, dates become Excel serials with an inferred number format, formulas remain
formulas, and literal text is preserved. Every durable change is reduced through
an atomic command result (`committed`, `rejected`, or `conflict`); rejected
transactions publish neither a partial workbook nor a stale formula projection.

Workbook history is bounded by both retained snapshot count and deterministic
logical weight. Defaults are 100 retained entries and 2,000,000 logical entries;
the present snapshot is always kept while the oldest retained past/future
snapshots are trimmed. This prevents a handful of very large workbooks from
defeating a count-only limit.

## Measured baseline (2026-07-07, 100k-row CSV ≈ 300k cells)

| Metric | Production build | Dev server |
| --- | --- | --- |
| Import 100k rows | ~1.1s | ~9.5s |
| Single cell edit | ~1.2s | ~11s |
| Long frames (>50ms) during scroll stress | 2 | ~40 |
| JS heap | ~90–106 MB | similar |

**The dev/prod gap is React's development build, not app code.** CPU profiling attributes the bulk of dev-mode edit time to React dev instrumentation diffing the large `sheet` prop (300k-key cells record) passed to row components. Demo and test large datasets against `pnpm exec vite preview`.

## How the app stays fast

- **Incremental formula engine** (`src/lib/formulaEngine.ts`): one persistent HyperFormula instance. Edits are applied as cell-level diffs between immutable workbook snapshots — O(changed cells), not O(grid). Initial load bulk-feeds `buildFromSheets` with ragged arrays sized to actual content. Grid limits are Excel's (1,048,576 × 16,384).
- **Virtualized grid** (`src/components/Grid.tsx`): both rows and columns are measured and windowed with bounded overscan. A 100,000 × 10,000 sheet renders only the visible cross-product while keeping frozen/hidden indexes, merged-cell anchors, selection/editor overlays, and `ensureCellVisible` geometry correct. All sheet-derived data is memoized on the sheet snapshot's identity; scroll updates coalesce to one render per frame (leading+trailing rAF throttle).
- **Memoized rows**: `RowFragment` is wrapped in `React.memo` with referentially stable callbacks (ref-proxied). Typing re-renders only the row being edited; scrolling renders only newly exposed rows.
- **Bounded derivations**: pivot source rows, formula-audit dependents, and the status-bar selection summary are gated on their panels being open and bounded by `min(selection area, populated cells)`.

## Memory guards

- **Undo history is count- and weight-bounded** (`src/core/workbook/history.ts`). Each entry can pin large sheet records, so the session trims against both 100 retained entries and 2,000,000 deterministic logical entries by default.
- **Autosave skips workbooks over 100k populated cells** (`AUTOSAVE_CELL_LIMIT` in `src/App.tsx`) with a status-bar notice — serializing beyond that stalls every edit and exceeds localStorage quota anyway. Export to XLSX for durable storage of large workbooks.
- Threshold checks count cells with early-exit `for…in` loops — no key-array allocation on the hot path.

## Known remaining work (see docs/audit-2026-07-07.md for the full backlog)

- Rows still receive the whole `sheet` object; passing narrower props would also eliminate the React dev-mode diff cost.
- Whole-column formatting materializes per-cell entries; range-based formats are the planned fix.
- Regression benchmarks: `src/lib/formulaEngine.perf.test.ts` (build < 30s, single-cell update < 2s on a 200k-cell sheet) and `src/core/workbook/WorkbookSession.perf.test.ts` (100k-row atomic session and history pressure).
