# Performance Notes

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
- **Virtualized grid** (`src/components/Grid.tsx`): only the visible row window (+16 rows overscan) renders. All sheet-derived data (filter results, row measurements, column widths) is memoized on the sheet snapshot's identity; scroll updates coalesce to one render per frame (leading+trailing rAF throttle); spacer regions paint phantom gridlines so fast flings never flash white.
- **Memoized rows**: `RowFragment` is wrapped in `React.memo` with referentially stable callbacks (ref-proxied). Typing re-renders only the row being edited; scrolling renders only newly exposed rows.
- **Bounded derivations**: pivot source rows, formula-audit dependents, and the status-bar selection summary are gated on their panels being open and bounded by `min(selection area, populated cells)`.

## Memory guards

- **Undo history is capped at 100 entries** (`MAX_UNDO_HISTORY` in `src/lib/workbook.ts`). Each entry pins a copy of the edited sheet's cells-record keys, so an unbounded past grows without limit on big sheets.
- **Autosave skips workbooks over 100k populated cells** (`AUTOSAVE_CELL_LIMIT` in `src/App.tsx`) with a status-bar notice — serializing beyond that stalls every edit and exceeds localStorage quota anyway. Export to XLSX for durable storage of large workbooks.
- Threshold checks count cells with early-exit `for…in` loops — no key-array allocation on the hot path.

## Known remaining work (see docs/audit-2026-07-07.md for the full backlog)

- Rows still receive the whole `sheet` object; passing narrower props would also eliminate the React dev-mode diff cost.
- No column virtualization yet — very wide sheets (hundreds of columns) multiply render cost.
- Whole-column formatting materializes per-cell entries; range-based formats are the planned fix.
- Regression benchmark: `src/lib/formulaEngine.perf.test.ts` (build < 30s, single-cell update < 2s on a 200k-cell sheet).
