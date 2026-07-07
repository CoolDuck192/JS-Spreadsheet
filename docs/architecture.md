# Architecture

## Layout

```
src/
  App.tsx                 orchestration: state, handlers, panels, autosave (2.6k lines)
  components/
    Grid.tsx              virtualized grid + cell editor (memoized rows)
    Toolbar.tsx           ribbon (tabs, groups, buttons)
    FormulaBar.tsx        name box + formula input with autocomplete
    PivotPanel.tsx, ...   feature panels (charts, validation, conditional formats, …)
  lib/
    workbook.ts           THE data model: immutable WorkbookModel snapshots + every
                          mutation helper (cells, formats, sheets, merges, history)
    formulaEngine.ts      persistent HyperFormula wrapper with cell-diff update()
    xlsx.ts               ExcelJS-based .xlsx import/export
    googleSheets.ts       Google Sheets import connector
    googleAuth.ts         TokenProvider interface + Google Identity Services impl
    pivot.ts              pivot aggregation + drill-down source-row tracking
    addressing.ts         A1 <-> coordinate conversion, ranges
    persistence.ts        localStorage save/load with validation + migration
    filters.ts, conditionalFormatting.ts, displayFormat.ts, validation.ts, ...
  index.ts                library entry point (embedding)
```

## Core design: immutable snapshots + incremental engine

Every mutation produces a **new `WorkbookModel`** with structural sharing (unchanged sheets keep identity; the edited sheet gets a fresh `cells` record). Undo/redo is a bounded stack of snapshots (`past`/`present`/`future`, capped at 100).

The formula engine (`formulaEngine.ts`) holds **one persistent HyperFormula instance**. `update(nextWorkbook)` diffs the previous and next snapshots by identity — unchanged sheets are skipped entirely; a changed sheet's cells records are diffed key-by-key and only the delta is applied via `setCellContents`. Structural changes (sheet add/remove/rename, named ranges) trigger a full rebuild through `buildFromSheets` with ragged content-sized arrays. `destroy()` releases HyperFormula but self-revives on the next call (React StrictMode safety).

The engine's `update` runs inside a `useMemo` during render **deliberately**: it's an idempotent diff, and the grid must read the new snapshot's values in the same render pass (the engine's identity never changes, so an effect-based update would strand the UI on stale values). The planned evolution for the embeddable component is `useSyncExternalStore`.

## Rendering strategy

`Grid.tsx` virtualizes rows over prefix-summed row measurements, memoized on the sheet snapshot. Scroll events coalesce to one commit per frame (leading+trailing rAF). Rows are `React.memo` components fed referentially stable callbacks through a ref proxy, so typing re-renders one row and scrolling renders only newly exposed rows. See [performance.md](performance.md) for measured numbers and remaining work.

## Known-issues backlog

`docs/audit-2026-07-07.md` holds a 164-item audit (deduplicated only partially — repeats corroborate) with failure scenarios and fix sketches, ordered by severity. The highest-impact open items: HyperFormula-backed structural operations (the current regex-based formula rewriting on row/column insert/delete can corrupt formulas), cross-sheet reference rewrites, column virtualization, and the controlled-props embedding API.
