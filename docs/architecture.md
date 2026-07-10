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
    xlsx.ts               .xlsx import/export with native structured-table metadata
    googleSheets.ts       Google Sheets import connector
    googleAuth.ts         TokenProvider interface + Google Identity Services impl
    pivot.ts              pivot aggregation + drill-down source-row tracking
    addressing.ts         A1 <-> coordinate conversion, ranges
    persistence.ts        localStorage save/load with validation + migration
    filters.ts, conditionalFormatting.ts, displayFormat.ts, validation.ts, ...
  table/
    core/                 source-neutral queries, capabilities, sessions, commands
    local/                complete in-memory record source and history
    remote/               abortable queries, optimistic edits, conflicts, subscriptions
    workbook/             live adapter from structured workbook tables to DataTable
  react/
    DataTable.tsx         embeddable local/session-backed React table surface
    viewport/             shared two-axis virtualized grid kernel
  entry/
    core.ts               DOM-free package entry point
    react.ts              Spreadsheet and DataTable package/root entry point
    google.ts             optional Google connector entry point
```

The application and library are built separately. `dist/app` is the standalone
Vite application; `dist/lib` and `dist/types` contain the ESM package,
declarations, and scoped stylesheet. React is a peer dependency, and the core
entry does not import React, CSS, storage, or browser globals.

## Table architecture

`DataTable` consumes the same `TableSession` contract for three source types.
Local sessions own a complete record set and can safely run whole-dataset
operations. Remote sessions expose only server-declared capabilities and
coordinate cancellation, pagination, optimistic overlays, versioned conflicts,
subscriptions, and compensating undo. Workbook-table sessions are live adapters
over a parent `WorkbookSession`, so edits and undo history remain shared with the
spreadsheet instead of being copied into a second state store.

Both the spreadsheet and `DataTable` render through the shared two-axis viewport
kernel. Source behavior stays in the session layer; React owns presentation,
keyboard interaction, extension boundaries, and `Blob` conversion for downloads.

## Core design: immutable snapshots + incremental engine

Every mutation produces a **new `WorkbookModel`** with structural sharing (unchanged sheets keep identity; the edited sheet gets a fresh `cells` record). Undo/redo is a bounded stack of snapshots (`past`/`present`/`future`, capped at 100).

The formula engine (`formulaEngine.ts`) holds **one persistent HyperFormula instance**. `update(nextWorkbook)` diffs the previous and next snapshots by identity — unchanged sheets are skipped entirely; a changed sheet's cells records are diffed key-by-key and only the delta is applied via `setCellContents`. Structural changes (sheet add/remove/rename, named ranges) trigger a full rebuild through `buildFromSheets` with ragged content-sized arrays. `destroy()` releases HyperFormula but self-revives on the next call (React StrictMode safety).

The engine's `update` runs inside a `useMemo` during render **deliberately**: it's an idempotent diff, and the grid must read the new snapshot's values in the same render pass (the engine's identity never changes, so an effect-based update would strand the UI on stale values). The planned evolution for the embeddable component is `useSyncExternalStore`.

## Rendering strategy

`Grid.tsx` virtualizes rows over prefix-summed row measurements, memoized on the sheet snapshot. Scroll events coalesce to one commit per frame (leading+trailing rAF). Rows are `React.memo` components fed referentially stable callbacks through a ref proxy, so typing re-renders one row and scrolling renders only newly exposed rows. See [performance.md](performance.md) for measured numbers and remaining work.

## Known-issues backlog

`docs/audit-2026-07-07.md` holds a 164-item audit (deduplicated only partially — repeats corroborate) with failure scenarios and fix sketches, ordered by severity. The highest-impact open items: HyperFormula-backed structural operations (the current regex-based formula rewriting on row/column insert/delete can corrupt formulas), cross-sheet reference rewrites, column virtualization, and the controlled-props embedding API.
