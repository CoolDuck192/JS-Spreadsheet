# Embeddable DataTable Implementation Program

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Ship a pressure-tested React `DataTable` for local and remote records plus live Excel-style workbook tables without regressing the existing spreadsheet.

**Architecture:** A narrow shared interaction and command lifecycle connects React views to source-specific sessions. `RecordTableSession` owns local or remote record semantics; `WorkbookSession` owns workbook and HyperFormula semantics; `WorkbookTableSession` is a live range projection and never copies worksheet data.

**Tech Stack:** TypeScript 6, React 19, Vite 8, Vitest 4, Testing Library, Playwright, HyperFormula 3.3, ExcelJS 4.4, pnpm 11.

## Global Constraints

- Node.js must be 22.13 or newer and pnpm must remain pinned to 11.7.0.
- Bind every user-facing dev or preview server to `0.0.0.0`; when handing a running build to the user, print the machine IP and port instead of a localhost URL. Automated browser tests may still use loopback internally.
- The first release is React-only; do not add Laravel, vanilla mount, Web Component, Vue, Svelte, or Angular bindings.
- Keep the headless core free of React, DOM, storage, authentication, and module-level browser side effects.
- Never copy a workbook table range into a second authoritative row model.
- Never present loaded-page sorting, filtering, formulas, grouping, or aggregation as dataset-wide behavior.
- Keep workbook and remote write, history, formula, validation-authority, and conflict semantics source-specific.
- Preserve the approved neutral slate, white, and muted-green tokens; scope all library styles beneath component roots.
- Preserve GPL-3.0-or-later and HyperFormula licensing notices in package and embedding documentation.
- Do not publish or open the pull request until unit, contract, property/fuzz, race, XLSX, React, E2E, accessibility, performance, memory, and consumer-build gates pass.
- Do not commit `.superpowers/` visual-companion state.

---

## Plan Set

Execute these plans in dependency order:

1. [Correctness and Workbook Session Foundation](2026-07-09-data-table-foundation.md)
2. [Local React DataTable](2026-07-09-local-react-data-table.md)
3. [Structured Workbook Tables](2026-07-09-workbook-structured-tables.md)
4. [Remote DataTable Source](2026-07-09-remote-data-table.md)
5. [Packaging and Pressure Gate](2026-07-09-data-table-packaging-pressure.md)

Plans 3 and 4 may run in parallel after Plan 2 is merged into the feature branch. Plan 5 begins only after both are integrated.

## Execution Topology

```text
Plan 1: correctness + WorkbookSession + controlled Spreadsheet
  |
  v
Plan 2: shared viewport + local RecordTableSession + React DataTable
  |-----------------------------|
  v                             v
Plan 3: workbook tables         Plan 4: remote source
  |-----------------------------|
                v
Plan 5: packaging + complete pressure gate + documentation
                |
                v
Independent code review -> clean verification run -> new pull request
```

## Program Checkpoints

- [ ] **Checkpoint 1: Establish an isolated execution worktree**

Use `superpowers:using-git-worktrees` before implementation. Create a feature branch from the plan commit that descends from `87874f2`, then verify the baseline:

```bash
corepack pnpm test
corepack pnpm run build
corepack pnpm run test:e2e
```

Expected: all existing tests pass before source changes.

- [ ] **Checkpoint 2: Complete Plan 1 and run spreadsheet parity**

Run:

```bash
corepack pnpm test
corepack pnpm run build
corepack pnpm run test:e2e
```

Expected: the atomic session and correctness fixes pass without changing existing visible spreadsheet behavior.

- [ ] **Checkpoint 3: Complete Plan 2 and validate the public local API**

Run the local table unit, React, accessibility, and 100,000-row viewport tests named in Plan 2, followed by the full unit suite.

Expected: simple local `rows/columns/getRowId/onRowsChange` usage and session-backed usage both pass.

- [ ] **Checkpoint 4: Integrate Plans 3 and 4**

Resolve integration only through the shared contracts from Plan 2. Do not make workbook sources implement remote cache APIs or make remote sources mutate workbook state.

Expected: workbook and remote adapter contract suites share observable cases only where both advertise support.

- [ ] **Checkpoint 5: Complete Plan 5 release gate**

Run the single release command introduced by Plan 5:

```bash
corepack pnpm run verify:release
```

Expected: exit 0 with unit, type, build, consumer, browser, accessibility, race, fuzz, XLSX, performance, and memory summaries.

- [ ] **Checkpoint 6: Request independent review**

Use `superpowers:requesting-code-review` and provide reviewers with the design spec, these five plans, the full diff, and verification output. Address every confirmed P0/P1 finding and rerun `verify:release`.

- [ ] **Checkpoint 7: Prepare and open a new pull request**

Use the repository PR workflow only after the branch is clean and the release gate passes. The PR must include:

- Scope and architecture summary.
- Local, remote, and workbook usage examples.
- Structured-table XLSX behavior and identity caveats.
- Verification commands and measured performance results.
- Licensing note.
- Known non-goals, including Laravel and non-React bindings.

Do not merge the PR automatically. Leave it ready for user review.

## Program Definition of Done

- The simple local React API works with typed row inference.
- Frappe-parity table interactions are explicit: checkbox row selection, expandable tree rows, inline filters, dynamic row height, custom header actions, serial row headers, visible column reorder/hide/resize, flexible layouts, clipboard, and custom editors/formatters.
- Remote queries, mutations, cancellation, versions, conflicts, subscriptions, and compensating undo follow their declared capabilities.
- Workbook tables are live projections, support the contextual Table tab, and import/export as native Excel tables.
- Existing spreadsheet workflows retain parity.
- Public ESM, declaration, connector, and CSS subpaths build in Vite and webpack consumers.
- Styles do not leak into the host.
- Chromium, Firefox, and WebKit pass keyboard and core interaction flows.
- The pinned performance environment meets accepted budgets at 100,000 local rows and million-row remote totals with bounded loaded pages.
- No unresolved P0/P1 review finding remains.
- `verify:release` passes from a clean checkout.
