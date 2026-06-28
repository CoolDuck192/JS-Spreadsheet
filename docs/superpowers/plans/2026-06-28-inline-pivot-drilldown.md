# Inline Pivot Drilldown Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add inline expand/collapse drilldown rows to generated pivot sheets and open a GitHub PR.

**Architecture:** Extend the pivot builder to return pivot metadata alongside the existing output matrix. Store that metadata on the generated sheet, materialize expanded detail rows from it, and render a small drilldown button in drillable pivot value cells.

**Tech Stack:** React, TypeScript, Vite, Vitest, Playwright, GitHub CLI.

---

### Task 1: Pivot Metadata Tests

**Files:**
- Modify: `src/lib/pivot.test.ts`

- [ ] **Step 1: Write failing unit tests**

Add tests that import `createPivotTableWithDrilldowns` and `materializePivotRows`. The tests should assert that an East/Hardware nested pivot entry contains the matching raw source row and that expanding the entry inserts a source header plus source row below the pivot row.

- [ ] **Step 2: Verify red**

Run: `pnpm exec vitest run src/lib/pivot.test.ts -t drilldown --reporter=dot`

Expected: fail because the new pivot drilldown exports do not exist yet.

### Task 2: Pivot Metadata Implementation

**Files:**
- Modify: `src/types.ts`
- Modify: `src/lib/pivot.ts`

- [ ] **Step 1: Add serializable pivot metadata types**

Add `PivotSheetMetadata`, `PivotDrilldownEntry`, and related row-kind types to `src/types.ts`.

- [ ] **Step 2: Build metadata in pivot creation**

Add `createPivotTableWithDrilldowns(rows, config)` in `src/lib/pivot.ts`. It should preserve `createPivotTable(rows, config).rows` behavior while also storing source rows per drillable summary cell.

- [ ] **Step 3: Materialize expanded rows**

Add `materializePivotRows(metadata)`, `getPivotDrilldownCell(metadata, row, column)`, and `getPivotMaterializedRowKind(metadata, row)` helpers.

- [ ] **Step 4: Verify green**

Run: `pnpm exec vitest run src/lib/pivot.test.ts -t drilldown --reporter=dot`

Expected: pass.

### Task 3: App And Grid Integration

**Files:**
- Modify: `src/App.tsx`
- Modify: `src/components/Grid.tsx`
- Modify: `src/App.css`

- [ ] **Step 1: Add failing browser test**

Add a Playwright test that creates a nested pivot, clicks the C2 drilldown expand button, verifies source rows appear below it, clicks collapse, and verifies the rows are removed.

- [ ] **Step 2: Verify red**

Run: `pnpm exec playwright test tests/spreadsheet.spec.ts -g "expands pivot drilldown rows inline" --reporter=line`

Expected: fail because the UI control does not exist.

- [ ] **Step 3: Wire pivot creation and toggling**

Update `handleCreatePivotTable` to use `createPivotTableWithDrilldowns`, store metadata on the new sheet, and render the materialized matrix. Add a toggle handler that flips the expanded entry and rebuilds the generated pivot sheet.

- [ ] **Step 4: Render compact controls**

Update `Grid.tsx` to render an unobtrusive expand/collapse button in drillable pivot value cells. Style detail rows and the small toggle in `src/App.css`.

- [ ] **Step 5: Verify green**

Run: `pnpm exec playwright test tests/spreadsheet.spec.ts -g "expands pivot drilldown rows inline" --reporter=line`

Expected: pass.

### Task 4: Full Verification And PR

**Files:**
- Verify all changed files
- Commit and push branch `codex/inline-pivot-drilldown`

- [ ] **Step 1: Run build**

Run: `pnpm run build`

Expected: exit 0.

- [ ] **Step 2: Run unit suite**

Run: `pnpm test -- --reporter=dot`

Expected: exit 0.

- [ ] **Step 3: Run spreadsheet browser suite**

Run: `pnpm exec playwright test tests/spreadsheet.spec.ts --reporter=line`

Expected: exit 0.

- [ ] **Step 4: Commit and push**

Run: `git status -sb`, stage intended files, commit with `feat: add inline pivot drilldown`, then `git push -u origin codex/inline-pivot-drilldown`.

- [ ] **Step 5: Open draft PR**

Use `gh pr create --draft` with title `[codex] add inline pivot drilldown` and a Markdown body summarizing changes and verification.
