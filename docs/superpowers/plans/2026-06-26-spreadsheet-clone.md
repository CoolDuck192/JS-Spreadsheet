# JavaScript Spreadsheet Clone Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build a standalone JavaScript spreadsheet app with an Excel-like UI, formula support, sheet management, CSV import/export, persistence, and verification.

**Architecture:** Use Vite, React, and TypeScript for the browser app. Keep spreadsheet logic in focused TypeScript modules: addressing, CSV, workbook state, persistence, formula engine, and React components. HyperFormula powers formula evaluation behind a small adapter so the UI can treat it as a spreadsheet engine.

**Tech Stack:** Vite, React, TypeScript, Vitest, Testing Library, Playwright, HyperFormula, lucide-react.

---

## File Structure

- `package.json`: scripts and dependencies.
- `index.html`: Vite entry document.
- `vite.config.ts`: Vite and Vitest config.
- `tsconfig.json`, `tsconfig.node.json`: TypeScript config.
- `playwright.config.ts`: browser smoke test config.
- `src/main.tsx`: React entrypoint.
- `src/App.tsx`: top-level app shell and workbook orchestration.
- `src/App.css`: spreadsheet layout and responsive styling.
- `src/types.ts`: workbook, sheet, selection, and cell types.
- `src/lib/addressing.ts`: A1/coordinate conversion and range utilities.
- `src/lib/csv.ts`: CSV parser and serializer.
- `src/lib/workbook.ts`: workbook creation, immutable edits, selection helpers, fill, copy, paste, undo/redo.
- `src/lib/persistence.ts`: local-storage save and restore.
- `src/lib/formulaEngine.ts`: HyperFormula wrapper.
- `src/components/Toolbar.tsx`: command toolbar.
- `src/components/FormulaBar.tsx`: active address and raw value editor.
- `src/components/Grid.tsx`: spreadsheet grid rendering and interactions.
- `src/components/SheetTabs.tsx`: sheet tab controls.
- `src/components/StatusBar.tsx`: current status and errors.
- `src/**/*.test.ts`: unit tests.
- `src/App.test.tsx`: React interaction tests.
- `tests/spreadsheet.spec.ts`: Playwright smoke test.

---

## Task 1: Project Scaffold

**Files:**
- Create: `package.json`
- Create: `index.html`
- Create: `vite.config.ts`
- Create: `tsconfig.json`
- Create: `tsconfig.node.json`
- Create: `playwright.config.ts`
- Create: `src/main.tsx`
- Create: `src/App.tsx`
- Create: `src/App.css`

- [ ] **Step 1: Create package and config files**

Add scripts:

```json
{
  "scripts": {
    "dev": "vite --host 127.0.0.1",
    "build": "tsc -b && vite build",
    "test": "vitest run",
    "test:watch": "vitest",
    "test:e2e": "playwright test"
  }
}
```

- [ ] **Step 2: Create minimal React entrypoint**

Render a placeholder `SpreadsheetApp` shell so Vite can boot.

- [ ] **Step 3: Install dependencies**

Run: `npm install`

- [ ] **Step 4: Verify scaffold**

Run: `npm run build`

Expected: TypeScript and Vite complete with exit code 0.

---

## Task 2: Addressing Module

**Files:**
- Create: `src/types.ts`
- Create: `src/lib/addressing.test.ts`
- Create: `src/lib/addressing.ts`

- [ ] **Step 1: Write failing addressing tests**

Cover:

```ts
expect(columnIndexToName(0)).toBe("A");
expect(columnIndexToName(25)).toBe("Z");
expect(columnIndexToName(26)).toBe("AA");
expect(parseCellAddress("BC23")).toEqual({ row: 22, column: 54 });
expect(formatCellAddress({ row: 4, column: 2 })).toBe("C5");
expect(parseRangeAddress("A1:C3")).toEqual({
  start: { row: 0, column: 0 },
  end: { row: 2, column: 2 }
});
```

- [ ] **Step 2: Run the tests and verify failure**

Run: `npm test -- src/lib/addressing.test.ts`

Expected: FAIL because `src/lib/addressing.ts` is missing.

- [ ] **Step 3: Implement addressing utilities**

Create column conversion, address parsing, range parsing, range normalization, and address iteration.

- [ ] **Step 4: Run tests and verify pass**

Run: `npm test -- src/lib/addressing.test.ts`

Expected: PASS.

---

## Task 3: CSV Module

**Files:**
- Create: `src/lib/csv.test.ts`
- Create: `src/lib/csv.ts`

- [ ] **Step 1: Write failing CSV tests**

Cover plain values, quoted fields, escaped quotes, blank cells, commas inside quotes, and line breaks inside quoted fields.

- [ ] **Step 2: Run the tests and verify failure**

Run: `npm test -- src/lib/csv.test.ts`

Expected: FAIL because parser functions are missing.

- [ ] **Step 3: Implement CSV parser and serializer**

Expose `parseCsv(text: string): string[][]` and `serializeCsv(rows: string[][]): string`.

- [ ] **Step 4: Run tests and verify pass**

Run: `npm test -- src/lib/csv.test.ts`

Expected: PASS.

---

## Task 4: Workbook State

**Files:**
- Create: `src/lib/workbook.test.ts`
- Create: `src/lib/workbook.ts`

- [ ] **Step 1: Write failing workbook tests**

Cover blank workbook creation, cell edits, clearing ranges, copy matrix, paste matrix, fill-down, fill-right, sheet add, rename, duplicate, delete, undo, and redo.

- [ ] **Step 2: Run the tests and verify failure**

Run: `npm test -- src/lib/workbook.test.ts`

Expected: FAIL because workbook functions are missing.

- [ ] **Step 3: Implement immutable workbook operations**

Expose helpers including `createBlankWorkbook`, `setCellContent`, `clearRange`, `copyRange`, `pasteMatrix`, `fillDown`, `fillRight`, `addSheet`, `renameSheet`, `duplicateSheet`, `deleteSheet`, `createHistory`, `commitHistory`, `undoHistory`, and `redoHistory`.

- [ ] **Step 4: Run tests and verify pass**

Run: `npm test -- src/lib/workbook.test.ts`

Expected: PASS.

---

## Task 5: Formula Engine

**Files:**
- Create: `src/lib/formulaEngine.test.ts`
- Create: `src/lib/formulaEngine.ts`

- [ ] **Step 1: Write failing formula tests**

Cover raw values, formulas, ranges, references, `SUM`, `AVERAGE`, `IF`, recalculation after dependency edit, and visible formula errors.

- [ ] **Step 2: Run the tests and verify failure**

Run: `npm test -- src/lib/formulaEngine.test.ts`

Expected: FAIL because the formula adapter is missing.

- [ ] **Step 3: Implement HyperFormula adapter**

Create `createFormulaEngine(workbook)` with `getDisplayValue(sheetId, address)`, `getRawContent(sheetId, address)`, and `rebuild(workbook)`.

- [ ] **Step 4: Run tests and verify pass**

Run: `npm test -- src/lib/formulaEngine.test.ts`

Expected: PASS.

---

## Task 6: Persistence

**Files:**
- Create: `src/lib/persistence.test.ts`
- Create: `src/lib/persistence.ts`

- [ ] **Step 1: Write failing persistence tests**

Cover saving, restoring valid workbook JSON, fallback on invalid JSON, and fallback on unsupported version.

- [ ] **Step 2: Run the tests and verify failure**

Run: `npm test -- src/lib/persistence.test.ts`

Expected: FAIL because persistence functions are missing.

- [ ] **Step 3: Implement persistence helpers**

Expose `saveWorkbook(storage, workbook)` and `loadWorkbook(storage)`.

- [ ] **Step 4: Run tests and verify pass**

Run: `npm test -- src/lib/persistence.test.ts`

Expected: PASS.

---

## Task 7: React Spreadsheet UI

**Files:**
- Modify: `src/App.tsx`
- Modify: `src/App.css`
- Create: `src/components/Toolbar.tsx`
- Create: `src/components/FormulaBar.tsx`
- Create: `src/components/Grid.tsx`
- Create: `src/components/SheetTabs.tsx`
- Create: `src/components/StatusBar.tsx`
- Create: `src/App.test.tsx`

- [ ] **Step 1: Write failing React interaction tests**

Cover editing `A1`, entering `=SUM(A1:A3)`, formula bar raw formula visibility, adding a sheet, renaming a sheet, and copy/paste via tabular clipboard text.

- [ ] **Step 2: Run tests and verify failure**

Run: `npm test -- src/App.test.tsx`

Expected: FAIL because UI behavior is missing.

- [ ] **Step 3: Implement components**

Build the toolbar, formula bar, grid, sheet tabs, status bar, and top-level app state. Wire keyboard shortcuts for navigation, edit commit/cancel, copy, paste, fill-down, fill-right, undo, redo, and clear.

- [ ] **Step 4: Run tests and verify pass**

Run: `npm test -- src/App.test.tsx`

Expected: PASS.

---

## Task 8: Browser Smoke Test And Final Verification

**Files:**
- Create: `tests/spreadsheet.spec.ts`
- Modify: `README.md`

- [ ] **Step 1: Write Playwright smoke test**

Test that the app loads, cells can be edited, formulas recalculate, a second sheet can be added, and CSV export button exists.

- [ ] **Step 2: Run Playwright test and verify failure if app behavior is incomplete**

Run: `npm run test:e2e`

Expected before implementation completion: FAIL on missing behavior. Expected after Task 7: PASS.

- [ ] **Step 3: Write README usage notes**

Document `npm install`, `npm run dev`, `npm test`, `npm run build`, and `npm run test:e2e`.

- [ ] **Step 4: Run full verification**

Run:

```bash
npm test
npm run build
npm run test:e2e
```

Expected: all commands exit 0.

---

## Spec Coverage Review

- Grid, formula bar, formulas, error display, copy/paste, fill-down, fill-right, undo/redo, sheet tabs, CSV import/export, persistence, and tests are covered by Tasks 2 through 8.
- XLSX is intentionally out of first implementation scope per the design.
- The current workspace is not a git repository, so commit steps are omitted unless a git repository is initialized later.

