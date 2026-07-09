# DataTable Correctness and Workbook Session Foundation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Establish a correct, atomic, typed, controlled workbook foundation before any table surface exposes structural spreadsheet operations.

**Architecture:** Preserve `WorkbookModel` and HyperFormula authority while moving durable actions behind `WorkbookSession`. Fix confirmed formula, value, validation, sort, filter, history, and column-virtualization defects first; then migrate `App` handlers incrementally to typed commands and `useSyncExternalStore`.

**Tech Stack:** Node 22.13+, pnpm 11.7, TypeScript 6, React 19, HyperFormula 3.3, ExcelJS 4.4, Vitest 4, Testing Library, Playwright.

## Global Constraints

- Preserve existing spreadsheet appearance and no-prop standalone behavior.
- Keep HyperFormula derived from `WorkbookModel`; never make it canonical state.
- Reduce every workbook transaction atomically against the latest snapshot.
- Keep existing fixed XLSX shared-formula and native number/boolean regressions green.
- Use deterministic logical-weight history assertions instead of flaky heap-only tests.
- Do not add table metadata or remote data behavior in this plan.
- Do not begin public DataTable work until this plan's pressure gate passes.

---

### Task 1: Pin the supported executable environment

**Files:**
- Create: `.nvmrc`
- Modify: `package.json`
- Modify: `README.md`

**Interfaces:**
- Produces: Node 22.13.1 development default and an enforced Node floor of 22.13.

- [ ] **Step 1: Add the runtime files**

`.nvmrc` must contain:

```text
22.13.1
```

Add to `package.json` without changing `packageManager`:

```json
"engines": {
  "node": ">=22.13"
}
```

Also change the user-facing server scripts to:

```json
"dev": "vite --host 0.0.0.0",
"preview": "vite preview --host 0.0.0.0"
```

Browser automation may use `127.0.0.1` internally, but every manual handoff must report the machine IP and port.

- [ ] **Step 2: Document the exact activation command**

Add setup instructions for both version-manager and portable environments:

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

The portable fallback is required in the current execution host, which has Node 20 and no nvm/fnm/volta/mise/asdf. It runs the pinned `node@22.13.1` binary and places its directory first for child pnpm/Vite/Playwright processes; do not run pnpm 11 under the system Node 20.

- [ ] **Step 3: Verify the environment and baseline**

Run:

```bash
node --version
corepack pnpm --version
corepack pnpm test
corepack pnpm run build
```

Expected: Node `v22.13.1`, pnpm `11.7.0`, and passing baseline tests/build. Node 20 is not accepted because pnpm 11 requires `node:sqlite`.

- [ ] **Step 4: Commit the environment floor**

```bash
git add .nvmrc package.json README.md
git commit -m "chore: pin supported Node runtime"
```

### Task 2: Lock in already-correct baseline behavior

**Files:**
- Modify: `src/lib/workbook.test.ts`
- Modify: `src/lib/xlsx.test.ts`
- Modify: `src/lib/formulaEngine.perf.test.ts`

**Interfaces:**
- Produces: green guards for previously fixed audit findings. Live defects are introduced red and made green only inside their owning Tasks 3, 4, 5, 7, and 10 so no commit intentionally leaves the branch red.

- [ ] **Step 1: Add passing guards for fixed behavior**

Keep or strengthen assertions that native model numbers and booleans round-trip through XLSX, real shared-formula slaves translate to their own formulas, history depth never exceeds 100, and cell-only changes update HyperFormula incrementally.

- [ ] **Step 2: Run the guards**

Run:

```bash
corepack pnpm exec vitest run src/lib/workbook.test.ts src/lib/xlsx.test.ts src/lib/formulaEngine.perf.test.ts
```

Expected: PASS. If any guard fails, restore the baseline before continuing.

- [ ] **Step 3: Commit the green baseline guards**

```bash
git add src/lib/workbook.test.ts src/lib/xlsx.test.ts src/lib/formulaEngine.perf.test.ts
git commit -m "test: lock in workbook correctness baseline"
```

### Task 3: Parse and preserve typed spreadsheet input

**Files:**
- Create: `src/core/values/parseCellInput.ts`
- Create: `src/core/values/parseCellInput.test.ts`
- Create: `src/core/values/excelDate.ts`
- Create: `src/core/values/excelDate.test.ts`
- Modify: `src/types.ts`
- Modify: `src/lib/displayFormat.ts`
- Modify: `src/lib/xlsx.ts`
- Modify: `src/App.tsx`

**Interfaces:**
- Produces: `ParsedCellInput`, `parseCellInput`, Excel serial conversion, and `dateTime` display format.

- [ ] **Step 1: Write the parser tests**

```ts
expect(parseCellInput("")).toMatchObject({ kind: "blank", stored: null });
expect(parseCellInput("=A1*2")).toMatchObject({ kind: "formula", stored: "=A1*2", formula: "=A1*2" });
expect(parseCellInput("1.25e3")).toMatchObject({ kind: "number", stored: 1250 });
expect(parseCellInput("FALSE")).toMatchObject({ kind: "boolean", stored: false });
expect(parseCellInput("2026-01-15")).toMatchObject({ kind: "date", inferredNumberFormat: "date" });
expect(parseCellInput("2026-01-15T12:00:00Z")).toMatchObject({ kind: "dateTime", inferredNumberFormat: "dateTime" });
expect(parseCellInput("'00123")).toMatchObject({ kind: "text", stored: "00123" });
expect(parseCellInput("1,23")).toMatchObject({ kind: "text", stored: "1,23" });
```

Add the owning App regression `it("stores typed grid input as number, boolean, date, and date-time values")` in this step; do not commit it before this task.

- [ ] **Step 2: Run and verify failure**

Run: `corepack pnpm exec vitest run src/core/values`

Expected: FAIL because typed parsing is missing.

- [ ] **Step 3: Implement the typed parser**

```ts
export type ParsedCellInput = {
  raw: string;
  kind: "blank" | "formula" | "number" | "boolean" | "date" | "dateTime" | "text";
  stored: CellContent;
  formula?: string;
  inferredNumberFormat?: CellFormat["numberFormat"];
};

export function parseCellInput(raw: string): ParsedCellInput {
  if (raw === "") return { raw, kind: "blank", stored: null };
  if (raw.startsWith("'")) return { raw, kind: "text", stored: raw.slice(1) };
  if (raw.startsWith("=")) return { raw, kind: "formula", stored: raw, formula: raw };
  if (/^(TRUE|FALSE)$/i.test(raw)) return { raw, kind: "boolean", stored: raw.toUpperCase() === "TRUE" };
  const temporal = parseExcelTemporalInput(raw);
  if (temporal) return { raw, kind: temporal.kind, stored: temporal.serial, inferredNumberFormat: temporal.kind };
  const number = parseUnambiguousNumber(raw);
  if (number !== null) return { raw, kind: "number", stored: number };
  return { raw, kind: "text", stored: raw };
}
```

Use Excel's 1900 date system, including the compatibility leap-day offset. Accept only unambiguous US/ISO formats configured by the project.

- [ ] **Step 4: Wire parsing into cell and formula-bar commits**

`Grid` continues to edit raw strings, but App/session commit paths store `parseCellInput(raw).stored` and apply the inferred number format only when the target has no explicit format.

- [ ] **Step 5: Extend native XLSX date and date-time tests**

Assert import-export-import preserves typed serial values and matching formats for midnight dates and date-times.

- [ ] **Step 6: Run focused and existing XLSX tests**

Run: `corepack pnpm exec vitest run src/core/values src/lib/xlsx.test.ts src/lib/displayFormat.test.ts src/components/Grid.test.tsx src/App.test.tsx`

Expected: PASS.

- [ ] **Step 7: Commit typed input**

```bash
git add src/core/values src/types.ts src/lib/displayFormat.ts src/lib/xlsx.ts src/lib/xlsx.test.ts src/App.tsx src/App.test.tsx
git commit -m "fix: preserve typed spreadsheet input values"
```

### Task 4: Replace regex-only structural formula rewriting

**Files:**
- Modify: `src/lib/formulaReferences.ts`
- Modify: `src/lib/formulaReferences.test.ts`
- Modify: `src/lib/workbook.ts`
- Modify: `src/lib/workbook.test.ts`

**Interfaces:**
- Produces: `rewriteFormulaForStructure(formula, context)` used for every workbook structural edit.

- [ ] **Step 1: Expand token-aware tests**

Cover double-quoted strings and escaped quotes; single-quoted sheet names and doubled apostrophes; `LOG10`; `1E5`; identifiers containing digits; absolute/relative references; shrinking ranges; local and sheet-qualified references; and formulas on other sheets that target the edited sheet.

Add the owning regressions `it("does not rewrite LOG10, scientific notation, or quoted A1-like text during row insertion")` and `it("updates formulas on other sheets that reference the structurally edited sheet")` in this step.

- [ ] **Step 2: Run and verify the characterization failures remain**

Run: `corepack pnpm exec vitest run src/lib/formulaReferences.test.ts src/lib/workbook.test.ts`

Expected: targeted structural cases FAIL.

- [ ] **Step 3: Implement token-aware transformation**

```ts
export type FormulaStructureContext = {
  formulaSheetName: string;
  editedSheetName: string;
  axis: "row" | "column";
  mode: "insert" | "delete";
  index: number;
  count: number;
};

export function rewriteFormulaForStructure(formula: string, context: FormulaStructureContext): string {
  return tokenizeFormulaReferences(formula).map((token) =>
    token.kind === "reference" && referenceTargetsEditedSheet(token, context)
      ? shiftReferenceToken(token, context)
      : token.raw
  ).join("");
}
```

The tokenizer must never classify quoted text, function names, scientific notation, or identifiers as references.

- [ ] **Step 4: Rewrite formulas across every sheet**

Update structural operations so cells and metadata move only on the edited sheet, while formulas on all sheets pass through the structural reference service using formula-sheet and edited-sheet context.

- [ ] **Step 5: Run formula and workbook tests**

Run: `corepack pnpm exec vitest run src/lib/formulaReferences.test.ts src/lib/workbook.test.ts src/lib/formulaEngine.test.ts`

Expected: PASS.

- [ ] **Step 6: Commit structural formula safety**

```bash
git add src/lib/formulaReferences.ts src/lib/formulaReferences.test.ts src/lib/workbook.ts src/lib/workbook.test.ts
git commit -m "fix: make formula structure edits token aware"
```

### Task 5: Use evaluated typed values for validation, sorting, and filtering

**Files:**
- Modify: `src/lib/formulaEngine.ts`
- Modify: `src/lib/formulaEngine.test.ts`
- Modify: `src/lib/validation.ts`
- Modify: `src/lib/validation.test.ts`
- Modify: `src/lib/filters.ts`
- Modify: `src/lib/filters.test.ts`
- Modify: `src/lib/workbook.ts`
- Modify: `src/lib/workbook.test.ts`
- Modify: `src/components/Grid.tsx`

**Interfaces:**
- Produces: `ComputedCellValue`, `FormulaEngine.getComputedValue`, `ValidationCandidate`, and resolver-backed `sortRange`.

- [ ] **Step 1: Add computed-value tests**

```ts
expect(engine.getComputedValue(sheetId, "A1")).toBe(10);
expect(validateCellCandidate({ raw: "=5+5", parsed: "=5+5", evaluated: 10, formula: "=5+5" }, rule).valid).toBe(true);
expect(matchesFilterValue("", { operator: "equals", value: "0" })).toBe(false);
```

Add a sort test where formula strings are identical in shape but evaluated results require a different order, then assert moved relative formulas still evaluate correctly.

Add the owning formula-validation, blank-versus-zero filter, and evaluated-sort regressions in this step; do not commit them earlier.

- [ ] **Step 2: Run and verify failure**

Run: `corepack pnpm exec vitest run src/lib/formulaEngine.test.ts src/lib/validation.test.ts src/lib/filters.test.ts src/lib/workbook.test.ts`

Expected: new cases FAIL.

- [ ] **Step 3: Add computed and validation types**

```ts
export type ComputedCellValue = CellContent | { kind: "error"; code: string };
export type ValidationCandidate = { raw: string; parsed: CellContent; evaluated: ComputedCellValue; formula?: string };
```

List/text-length validation uses parsed text; numeric validation uses the evaluated number when a formula is present; formula errors are invalid with an explicit message.

- [ ] **Step 4: Make sort resolver-backed and formula-safe**

```ts
sortRange(workbook, sheetId, range, {
  direction,
  sortColumn,
  readValue: (address) => formulaEngine.getComputedValue(sheetId, address)
});
```

Translate relative formulas by the row movement through the same token-aware formula service.

- [ ] **Step 5: Make filter blank/null semantics explicit**

Return `null` before numeric conversion for blank strings. Add typed date, boolean, error, and formula-empty matching. Move AutoFilter distinct-choice collection behind menu-open state and cap displayed distinct choices with search.

- [ ] **Step 6: Run focused and UI suites**

Run: `corepack pnpm exec vitest run src/lib/formulaEngine.test.ts src/lib/validation.test.ts src/lib/filters.test.ts src/lib/workbook.test.ts src/components/Grid.test.tsx src/App.test.tsx`

Expected: PASS.

- [ ] **Step 7: Commit evaluated operations**

```bash
git add src/lib/formulaEngine.ts src/lib/formulaEngine.test.ts src/lib/validation.ts src/lib/validation.test.ts src/lib/filters.ts src/lib/filters.test.ts src/lib/workbook.ts src/lib/workbook.test.ts src/components/Grid.tsx
git commit -m "fix: use evaluated values for workbook operations"
```

### Task 6: Add count- and weight-bounded workbook history

**Files:**
- Create: `src/core/workbook/history.ts`
- Create: `src/core/workbook/history.test.ts`
- Modify: `src/lib/workbook.ts`
- Modify: `src/types.ts`

**Interfaces:**
- Produces: `WorkbookHistory`, `commitWorkbookHistory`, `undoWorkbookHistory`, `redoWorkbookHistory`, and retention statistics.

- [ ] **Step 1: Write deterministic retention tests**

Assert at most 100 snapshots, total estimated weight at most 2,000,000 logical entries, valid undo/redo after trimming, future clearing after a new edit, and structural identity sharing for untouched sheets.

- [ ] **Step 2: Run and verify failure**

Run: `corepack pnpm exec vitest run src/core/workbook/history.test.ts`

Expected: FAIL because weighted history is missing.

- [ ] **Step 3: Implement weighted trimming**

```ts
const DEFAULT_MAX_ENTRIES = 100;
const DEFAULT_MAX_WEIGHT = 2_000_000;

export function commitWorkbookHistory(history: WorkbookHistory, workbook: WorkbookModel): WorkbookHistory {
  const past = trimHistory([...history.past, weighted(history.present)], history.limits);
  return { ...history, past, present: workbook, future: [] };
}
```

Estimate weight from cell and metadata-map entry counts. Expose retained count and weight for tests and diagnostics.

- [ ] **Step 4: Migrate old helpers and run tests**

Run: `corepack pnpm exec vitest run src/core/workbook/history.test.ts src/lib/workbook.test.ts src/App.test.tsx`

Expected: PASS.

- [ ] **Step 5: Commit bounded history**

```bash
git add src/core/workbook/history.ts src/core/workbook/history.test.ts src/lib/workbook.ts src/types.ts
git commit -m "perf: bound workbook snapshot history"
```

### Task 7: Add typed commands and atomic WorkbookSession

**Files:**
- Create: `src/core/commands/types.ts`
- Create: `src/core/workbook/commands.ts`
- Create: `src/core/workbook/commands.test.ts`
- Create: `src/core/workbook/WorkbookSession.ts`
- Create: `src/core/workbook/WorkbookSession.test.ts`
- Create: `src/core/workbook/WorkbookSession.property.test.ts`
- Modify: `package.json`
- Modify: `pnpm-lock.yaml`

**Interfaces:**
- Produces: `CommandEnvelope`, `CommandResult`, `WorkbookCommand`, `WorkbookSession`, and `createWorkbookSession`.

Define the complete pre-table command union in `src/core/workbook/commands.ts`; the workbook-table plan extends it with `StructuredTableCommand`:

```ts
export type WorkbookCommand =
  | { type: "transaction"; commands: readonly WorkbookCommand[] }
  | { type: "selection.set"; selection: CellRange }
  | { type: "cell.set"; sheetId: string; address: string; input: string }
  | { type: "cell.comment.set"; sheetId: string; address: string; comment: string | null }
  | { type: "cell.hyperlink.set"; sheetId: string; address: string; hyperlink: string | null }
  | { type: "range.clear"; sheetId: string; range: CellRange; mode: "contents" | "formats" | "comments" | "hyperlinks" | "all" }
  | { type: "range.format"; sheetId: string; range: CellRange; format: Partial<CellFormat> }
  | { type: "range.borders"; sheetId: string; range: CellRange; preset: BorderPreset }
  | { type: "range.validation.set"; sheetId: string; range: CellRange; rule: DataValidationRule }
  | { type: "range.validation.clear"; sheetId: string; range: CellRange }
  | { type: "range.conditionalFormat.add"; sheetId: string; range: CellRange; rule: ConditionalFormatRule }
  | { type: "range.conditionalFormat.remove"; sheetId: string; ruleId: string }
  | { type: "range.readOnly.set"; sheetId: string; range: CellRange; readOnly: boolean }
  | { type: "range.merge" | "range.unmerge"; sheetId: string; range: CellRange }
  | { type: "range.fill"; sheetId: string; range: CellRange; direction: "down" | "right" }
  | { type: "range.autoFill"; sheetId: string; source: CellRange; target: CellRange }
  | { type: "range.sort"; sheetId: string; range: CellRange; direction: "asc" | "desc"; sortColumn?: number }
  | { type: "range.removeDuplicates"; sheetId: string; range: CellRange }
  | { type: "clipboard.paste"; sheetId: string; target: CellCoord; payload: RichClipboardRange; mode: RichPasteMode }
  | { type: "clipboard.move"; sourceSheetId: string; source: CellRange; targetSheetId: string; target: CellCoord }
  | { type: "rows.insert" | "rows.delete"; sheetId: string; index: number; count: number }
  | { type: "columns.insert" | "columns.delete"; sheetId: string; index: number; count: number }
  | { type: "rows.resize"; sheetId: string; rows: readonly number[]; height: number }
  | { type: "columns.resize"; sheetId: string; columns: readonly number[]; width: number }
  | { type: "rows.hidden.set"; sheetId: string; rows: readonly number[]; hidden: boolean }
  | { type: "columns.hidden.set"; sheetId: string; columns: readonly number[]; hidden: boolean }
  | { type: "sheet.add"; name?: string }
  | { type: "sheet.rename"; sheetId: string; name: string }
  | { type: "sheet.duplicate" | "sheet.delete" | "sheet.activate"; sheetId: string }
  | { type: "sheet.move"; sheetId: string; targetIndex: number }
  | { type: "sheet.hidden.set"; sheetId: string; hidden: boolean }
  | { type: "sheet.tabColor.set"; sheetId: string; color: string }
  | { type: "sheet.freeze.set"; sheetId: string; rows: number; columns: number }
  | { type: "sheet.protection.set"; sheetId: string; protected: boolean }
  | { type: "sheet.filter.set"; sheetId: string; filter: SheetFilter }
  | { type: "sheet.filter.clear"; sheetId: string; column?: number }
  | { type: "sheet.chart.add"; sheetId: string; chart: SheetChart }
  | { type: "sheet.chart.delete"; sheetId: string; chartId: string }
  | { type: "namedRange.define"; namedRange: NamedRange }
  | { type: "namedRange.remove"; name: string }
  | { type: "history.undo" | "history.redo" }
  | { type: "persistence.status"; status: "idle" | "saving" | "failed"; message?: string }
  | { type: "workbook.replace"; workbook: WorkbookModel; history: "commit" | "reset" };
```

Every variant is serializable data. Pivot output and multi-range paste compile to one `transaction` of the primitives above; reducers never accept arbitrary callbacks.

- [ ] **Step 1: Install the property-test dependency**

Run: `corepack pnpm add -D --save-exact fast-check@4.9.0`

Expected: dependency and lockfile update before any `fast-check` import.

- [ ] **Step 2: Write atomicity and transaction tests**

Cover two synchronous dispatches, one-notification transactions, child rollback, no-op suppression, expected-revision conflict, recalculation before publication, destroy cleanup, and seeded edit/undo property sequences.

Every `WorkbookSession.property.test.ts` assertion uses the same explicit deterministic run parameters:

```ts
const replayPath = process.env.FC_REPLAY_PATH;

fc.assert(editUndoProperty, {
  seed: 20260709,
  path: replayPath,
  numRuns: 1000,
  endOnFailure: true,
});
```

On failure, preserve fast-check's `Seed: 20260709` and `Path: "<shrunk-path>"` lines in the Vitest output/artifact. Replay the exact shrink locally with `FC_REPLAY_PATH='<shrunk-path>' corepack pnpm exec vitest run src/core/workbook/WorkbookSession.property.test.ts`; the replay must still use seed `20260709`, `numRuns: 1000`, and `endOnFailure: true`. Add a small test around the options builder so an accidental seed/run-count change fails review rather than silently weakening the pressure test.

- [ ] **Step 3: Run and verify failure**

Run: `corepack pnpm exec vitest run src/core/workbook/WorkbookSession.test.ts src/core/workbook/WorkbookSession.property.test.ts`

Expected: FAIL because the session is missing.

- [ ] **Step 4: Define shared command lifecycle types**

```ts
export type CommandEnvelope<TIntent> = { id: string; intent: TIntent; transactionId?: string; expectedRevision?: string };
export type TableIssue = { code: string; message: string; sheetId?: string; address?: string };
export type CommandResult<TCurrent = unknown> =
  | { status: "committed"; revision: string; changed?: boolean }
  | { status: "pending"; operationId: string }
  | { status: "rejected"; reason: "validation" | "permission" | "unsupported"; issues?: readonly TableIssue[] }
  | { status: "conflict"; revision: string; current: TCurrent };
```

- [ ] **Step 5: Implement WorkbookSession around current workbook helpers**

```ts
export type WorkbookSnapshot = {
  workbook: WorkbookModel;
  selection: CellRange;
  revision: string;
  canUndo: boolean;
  canRedo: boolean;
  persistence: { status: "idle" | "saving" | "failed"; message?: string };
};

export type WorkbookCommandResult = CommandResult<WorkbookSnapshot>;

export type WorkbookDiagnosticEvent = {
  category: "command" | "validation" | "persistence" | "performance";
  commandId?: string;
  durationMs?: number;
  metadata: Readonly<Record<string, string | number | boolean | null>>;
};

export type CreateWorkbookSessionOptions = {
  workbook: WorkbookModel;
  selection?: CellRange;
  formulaEngineFactory?: typeof createFormulaEngine;
  now?: () => number;
  createCommandId?: () => string;
  history?: { maxEntries?: number; maxWeight?: number };
  onDiagnostic?: (event: WorkbookDiagnosticEvent) => void;
};

export interface WorkbookSession {
  getSnapshot(): WorkbookSnapshot;
  getCellEvaluation(sheetId: string, address: string): ComputedCellValue;
  subscribe(listener: () => void): () => void;
  subscribeDiagnostics(listener: (event: WorkbookDiagnosticEvent) => void): () => void;
  dispatch(command: WorkbookCommand | CommandEnvelope<WorkbookCommand>): WorkbookCommandResult;
  replaceWorkbook(workbook: WorkbookModel, options?: { history?: "reset" | "preserve"; origin?: "external" | "import" }): WorkbookCommandResult;
  destroy(): void;
}

export function createWorkbookSession(options: CreateWorkbookSessionOptions): WorkbookSession;
```

Dispatch reduces against private current state, updates FormulaEngine, then publishes. A transaction reduces children into a temporary workbook and commits once only if every child succeeds.

`dispatch` distinguishes a plain command from an envelope with `"intent" in value`; only the envelope form can carry `expectedRevision`. A plain command is wrapped with a per-session command ID from `options.createCommandId ?? crypto.randomUUID`, never a module counter. Add compile-time/runtime tests for deterministic injected IDs, uniqueness across two sessions, and expected-revision mismatch returning `conflict` without mutating state.

Time dispatch with `options.now ?? performance.now` when available (fall back to a monotonic zero-duration measurement in runtimes without either) and emit one sanitized diagnostic containing command ID/type, duration, outcome category, and affected sheet/range/cell counts. Publish it to both `options.onDiagnostic` and `subscribeDiagnostics` listeners. Never include stored/display values, formulas, raw input, comments, hyperlinks, validation payloads, clipboard contents, workbook objects, or storage data. Catch diagnostic-callback/listener exceptions independently so one host cannot suppress another. Add tests for committed, rejected, conflict, transaction, persistence failure, diagnostic subscribe/unsubscribe, destroy cleanup, and callback-throws paths.

- [ ] **Step 6: Run session and property tests**

Run: `corepack pnpm exec vitest run src/core/workbook`

Expected: PASS.

- [ ] **Step 7: Commit atomic sessions**

```bash
git add package.json pnpm-lock.yaml src/core/commands src/core/workbook/commands.ts src/core/workbook/commands.test.ts src/core/workbook/WorkbookSession.ts src/core/workbook/WorkbookSession.test.ts src/core/workbook/WorkbookSession.property.test.ts
git commit -m "feat: add atomic workbook session"
```

### Task 8: Add useWorkbookSession and controlled Spreadsheet

**Files:**
- Create: `src/react/useWorkbookSession.ts`
- Create: `src/react/useWorkbookSession.test.tsx`
- Create: `src/react/browserWorkbookStorage.ts`
- Create: `src/react/WorkbookSessionContext.tsx`
- Create: `src/react/WorkbookSessionContext.test.tsx`
- Create: `src/react/Spreadsheet.tsx`
- Create: `src/core/workbook/services.ts`
- Create: `src/Spreadsheet.controlled.test.tsx`
- Create: `src/react/Spreadsheet.types.test.tsx`
- Modify: `src/App.tsx`
- Modify: `src/index.ts`
- Modify: `src/main.tsx`

**Interfaces:**
- Consumes: `WorkbookSession`.
- Produces: mutually exclusive `SpreadsheetProps`, `useWorkbookSession`, `WorkbookSessionProvider`, `useWorkbookSessionContext`, and controlled/no-prop modes.

- [ ] **Step 1: Write ownership, public-prop, and controlled-mode tests**

Cover session ownership, controlled rerender, uncontrolled default, `storage={false}`, shared subscribers, StrictMode, persistence failure, cleanup, feature flags, service injection, event delivery, and root-scoped theme overrides. In `Spreadsheet.types.test.tsx`, add positive compile-time cases showing that every ownership mode accepts the common props and `@ts-expect-error` cases for every invalid ownership pair (`session` plus `workbook`, `session` plus `defaultWorkbook`, `workbook` without `onWorkbookChange`, controlled plus `defaultWorkbook`, and uncontrolled plus `onWorkbookChange`).

- [ ] **Step 2: Run and verify failure**

Run: `corepack pnpm exec vitest run src/react/useWorkbookSession.test.tsx src/react/Spreadsheet.types.test.tsx src/Spreadsheet.controlled.test.tsx`

Run: `corepack pnpm exec tsc --noEmit`

Expected: FAIL because controlled APIs do not exist.

- [ ] **Step 3: Add the mutually exclusive props**

```ts
export type WorkbookStorage = {
  load(): WorkbookModel | null | Promise<WorkbookModel | null>;
  save(workbook: WorkbookModel): void | Promise<void>;
  clear?(): void | Promise<void>;
};

export type WorkbookFeatureConfiguration = Readonly<{
  toolbar?: boolean;
  formulaBar?: boolean;
  sheetTabs?: boolean;
  import?: boolean;
  export?: boolean;
  charts?: boolean;
  structuredTables?: boolean;
  googleSheets?: boolean;
}>;

export type WorkbookThemeToken =
  | "font-family"
  | "font-size"
  | "surface"
  | "surface-muted"
  | "text"
  | "text-muted"
  | "border"
  | "accent"
  | "accent-contrast"
  | "selection"
  | "danger";

// Define these contracts in the DOM-free src/core/workbook/services.ts module.
export type WorkbookImportPayload =
  | { kind: "bytes"; bytes: Uint8Array; fileName?: string }
  | { kind: "text"; text: string; fileName?: string };

export type WorkbookExportArtifact = Readonly<{
  bytes: Uint8Array;
  mediaType: string;
  fileName: string;
}>;

export type WorkbookServiceOptions = Readonly<Record<string, string | number | boolean | null>>;

export interface WorkbookImporter {
  import(payload: WorkbookImportPayload, options?: WorkbookServiceOptions): WorkbookModel | Promise<WorkbookModel>;
}

export interface WorkbookExporter {
  export(workbook: WorkbookModel, options?: WorkbookServiceOptions): WorkbookExportArtifact | Promise<WorkbookExportArtifact>;
}

export type SpreadsheetServices = Readonly<{
  formulaEngineFactory?: typeof createFormulaEngine;
  importers?: Readonly<Record<string, WorkbookImporter>>;
  exporters?: Readonly<Record<string, WorkbookExporter>>;
  now?: () => number;
  createCommandId?: () => string;
  googleTokenProviderFactory?: (clientId: string) => TokenProvider;
}>;

export type WorkbookChangeEvent = Readonly<{
  workbook: WorkbookModel;
  revision: string;
  previousRevision: string;
  origin: "command" | "undo" | "redo" | "import" | "external" | "storage";
  commandId?: string;
}>;

export type SpreadsheetErrorEvent = Readonly<{
  code: string;
  message: string;
  recoverable: boolean;
}>;

export type SpreadsheetCommonProps = Readonly<{
  className?: string;
  style?: React.CSSProperties;
  features?: WorkbookFeatureConfiguration;
  services?: SpreadsheetServices;
  theme?: Partial<Record<WorkbookThemeToken, string>>;
  onDiagnostic?: (event: WorkbookDiagnosticEvent) => void;
  onCommandResult?: (event: Readonly<{ command: WorkbookCommand; result: WorkbookCommandResult }>) => void;
  onWorkbookChangeEvent?: (event: WorkbookChangeEvent) => void;
  onError?: (event: SpreadsheetErrorEvent) => void;
}>;

export type SpreadsheetProps = SpreadsheetCommonProps & (
  | { session: WorkbookSession; workbook?: never; defaultWorkbook?: never; onWorkbookChange?: never; storage?: never }
  | { workbook: WorkbookModel; onWorkbookChange(workbook: WorkbookModel): void; session?: never; defaultWorkbook?: never; storage?: WorkbookStorage | false }
  | { defaultWorkbook?: WorkbookModel; session?: never; workbook?: never; onWorkbookChange?: never; storage?: WorkbookStorage | false }
);
```

`WorkbookImporter` and `WorkbookExporter` are named, platform-neutral interfaces whose requests and results are serializable metadata plus `Uint8Array`; neither contract may mention `File`, `Blob`, `HTMLElement`, or another browser global, and browser adapters belong in `src/react`. Validate blank/reserved registry keys once when creating the component. Pass `formulaEngineFactory`, `now`, and `createCommandId` into internally created sessions; a supplied session remains authoritative for evaluation, time, and command IDs, while the `services` import/export registry and `googleTokenProviderFactory` remain available to the host-facing toolbar. Never silently replace services owned by a supplied session. Catch service exceptions and emit a sanitized `SpreadsheetErrorEvent` without including workbook values, formulas, credentials, source payloads, or raw thrown messages.

Apply `features` as capability visibility/enablement only: disabling a surface must not delete workbook state. Translate `theme` entries to root-local `--js-spreadsheet-*` custom properties, reject unknown keys through the type, and let an explicit `style` custom property win if it sets the same root property. Deliver `onCommandResult` once for every public dispatch outcome, including rejection/conflict; deliver `onWorkbookChangeEvent` only after a committed workbook revision; forward sanitized session diagnostics to `onDiagnostic`; catch host callback exceptions so one callback cannot prevent publication to subscribers. `onWorkbookChange` remains the controlled-value acknowledgement callback and is not replaced by `onWorkbookChangeEvent`.

Use `useSyncExternalStore`. Never destroy a supplied session; always destroy an internally created session. Controlled mode does not touch browser storage unless explicitly supplied. No-prop mode preserves current storage behavior.

`persistence.status` updates only `WorkbookSnapshot.persistence`; it creates no workbook revision or history entry. Promise-based storage hydration records the initial session revision before `load()`. Apply a loaded workbook with `history: "reset"` only if the session is still at that revision and uncontrolled; if the user edited or props changed first, discard the stale load. Controlled `workbook` mode never lets `load()` replace host state, even when storage is supplied; it may save acknowledged controlled revisions. Serialize async saves in revision order, coalesce superseded queued revisions, ignore completions after destroy, and dispatch `persistence.status` for saving/failure/recovery. Add deferred-promise tests for stale load, out-of-order save completion, quota rejection, StrictMode remount, and unmount cleanup.

Create `WorkbookSessionContext.tsx` with `WorkbookSessionProvider({ session, children })` and `useWorkbookSessionContext()`. The provider never owns or destroys its session. The hook throws `useWorkbookSessionContext must be used inside WorkbookSessionProvider` when absent. Test nested provider isolation, StrictMode, live snapshot subscription, and no destroy on provider unmount.

- [ ] **Step 4: Export the React surface without moving the App monolith yet**

`src/react/Spreadsheet.tsx` re-exports the controlled component from `App.tsx` and its props. Keep `App` as the standalone wrapper so existing tests remain stable.

The embeddable workbook root must already render `class="js-spreadsheet-root js-spreadsheet-workbook"` and `data-js-spreadsheet-root="workbook"` in this task. Preserve existing compatibility classes inside it. Workbook-table styles/tests depend on this contract; the packaging plan later consolidates tokens and performs the complete leak audit.

- [ ] **Step 5: Run controlled and existing App suites**

Run: `corepack pnpm exec vitest run src/react/useWorkbookSession.test.tsx src/react/Spreadsheet.types.test.tsx src/Spreadsheet.controlled.test.tsx src/App.test.tsx`

Run: `corepack pnpm exec tsc --noEmit`

Expected: PASS.

- [ ] **Step 6: Commit controlled embedding**

```bash
git add src/core/workbook/services.ts src/react src/Spreadsheet.controlled.test.tsx src/App.tsx src/index.ts src/main.tsx
git commit -m "feat: support controlled spreadsheet sessions"
```

### Task 9: Route every durable App mutation through commands

**Files:**
- Modify: `src/App.tsx`
- Modify: `src/core/workbook/commands.ts`
- Modify: `src/core/workbook/commands.test.ts`
- Modify: `src/App.test.tsx`

**Interfaces:**
- Consumes: `WorkbookSession.dispatch`.
- Produces: no direct durable `setHistory` or stale closure commit path.

- [ ] **Step 1: Add the owning stale-closure regression**

Add `it("retains both synchronous workbook commits in one React batch")` to `src/App.test.tsx`, run only that test, and confirm it fails for the current non-functional `setHistory` path.

- [ ] **Step 2: Migrate cell, format, validation, conditional-format, and filter handlers**

Replace handlers with exact typed commands and keep their existing status messages derived from `CommandResult`. Leave the now-unused helper definitions in place until Step 5 so their deletion is an explicit reviewed commit. Run their focused App tests, then commit:

```bash
git add src/App.tsx src/App.test.tsx src/core/workbook/commands.ts src/core/workbook/commands.test.ts
git diff --cached --check
git commit -m "refactor: route cell and range edits through commands"
```

- [ ] **Step 3: Migrate sort, fill, row/column, protection, merge, and sizing handlers**

Use transactions when one user action changes multiple workbook concerns. Run workbook, Grid, and matching App tests, then commit:

```bash
git add src/App.tsx src/App.test.tsx src/core/workbook/commands.ts src/core/workbook/commands.test.ts
git diff --cached --check
git commit -m "refactor: route structural edits through commands"
```

- [ ] **Step 4: Migrate clipboard, sheets, imports, and workbook replacement**

Reset stale cut state on workbook/sheet replacement, remove CSV argument spreading, and make imports dispatch against current session state. Replace the module-level Google token provider with an injected instance service. Commit:

```bash
git add src/App.tsx src/App.test.tsx src/core/workbook/commands.ts src/core/workbook/commands.test.ts src/core/workbook/services.ts src/react/Spreadsheet.tsx
git diff --cached --check
git commit -m "refactor: route workbook lifecycle through commands"
```

- [ ] **Step 5: Delete stale closure helpers and prove no bypass remains**

Remove `commitWorkbook` and `applyHistoryTransition`. Search for direct `setHistory` and durable workbook mutation bypasses; only session internals and controlled external replacement may remain. Stage and commit the helper deletions and their regression assertions rather than leaving them in the worktree:

```bash
git add src/App.tsx src/App.test.tsx src/core/workbook/commands.ts src/core/workbook/commands.test.ts
git diff --cached --check
git commit -m "refactor: remove workbook mutation bypasses"
```

- [ ] **Step 6: Run full App and unit suites**

Run: `corepack pnpm test`

Expected: PASS.

### Task 10: Add reusable two-axis virtualization primitives

**Files:**
- Create: `src/core/viewport/axis.ts`
- Create: `src/core/viewport/axis.test.ts`
- Modify: `src/components/Grid.tsx`
- Modify: `src/components/Grid.test.tsx`

**Interfaces:**
- Produces: `measureAxis`, `findVisibleRange`, and bounded horizontal rendering used later by DataTable.

- [ ] **Step 1: Write axis and 10,000-column tests**

Test binary-search visibility, overscan, hidden columns, frozen-first-column retention, horizontal scroll, `ensureCellVisible`, selection overlays, editing, and bounded DOM count.

- [ ] **Step 2: Run and verify failure**

Run: `corepack pnpm exec vitest run src/core/viewport/axis.test.ts src/components/Grid.test.tsx`

Expected: the wide-column bound FAILS.

- [ ] **Step 3: Implement axis measurement and horizontal viewport state**

```ts
export function findVisibleRange(measurements: readonly AxisMeasurement[], start: number, end: number, overscan: number): AxisRange {
  const first = Math.max(0, lowerBoundByEnd(measurements, start) - overscan);
  const last = Math.min(measurements.length, upperBoundByStart(measurements, end) + overscan);
  return { first, last };
}
```

Render left/right spacers and the visible column window while retaining the full measurement array for hit testing and overlay geometry.

- [ ] **Step 4: Run viewport and Grid tests**

Run: `corepack pnpm exec vitest run src/core/viewport/axis.test.ts src/components/Grid.test.tsx`

Expected: PASS with bounded visible cells.

- [ ] **Step 5: Commit two-axis virtualization**

```bash
git add src/core/viewport src/components/Grid.tsx src/components/Grid.test.tsx
git commit -m "perf: virtualize spreadsheet columns"
```

### Task 11: Run the foundation pressure gate

**Files:**
- Create: `src/core/workbook/WorkbookSession.perf.test.ts`
- Modify: `docs/performance.md`

**Interfaces:**
- Produces: verified foundation for all later plans.

- [ ] **Step 1: Add 100,000-row command and history benchmarks**

Measure cell edit, formula recalculation, 100-command burst, weighted history trimming, and subscriber publication order in production-equivalent conditions. Keep correctness assertions active when timing assertions are disabled explicitly.

- [ ] **Step 2: Run focused foundation verification**

Run:

```bash
corepack pnpm exec vitest run src/core src/lib/workbook.test.ts src/lib/formulaReferences.test.ts src/lib/validation.test.ts src/lib/filters.test.ts src/lib/formulaEngine.test.ts src/lib/formulaEngine.perf.test.ts src/components/Grid.test.tsx src/App.test.tsx
corepack pnpm test
corepack pnpm run build
corepack pnpm exec playwright test --project=chromium
```

Expected: all commands pass.

- [ ] **Step 3: Update performance documentation**

Record the pinned Node version, benchmark environment, typed-input behavior, history limits, atomic command results, and horizontal virtualization limits.

- [ ] **Step 4: Commit the foundation gate**

```bash
git add src/core/workbook/WorkbookSession.perf.test.ts docs/performance.md
git commit -m "test: pressure workbook session foundation"
```
