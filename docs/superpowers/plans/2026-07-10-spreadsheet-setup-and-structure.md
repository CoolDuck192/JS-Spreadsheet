# Spreadsheet Setup and Structure Editing Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Make Google Sheets import self-configuring, make Add sheet visibly and immediately activate a focused worksheet, add left/right table-aware column insertion everywhere users expect it, and close the confirmed DataTable overlay and first-click regressions.

**Architecture:** A strict `worksheetStructure` reducer becomes the only command path for row and column edits and composes the existing sheet-plane shifter with structured-table metadata projection. Standalone-only sizing and Google client-ID persistence stay at the `App` composition root, while reusable `Spreadsheet` instances remain host-sized and storage-neutral for Google configuration. Browser-specific Google state lives in a controller hook and accessible dialog; DataTable menus use the native top layer without weakening component containment.

**Tech Stack:** TypeScript 6, React 19, Vitest 4, Testing Library, fast-check 4, Playwright 1.61, HyperFormula 3.3, Vite 8, native `<dialog>` and Popover APIs, pnpm 11, Node 22.13 or newer.

## Global Constraints

- Implement on `feat/spreadsheet-ux-followup`, based directly on squash-merged `origin/main` commit `98b48a1`.
- Use `/home/ranar2/.local/node22/bin` first in `PATH`; the system Node 20 cannot run pnpm 11.
- Prefix repository commands with `PATH=/home/ranar2/.local/node22/bin:$PATH corepack pnpm` unless the shell has already exported that path.
- Do not add Laravel, a backend, Google write-back, refresh, synchronization, or conflict handling.
- Google Sheets remains a one-time workbook import and the UI copy must say **Import**, never **Link** or **Connect**.
- Never request or persist a Google client secret, access token, sheet URL, spreadsheet payload, or Google error response body.
- Only standalone `App` may opt into Google client-ID storage and viewport ownership. Public `Spreadsheet` remains host-sized with its documented `420px` minimum.
- The current `http://192.168.6.232:4173` origin must receive an actionable built-in OAuth block; do not redirect or recommend the user to localhost.
- Keep row/column commands absolute and serializable. Left/right remain UI intent, represented by `index`, `count`, and boundary-only `expandTableIds`.
- Structural bounds are strict; invalid custom-app commands reject without changing workbook, selection, history, focus, revision, or consuming IDs.
- Preserve table IDs workbook-wide, column IDs within their table, row IDs, headers, formulas, totals, filters, sorts, keys, non-overlap, and exact range/column contiguity.
- A generic row edit may shift a table only when it stays outside the table. Table-body row edits continue through existing `table.insertRows` and `table.deleteRows` commands.
- Use the existing white, slate, and muted-green tokens. New UI must remain scoped below the spreadsheet or DataTable root.
- Use native browser top-layer APIs; add no overlay, OAuth, state-machine, or styling dependency.
- Keep one live review server. Playwright must use `E2E_BASE_URL=http://192.168.6.232:4173` without starting port 5173.
- Follow red-green-refactor for every behavior-changing task and make one focused commit per task.
- Do not open the pull request until typecheck, unit/property tests, app build, library build, and external-server browser tests pass.

## File Structure

New files and single responsibilities:

- `src/core/workbook/worksheetStructure.ts` — strict structural command reducer, table geometry projection, ID generation, and actionable rejections.
- `src/core/workbook/worksheetStructure.test.ts` — deterministic examples for bounds, protection, table insertion/deletion, row guards, formulas, and adjacent tables.
- `src/core/workbook/worksheetStructure.property.test.ts` — bounded generated structural operations checked against workbook/table invariants.
- `src/components/ColumnHeaderContextMenu.tsx` — spreadsheet column-header structural actions only.
- `src/components/ColumnHeaderContextMenu.test.tsx` — menu positioning, focus, actions, and dismissal.
- `src/components/SheetTabs.test.tsx` — active-tab scroll behavior.
- `src/standalone.css` — global document sizing imported only by the standalone workbook route.
- `src/demo/SpreadsheetEmbeddingDemo.tsx` — fixed-height host fixture for browser verification.
- `src/lib/googleErrors.ts` — safe typed Google connector errors.
- `src/lib/googleConfiguration.ts` — client-ID validation, origin assessment, and pure auth-source resolution.
- `src/lib/googleConfiguration.test.ts` — configuration precedence and origin matrix.
- `src/lib/googleAuth.test.ts` — GIS preload, synchronous token request, popup errors, and token cache.
- `src/react/browserGoogleClientIdStorage.ts` — versioned standalone client-ID storage adapter.
- `src/react/browserGoogleClientIdStorage.test.ts` — storage lifecycle and failure behavior.
- `src/react/useGoogleSheetsImport.ts` — per-instance configuration/provider cache and single-flight import controller.
- `src/react/useGoogleSheetsImport.test.tsx` — provider precedence, cancellation, retry, and stale-result isolation.
- `src/components/GoogleSheetsImportDialog.tsx` — accessible setup/import modal presentation.
- `src/components/GoogleSheetsImportDialog.test.tsx` — focus, validation, save/forget, blocked origin, progress, and errors.
- `tests/datatable.spec.ts` — DataTable top-layer and narrow tree-toggle browser regressions.

Existing files modified:

- `src/core/workbook/commands.ts`, `src/core/workbook/commands.test.ts` — route structure commands and add `expandTableIds`.
- `src/core/workbook/WorkbookSession.test.ts` — rejection/history/ID-call contracts.
- `src/core/workbook/structuredTables.ts` — export existing header normalization/name helpers.
- `src/core/workbook/services.ts` — nested Google service and client-ID storage contracts.
- `src/lib/workbook.ts`, `src/lib/workbook.test.ts` — expose the table-free sheet-plane primitive and guard legacy source-only helpers.
- `src/lib/googleAuth.ts`, `src/lib/googleSheets.ts`, `src/lib/googleSheets.test.ts` — prepared browser provider and typed failures.
- `src/App.tsx`, `src/App.test.tsx`, `src/App.css` — directional insert, context state, sheet focus, Google controller/dialog, and scoped UI.
- `src/components/Toolbar.tsx`, `src/components/CellContextMenu.tsx` — insert split button, right insertion, and Import naming.
- `src/components/Grid.tsx`, `src/components/Grid.test.tsx` — focused-cell API and column-header context behavior.
- `src/components/SheetTabs.tsx` — active-tab visibility.
- `src/react/viewport/types.ts`, `src/react/viewport/GridViewport.tsx`, `src/react/viewport/GridViewport.test.tsx` — header context and focused-cell APIs.
- `src/react/DataTable.tsx`, `src/react/DataTableColumnMenu.tsx`, `src/react/DataTableCell.tsx`, `src/react/DataTable.test.tsx` — top-layer menu and one-click tree controls.
- `src/styles/data-table.css`, `src/styles/styleIsolation.test.tsx` — scoped fixed popover and hit target.
- `src/Spreadsheet.controlled.test.tsx`, `src/react/Spreadsheet.types.test.tsx` — service compatibility and public types.
- `src/react/index.ts`, `src/entry/core.ts`, `src/entry/react.ts`, `src/entry/google.ts`, `src/entry/entrypoints.test.ts`, `src/index.ts` — public type/connector exports without adding DOM to core.
- `src/react/browserWorkbookStorage.ts` — no behavior change; reference pattern only.
- `src/test/setup.ts` — deterministic jsdom shims for native dialog/popover methods.
- `src/main.tsx`, `playwright.config.ts`, `tests/spreadsheet.spec.ts` — standalone/embed routes and external-server browser coverage.
- `README.md`, `docs/google-sheets-connector.md`, `docs/getting-started.md`, `docs/embedding.md`, `docs/features.md` — accurate Import/setup/embedding documentation.

## Shared Test Fixture Contracts

Define the structural fixtures once at the bottom of `worksheetStructure.test.ts` and import or copy only the minimal builder needed by the property/session tests:

~~~ts
const range = (startRow: number, startColumn: number, endRow: number, endColumn: number): CellRange => ({
  start: { row: startRow, column: startColumn },
  end: { row: endRow, column: endColumn }
});

function deterministicServices() {
  let sequence = 0;
  return {
    createId: vi.fn((kind: IdKind) => `${kind}-${++sequence}`)
  };
}

function structuredWorkbook(): WorkbookModel {
  const workbook = createBlankWorkbook();
  const sheet = workbook.sheets[0];
  return {
    ...workbook,
    sheets: [{
      ...sheet,
      cells: {
        A1: "Region", B1: "Sales", A2: "West", B2: 10, A3: "East", B3: 8,
        D1: "Department", E1: "Cost", D2: "Operations", E2: 5, D3: "Technology", E3: 7
      }
    }],
    tables: [
      {
        id: "table-sales",
        name: "SalesTable",
        sheetId: sheet.id,
        range: range(0, 0, 2, 1),
        headerRow: true,
        totalsRow: false,
        columns: [
          { id: "sales-region", name: "Region", sheetColumn: 0 },
          { id: "sales-value", name: "Sales", sheetColumn: 1 }
        ],
        rowIds: ["sales-west", "sales-east"]
      },
      {
        id: "table-costs",
        name: "CostsTable",
        sheetId: sheet.id,
        range: range(0, 3, 2, 4),
        headerRow: true,
        totalsRow: false,
        columns: [
          { id: "cost-department", name: "Department", sheetColumn: 3 },
          { id: "cost-value", name: "Cost", sheetColumn: 4 }
        ],
        rowIds: ["cost-operations", "cost-technology"]
      }
    ]
  };
}

function threeColumnCalculatedWorkbook(formula = "=B2*2"): WorkbookModel {
  const base = structuredWorkbook();
  const sales = base.tables[0];
  return {
    ...base,
    sheets: base.sheets.map((sheet) => sheet.id === sales.sheetId ? {
      ...sheet,
      cells: { ...sheet.cells, C1: "Total", C2: "=B2*2", C3: "=B3*2" }
    } : sheet),
    tables: [
      {
        ...sales,
        range: range(0, 0, 2, 2),
        columns: [
          { id: "column-a", name: "Region", sheetColumn: 0 },
          { id: "column-b", name: "Sales", sheetColumn: 1 },
          { id: "column-c", name: "Total", sheetColumn: 2, calculatedFormula: formula }
        ],
        keyColumnId: "column-b",
        sort: [{ columnId: "column-b", direction: "asc" }],
        filter: {
          kind: "comparison",
          columnId: "column-b",
          operator: "eq",
          value: { type: "number", value: 10 }
        }
      },
      base.tables[1]
    ]
  };
}

const calculatedFormulaWorkbook = threeColumnCalculatedWorkbook;
~~~

`adjacentSingleColumnTables()` uses two one-column tables at `A1:A3` and `B1:B3`; `tableStartingAtRowFive()` moves `SalesTable` and its cell plane to `A5:B7`; both must satisfy `migrateWorkbookModel` before the test begins. `generatedTwoTableWorkbook(gap)` starts the second table exactly `gap` columns after the first table's right edge. `sequentialServices()` is the same monotonic generator without Vitest spies. These builders must throw during test setup if their initial workbook does not migrate.

Define Google test primitives in `googleConfiguration.test.ts`:

~~~ts
const CLIENT_ID = "123-abc.apps.googleusercontent.com";
const provider = (token: string): TokenProvider => ({
  getAccessToken: vi.fn().mockResolvedValue(token)
});
const nestedFactory = vi.fn(() => provider("nested"));
const legacyFactory = vi.fn(() => provider("legacy"));
const builtInFactory = vi.fn(() => provider("built-in"));
~~~

Define a standards-shaped in-memory store in `browserGoogleClientIdStorage.test.ts`:

~~~ts
class MemoryStorage implements Storage {
  private readonly values = new Map<string, string>();
  get length() { return this.values.size; }
  clear() { this.values.clear(); }
  getItem(key: string) { return this.values.get(key) ?? null; }
  key(index: number) { return [...this.values.keys()][index] ?? null; }
  removeItem(key: string) { this.values.delete(key); }
  setItem(key: string, value: string) { this.values.set(key, value); }
}
~~~

Define `setupController(overrides)` in the dialog test as a complete `GoogleSheetsImportController` with `phase: "setup"`, `origin: "https://sheets.example.com"`, eligible origin assessment, missing client-ID source, empty drafts, and a `vi.fn()` for every action. Merge overrides last so each test changes one state without constructing an inconsistent controller.

---

### Task 0: Establish the Node 22 and Baseline Gate

**Files:**
- Verify only: `package.json`
- Verify only: `pnpm-lock.yaml`
- Verify only: current branch and committed design specification

**Interfaces:**
- Consumes: Node `v22.23.1` at `/home/ranar2/.local/node22/bin/node`, pnpm `11.7.0`, and the lockfile.
- Produces: installed dependencies and a recorded green baseline. This task changes no tracked file and creates no commit.

- [ ] **Step 1: Activate the supported toolchain**

Run:

~~~bash
export PATH=/home/ranar2/.local/node22/bin:$PATH
node --version
corepack pnpm --version
~~~

Expected: `v22.23.1` and `11.7.0`. Stop if Node resolves to `/usr/bin/node` or reports version 20.

- [ ] **Step 2: Install exactly the locked dependencies in the new worktree**

Run:

~~~bash
corepack pnpm install --frozen-lockfile
~~~

Expected: exit 0 with no `pnpm-lock.yaml` change.

- [ ] **Step 3: Run the focused baseline**

Run:

~~~bash
corepack pnpm exec vitest run src/core/workbook/commands.test.ts src/core/workbook/WorkbookSession.test.ts src/lib/workbook.test.ts src/App.test.tsx src/Spreadsheet.controlled.test.tsx src/react/DataTable.test.tsx src/react/viewport/GridViewport.test.tsx
~~~

Expected: all existing tests pass before feature changes.

- [ ] **Step 4: Confirm the branch is clean except for committed planning documents**

Run:

~~~bash
git status --short --branch
git log -2 --oneline
~~~

Expected: `feat/spreadsheet-ux-followup` is ahead of `origin/main` only by documentation commits and has no untracked source changes.

### Task 1: Add a Strict Worksheet-Structure Reducer Shell

**Files:**
- Create: `src/core/workbook/worksheetStructure.ts`
- Create: `src/core/workbook/worksheetStructure.test.ts`
- Modify: `src/core/workbook/commands.ts`
- Modify: `src/core/workbook/commands.test.ts`
- Modify: `src/lib/workbook.ts`
- Test: `src/lib/workbook.test.ts`

**Interfaces:**
- Consumes: `WorkbookModel`, `TableIssue`, `IdGenerator`, and the current sheet-plane shift behavior.
- Produces: `WorksheetStructureCommand`, `WorksheetStructureReduction`, `reduceWorksheetStructureCommand`, and internal `shiftSheetStructurePlanes`. Later tasks extend this reducer without changing its public signature.

- [ ] **Step 1: Write failing reducer tests for bounds, protection, and table-free shifts**

Create tests with these exact expectations:

~~~ts
import { describe, expect, it, vi } from "vitest";
import { createBlankWorkbook, getCellContent, setCellContent, setSheetProtection } from "../../lib/workbook";
import { reduceWorksheetStructureCommand } from "./worksheetStructure";

const services = { createId: vi.fn((kind: string) => `${kind}-generated`) };

describe("worksheet structure reducer", () => {
  it.each([
    { type: "columns.insert", index: 27, count: 1 },
    { type: "columns.delete", index: 25, count: 2 },
    { type: "rows.insert", index: -1, count: 1 },
    { type: "rows.delete", index: 0, count: 0 }
  ] as const)("rejects strict structural bounds: $type", (command) => {
    const workbook = createBlankWorkbook();
    const result = reduceWorksheetStructureCommand(workbook, {
      ...command,
      sheetId: workbook.activeSheetId
    }, services);
    expect(result).toMatchObject({
      status: "rejected",
      reason: "validation",
      workbook,
      issues: [{ code: expect.stringMatching(/^SHEET_STRUCTURE_/) }]
    });
    expect(services.createId).not.toHaveBeenCalled();
  });

  it("rejects protected sheets without mutation", () => {
    let workbook = createBlankWorkbook();
    workbook = setSheetProtection(workbook, workbook.activeSheetId, true);
    const result = reduceWorksheetStructureCommand(workbook, {
      type: "columns.insert",
      sheetId: workbook.activeSheetId,
      index: 0,
      count: 1
    }, services);
    expect(result).toEqual({
      status: "rejected",
      reason: "permission",
      workbook,
      issues: [{ code: "TABLE_PROTECTED", message: "Protected sheets cannot change worksheet structure" }]
    });
  });

  it("commits a valid table-free plane shift", () => {
    let workbook = createBlankWorkbook();
    const sheetId = workbook.activeSheetId;
    workbook = setCellContent(workbook, sheetId, "A1", "Name");
    const result = reduceWorksheetStructureCommand(workbook, {
      type: "columns.insert", sheetId, index: 0, count: 1
    }, services);
    expect(result.status).toBe("committed");
    if (result.status !== "committed") return;
    expect(getCellContent(result.workbook, sheetId, "B1")).toBe("Name");
  });
});
~~~

- [ ] **Step 2: Run the reducer tests and verify red**

Run:

~~~bash
corepack pnpm exec vitest run src/core/workbook/worksheetStructure.test.ts
~~~

Expected: FAIL because `worksheetStructure.ts` does not exist.

- [ ] **Step 3: Define the strict command and reduction contracts**

Create these exact exported types in `worksheetStructure.ts`:

~~~ts
export type WorksheetStructureCommand =
  | { type: "rows.insert" | "rows.delete"; sheetId: string; index: number; count: number }
  | {
      type: "columns.insert";
      sheetId: string;
      index: number;
      count: number;
      expandTableIds?: readonly string[];
    }
  | { type: "columns.delete"; sheetId: string; index: number; count: number };

export type WorksheetStructureServices = Readonly<{ createId: IdGenerator }>;

export type WorksheetStructureReduction =
  | { status: "committed"; workbook: WorkbookModel }
  | {
      status: "rejected";
      reason: "validation" | "permission";
      workbook: WorkbookModel;
      issues: readonly TableIssue[];
    };

export function reduceWorksheetStructureCommand(
  workbook: WorkbookModel,
  command: WorksheetStructureCommand,
  services: WorksheetStructureServices
): WorksheetStructureReduction;
~~~

Use these exact issue codes: `SHEET_STRUCTURE_INDEX_INVALID`, `SHEET_STRUCTURE_COUNT_INVALID`, `SHEET_STRUCTURE_OUT_OF_BOUNDS`, `TABLE_PROTECTED`, `TABLE_PARTIAL_STRUCTURAL_EDIT`, `TABLE_EXPANSION_CONTEXT_INVALID`, `TABLE_RANGE_OVERLAP`, and `TABLE_COLUMN_ID_CONFLICT`.

- [ ] **Step 4: Implement strict validation and the initial safe reducer**

Implement the table-free shell with this control flow:

~~~ts
export function reduceWorksheetStructureCommand(
  workbook: WorkbookModel,
  command: WorksheetStructureCommand,
  services: WorksheetStructureServices
): WorksheetStructureReduction {
  void services;
  const sheet = workbook.sheets.find((candidate) => candidate.id === command.sheetId);
  if (!sheet) return rejected(workbook, "validation", "SHEET_STRUCTURE_OUT_OF_BOUNDS", "Worksheet does not exist");
  if (!Number.isInteger(command.index) || command.index < 0) {
    return rejected(workbook, "validation", "SHEET_STRUCTURE_INDEX_INVALID", "Structure index must be a non-negative integer");
  }
  if (!Number.isInteger(command.count) || command.count <= 0) {
    return rejected(workbook, "validation", "SHEET_STRUCTURE_COUNT_INVALID", "Structure count must be a positive integer");
  }
  const dimension = command.type.startsWith("rows.") ? sheet.rowCount : sheet.columnCount;
  const withinBounds = command.type.endsWith(".insert")
    ? command.index <= dimension
    : command.index < dimension && command.index + command.count <= dimension;
  if (!withinBounds) {
    return rejected(workbook, "validation", "SHEET_STRUCTURE_OUT_OF_BOUNDS", "Structure edit is outside the worksheet");
  }
  if (sheet.protection.isProtected) {
    return rejected(workbook, "permission", "TABLE_PROTECTED", "Protected sheets cannot change worksheet structure");
  }
  if (workbook.tables.some((table) => table.sheetId === command.sheetId)) {
    return rejected(workbook, "validation", "TABLE_PARTIAL_STRUCTURAL_EDIT", "Structured tables require table-aware structure editing");
  }
  return {
    status: "committed",
    workbook: shiftSheetStructurePlanes(workbook, command.sheetId, operationFor(command))
  };
}
~~~

Define `rejected` to return the exact reduction shape, and `operationFor` to map `rows.*` to axis `row`, `columns.*` to axis `column`, and the suffix to mode `insert | delete`. Rename the current private `shiftSheetStructure` in `src/lib/workbook.ts` to exported source-internal `shiftSheetStructurePlanes`; remove its clamping and accept only the reducer's normalized operation.

Keep `insertRows`, `deleteRows`, `insertColumns`, and `deleteColumns` four-argument signatures for table-free source tests. Before delegating to the plane helper, throw `Error("Use WorkbookSession.dispatch for structured-table worksheet edits")` when the sheet owns a structured table; these helpers are not exported from package entrypoints and cannot express IDs or boundary intent.

- [ ] **Step 5: Route core commands through the reducer**

Replace the four raw structure cases in `applyWorkbookMutation` with:

~~~ts
if (isWorksheetStructureCommand(command)) {
  const reduction = reduceWorksheetStructureCommand(workbook, command, {
    createId: context.createId ?? createRandomId
  });
  return reduction.status === "committed"
    ? applied(reduction.workbook)
    : {
        status: "rejected",
        reason: reduction.reason,
        issues: reduction.issues
      };
}
~~~

Export `isWorksheetStructureCommand` from the new module and include `WorksheetStructureCommand` once in the `WorkbookCommand` union. Remove the old combined row/column union and stop calling `checkedStructure` for these commands.

- [ ] **Step 6: Run focused tests and make them green**

Run:

~~~bash
corepack pnpm exec vitest run src/core/workbook/worksheetStructure.test.ts src/core/workbook/commands.test.ts src/lib/workbook.test.ts
~~~

Expected: PASS. Existing table-free shifting remains unchanged; strict invalid commands return actionable validation results.

- [ ] **Step 7: Commit the strict reducer shell**

~~~bash
git add src/core/workbook/worksheetStructure.ts src/core/workbook/worksheetStructure.test.ts src/core/workbook/commands.ts src/core/workbook/commands.test.ts src/lib/workbook.ts src/lib/workbook.test.ts
git commit -m "refactor(workbook): route worksheet structure through strict reducer"
~~~

### Task 2: Implement Table-Aware Column Insertion

**Files:**
- Modify: `src/core/workbook/worksheetStructure.ts`
- Modify: `src/core/workbook/worksheetStructure.test.ts`
- Modify: `src/core/workbook/structuredTables.ts`
- Test: `src/lib/formulaReferences.test.ts`

**Interfaces:**
- Consumes: Task 1 reducer, `IdGenerator`, `rewriteFormulaForStructure`, current table header normalization, and `shiftSheetStructurePlanes`.
- Produces: automatic internal expansion and explicit boundary expansion through `expandTableIds`, deterministic table-column IDs, unique headers, and projected table ranges.

- [ ] **Step 1: Add failing insertion examples**

Add tests covering these exact outcomes on a fixture with `SalesTable` at `A1:B3` and `CostsTable` at `D1:E3`:

~~~ts
it("shifts later tables and expands only the intended boundary table", () => {
  const workbook = structuredWorkbook();
  const result = reduceWorksheetStructureCommand(workbook, {
    type: "columns.insert",
    sheetId: workbook.activeSheetId,
    index: 2,
    count: 1,
    expandTableIds: ["table-sales"]
  }, deterministicServices());
  expect(result.status).toBe("committed");
  if (result.status !== "committed") return;
  expect(result.workbook.tables[0].range).toEqual(range(0, 0, 2, 2));
  expect(result.workbook.tables[0].columns.map((column) => column.name)).toEqual(["Region", "Sales", "Column3"]);
  expect(result.workbook.tables[1].range).toEqual(range(0, 4, 2, 5));
  expect(result.workbook.sheets[0].cells.C1).toBe("Column3");
});

it("expands an internal insertion without boundary context", () => {
  const workbook = structuredWorkbook();
  const result = reduceWorksheetStructureCommand(workbook, {
    type: "columns.insert", sheetId: workbook.activeSheetId, index: 1, count: 2
  }, deterministicServices());
  expect(result.status).toBe("committed");
  if (result.status !== "committed") return;
  expect(result.workbook.tables[0].columns.map((column) => [column.id, column.name, column.sheetColumn])).toEqual([
    ["sales-region", "Region", 0],
    ["table-column-1", "Column2", 1],
    ["table-column-2", "Column3", 2],
    ["sales-value", "Sales", 3]
  ]);
});

it("rejects duplicate, stale, cross-sheet, or non-boundary expansion ids before generating ids", () => {
  const workbook = structuredWorkbook();
  const services = deterministicServices();
  const result = reduceWorksheetStructureCommand(workbook, {
    type: "columns.insert",
    sheetId: workbook.activeSheetId,
    index: 2,
    count: 1,
    expandTableIds: ["table-sales", "table-sales"]
  }, services);
  expect(result).toMatchObject({ status: "rejected", issues: [{ code: "TABLE_EXPANSION_CONTEXT_INVALID" }] });
  expect(services.createId).not.toHaveBeenCalled();
});
~~~

The fixture must include adjacent-table coverage where inserting at one table's right edge shifts the neighbor and never expands both accidentally.
Also assert that existing sort, filter, key, calculated-formula, totals, data-type, and style fields stay on their original column IDs; inserted columns have none of those fields. Add one `headerRow: false` case that creates metadata names without writing a physical header cell.

- [ ] **Step 2: Run insertion tests and verify red**

Run:

~~~bash
corepack pnpm exec vitest run src/core/workbook/worksheetStructure.test.ts -t "insertion|expands|boundary"
~~~

Expected: FAIL because Task 1 still rejects every workbook table.

- [ ] **Step 3: Export the existing header helpers**

In `structuredTables.ts`, export the existing normalization and generated-name logic under these stable names:

~~~ts
export function normalizeStructuredTableHeader(value: string): string;

export function nextStructuredTableColumnName(
  ordinal: number,
  normalizedNames: ReadonlySet<string>
): string;
~~~

Do not create a second header-equivalence rule. `nextStructuredTableColumnName` must advance until its normalized `ColumnN` candidate is absent.

- [ ] **Step 4: Implement preflight and deterministic insertion projection**

Add an insertion planner whose output is applied only after all validation succeeds:

~~~ts
type HeaderWrite = Readonly<{ sheetId: string; row: number; column: number; value: string }>;

type StructureProjection = Readonly<{
  tables: readonly StructuredTable[];
  headerWrites: readonly HeaderWrite[];
}>;

function projectColumnInsertion(
  workbook: WorkbookModel,
  command: Extract<WorksheetStructureCommand, { type: "columns.insert" }>,
  services: WorksheetStructureServices
): StructureProjection | readonly TableIssue[];
~~~

Implement this decision table exactly:

~~~text
index < start                         -> shift start/end and every sheetColumn by count
start < index <= end                  -> expand automatically at index
index == start                        -> expand only when table id is listed; otherwise shift
index == end + 1                      -> expand only when table id is listed; otherwise unchanged
index > end + 1                       -> unchanged
~~~

Preflight duplicate/stale/cross-sheet/non-boundary IDs and projected range overlap before requesting any new ID. Then generate IDs in workbook table order and ascending inserted physical-column order. Validate each returned ID against existing and earlier generated column IDs; a collision is a `TABLE_COLUMN_ID_CONFLICT` service-contract failure, not a command-preflight failure. New columns contain only `{ id, name, sheetColumn }`; write generated names to the projected header row only when `headerRow` is true.

Apply the result in this order inside the reducer:

~~~ts
const projection = projectColumnInsertion(workbook, command, services);
if (Array.isArray(projection)) {
  return { status: "rejected", reason: "validation", workbook, issues: projection };
}
const shifted = shiftSheetStructurePlanes(workbook, command.sheetId, operationFor(command));
return {
  status: "committed",
  workbook: applyHeaderWrites({ ...shifted, tables: [...projection.tables] }, projection.headerWrites)
};
~~~

`applyHeaderWrites` clones only the owning sheet and its `cells` map and writes the exact projected addresses.

- [ ] **Step 5: Run insertion and formula regression tests**

Run:

~~~bash
corepack pnpm exec vitest run src/core/workbook/worksheetStructure.test.ts src/lib/formulaReferences.test.ts src/core/workbook/migrateWorkbook.test.ts
~~~

Expected: PASS for before/internal/left-boundary/right-boundary/adjacent tables, multi-column counts, unique names, deterministic IDs, and ordinary A1 formula rewriting.

- [ ] **Step 6: Commit table-aware insertion**

~~~bash
git add src/core/workbook/worksheetStructure.ts src/core/workbook/worksheetStructure.test.ts src/core/workbook/structuredTables.ts src/lib/formulaReferences.test.ts
git commit -m "feat(workbook): insert columns through structured tables"
~~~

### Task 3: Complete Deletion, Row Guards, and Canonical Formula Semantics

**Files:**
- Modify: `src/core/workbook/worksheetStructure.ts`
- Modify: `src/core/workbook/worksheetStructure.test.ts`
- Modify: `src/core/workbook/commands.test.ts`
- Test: `src/core/workbook/structuredTableRows.test.ts`

**Interfaces:**
- Consumes: Tasks 1–2 reducer/projection and `rewriteFormulaForStructure`.
- Produces: interval-correct table column deletion, safe generic row shifting/rejection, filter/sort/key cleanup, and canonical calculated-formula rewriting.

- [ ] **Step 1: Write failing column-deletion tests**

Add exact assertions for half-open intervals:

~~~ts
it("deletes an internal table column and preserves surviving logical identities", () => {
  const workbook = threeColumnCalculatedWorkbook();
  const result = reduceWorksheetStructureCommand(workbook, {
    type: "columns.delete", sheetId: workbook.activeSheetId, index: 1, count: 1
  }, deterministicServices());
  expect(result.status).toBe("committed");
  if (result.status !== "committed") return;
  const table = result.workbook.tables[0];
  expect(table.range).toEqual(range(0, 0, 3, 1));
  expect(table.columns.map((column) => [column.id, column.sheetColumn])).toEqual([
    ["column-a", 0], ["column-c", 1]
  ]);
  expect(table.sort).toEqual([]);
  expect(table.filter).toBeUndefined();
  expect(table.keyColumnId).toBeUndefined();
});

it("rejects atomically when any affected table would lose every column", () => {
  const workbook = adjacentSingleColumnTables();
  const result = reduceWorksheetStructureCommand(workbook, {
    type: "columns.delete", sheetId: workbook.activeSheetId, index: 0, count: 2
  }, deterministicServices());
  expect(result).toMatchObject({ status: "rejected", issues: [{ code: "TABLE_PARTIAL_STRUCTURAL_EDIT" }] });
  expect(result.workbook).toBe(workbook);
});
~~~

Cover deletion wholly before, wholly after, entering from the left, leaving to the right, spanning multiple tables, and projected overlap. Assert that any filter leaf on a removed ID clears the entire Boolean filter, while sort entries are removed individually.

- [ ] **Step 2: Write failing generic-row and canonical-formula tests**

Add these contracts:

~~~ts
it("shifts tables for generic row edits before them and rejects edits inside them", () => {
  const workbook = tableStartingAtRowFive();
  const before = reduceWorksheetStructureCommand(workbook, {
    type: "rows.insert", sheetId: workbook.activeSheetId, index: 2, count: 2
  }, deterministicServices());
  expect(before.status).toBe("committed");
  if (before.status !== "committed") return;
  expect(before.workbook.tables[0].range.start.row).toBe(6);
  expect(before.workbook.tables[0].rowIds).toEqual(workbook.tables[0].rowIds);

  const inside = reduceWorksheetStructureCommand(workbook, {
    type: "rows.delete", sheetId: workbook.activeSheetId, index: 5, count: 1
  }, deterministicServices());
  expect(inside).toMatchObject({ status: "rejected", issues: [{ code: "TABLE_PARTIAL_STRUCTURAL_EDIT" }] });
  expect(inside.workbook).toBe(workbook);
});

it("rewrites canonical calculated formulas before later row regeneration", () => {
  const workbook = calculatedFormulaWorkbook("=B2*$C$2");
  const inserted = reduceWorksheetStructureCommand(workbook, {
    type: "columns.insert", sheetId: workbook.activeSheetId, index: 1, count: 1
  }, deterministicServices());
  expect(inserted.status).toBe("committed");
  if (inserted.status !== "committed") return;
  expect(inserted.workbook.tables[0].columns.find((column) => column.id === "calculated")?.calculatedFormula)
    .toBe("=C2*$D$2");
});
~~~

- [ ] **Step 3: Run the new cases and verify red**

Run:

~~~bash
corepack pnpm exec vitest run src/core/workbook/worksheetStructure.test.ts -t "deletes|generic row|canonical"
~~~

Expected: FAIL because Task 2 handles insertion only.

- [ ] **Step 4: Implement half-open deletion projection**

Use these exact interval rules, with `deleteStart = index`, `deleteEnd = index + count`, `tableStart = range.start.column`, and `tableEnd = range.end.column + 1`:

~~~text
deleteEnd <= tableStart  -> shift the table left by count
deleteStart >= tableEnd  -> unchanged
otherwise                -> remove overlapping logical columns; shift survivors at/after deleteEnd by count
zero survivors           -> reject the whole command before mutation
~~~

Rebuild an overlapping table's range from the minimum and maximum surviving `sheetColumn`. Remove deleted sort entries, clear the whole filter if any recursive comparison/blank filter leaf references a removed ID, clear a deleted key, and retain every other surviving column field. Implement a local `filterUsesOnlyColumns(filter, survivingIds)` recursion: comparison and blank nodes check their `columnId`; logical nodes require every operand to pass. Preflight all tables and overlap before calling `shiftSheetStructurePlanes`.

- [ ] **Step 5: Implement safe row projection and formula metadata rewriting**

Use these exact row rules:

~~~text
insert index <= table start             -> shift table down by count
insert index >= table end + 1           -> unchanged
table start < insert index <= table end -> reject

delete end <= table start               -> shift table up by count
delete start > table end                -> unchanged
otherwise                               -> reject
~~~

After every accepted row or column operation, map every table in the workbook—not only geometrically affected tables—and rewrite each surviving `calculatedFormula` with:

~~~ts
rewriteFormulaForStructure(column.calculatedFormula, {
  formulaSheetName: owningSheet.name,
  editedSheetName: editedSheet.name,
  axis: operation.axis,
  mode: operation.mode,
  index: operation.index,
  count: operation.count
})
~~~

This covers formulas in other tables that reference the edited sheet. Keep formulas beginning with anything other than `=` unchanged.

- [ ] **Step 6: Prove regeneration does not restore stale formulas**

After the canonical assertion passes, dispatch an existing `table.insertRows` and `table.sort` against the committed workbook and assert all generated body formulas use the shifted references. Run:

~~~bash
corepack pnpm exec vitest run src/core/workbook/worksheetStructure.test.ts src/core/workbook/structuredTableRows.test.ts src/core/workbook/commands.test.ts
~~~

Expected: PASS for deletion cleanup, row guards, canonical formulas, and downstream table regeneration.

- [ ] **Step 7: Commit complete structural semantics**

~~~bash
git add src/core/workbook/worksheetStructure.ts src/core/workbook/worksheetStructure.test.ts src/core/workbook/commands.test.ts src/core/workbook/structuredTableRows.test.ts
git commit -m "fix(workbook): preserve tables through worksheet structure edits"
~~~

### Task 4: Pressure-Test Structure History and Invariants

**Files:**
- Create: `src/core/workbook/worksheetStructure.property.test.ts`
- Modify: `src/core/workbook/WorkbookSession.test.ts`
- Modify: `src/core/workbook/commands.test.ts`
- Verify: `src/core/workbook/migrateWorkbook.ts`

**Interfaces:**
- Consumes: completed `reduceWorksheetStructureCommand`, deterministic `createId`, `WorkbookSession` history, and `migrateWorkbookModel` as the persisted-model oracle.
- Produces: bounded generated coverage and proof that rejected/undo/redo paths neither consume IDs nor leave invalid tables.

- [ ] **Step 1: Add failing session atomicity and ID-call tests**

Add this pattern to `WorkbookSession.test.ts` using the shared structured fixture:

~~~ts
it("undoes and redoes a table-aware insertion without regenerating ids", () => {
  const createId = vi.fn()
    .mockReturnValueOnce("table-column-new-1")
    .mockReturnValueOnce("table-column-new-2");
  const initial = structuredWorkbook();
  const session = createWorkbookSession({ workbook: initial, createId });
  const result = session.dispatch({
    type: "transaction",
    commands: [
      {
        type: "columns.insert",
        sheetId: initial.activeSheetId,
        index: 2,
        count: 2,
        expandTableIds: ["table-sales"]
      },
      {
        type: "selection.set",
        selection: { start: { row: 0, column: 2 }, end: { row: 99, column: 3 } }
      }
    ]
  });
  expect(result.status).toBe("committed");
  const inserted = session.getSnapshot().workbook;
  expect(createId).toHaveBeenCalledTimes(2);

  session.dispatch({ type: "history.undo" });
  expect(session.getSnapshot().workbook).toEqual(initial);
  session.dispatch({ type: "history.redo" });
  expect(session.getSnapshot().workbook).toEqual(inserted);
  expect(createId).toHaveBeenCalledTimes(2);
});

it("keeps selection, revision, and history unchanged when structure rejects", () => {
  const session = createWorkbookSession({ workbook: structuredWorkbook() });
  const before = session.getSnapshot();
  const result = session.dispatch({
    type: "columns.delete",
    sheetId: before.workbook.activeSheetId,
    index: 0,
    count: 2
  });
  expect(result).toMatchObject({ status: "rejected", reason: "validation" });
  expect(session.getSnapshot()).toMatchObject({
    revision: before.revision,
    selection: before.selection,
    canUndo: before.canUndo,
    canRedo: before.canRedo
  });
  expect(session.getSnapshot().workbook).toBe(before.workbook);
});
~~~

- [ ] **Step 2: Add a bounded property test**

Generate table-free gaps and structural operations without generating arbitrary invalid workbooks:

~~~ts
import fc from "fast-check";

it("preserves migrated workbook invariants for bounded multi-table column edits", () => {
  fc.assert(fc.property(
    fc.record({
      gap: fc.integer({ min: 1, max: 5 }),
      count: fc.integer({ min: 1, max: 3 }),
      direction: fc.constantFrom("before", "inside", "after")
    }),
    ({ gap, count, direction }) => {
      const workbook = generatedTwoTableWorkbook(gap);
      const first = workbook.tables[0];
      const index = direction === "before"
        ? first.range.start.column
        : direction === "inside"
          ? first.range.end.column
          : first.range.end.column + 1;
      const expandTableIds = direction === "after" ? [first.id] : undefined;
      const result = reduceWorksheetStructureCommand(workbook, {
        type: "columns.insert",
        sheetId: workbook.activeSheetId,
        index,
        count,
        expandTableIds
      }, sequentialServices());
      expect(result.status).toBe("committed");
      if (result.status !== "committed") return;
      expect(migrateWorkbookModel(result.workbook)).not.toBeNull();
      assertTableStructureInvariants(result.workbook);
    }
  ), { numRuns: 100, seed: 20260710 });
});
~~~

`assertTableStructureInvariants` must explicitly check width/column count, contiguous `sheetColumn`, normalized unique headers, workbook-wide table IDs, per-table column IDs, row-ID/body-height count, surviving filter/sort/key IDs, bounds, and non-overlap. Do not treat `migrateWorkbookModel` alone as sufficient evidence.

- [ ] **Step 3: Run property/session tests and verify any missing case fails**

Run:

~~~bash
corepack pnpm exec vitest run src/core/workbook/worksheetStructure.property.test.ts src/core/workbook/WorkbookSession.test.ts src/core/workbook/commands.test.ts
~~~

Expected before final fixes: at least one new assertion fails if history, ID generation, or strict rejection mapping is incomplete.

- [ ] **Step 4: Fix only invariant/history gaps surfaced by the tests**

Keep generation after complete preflight, keep the whole insert plus selection transaction to one history entry, and map structural reducer failures directly to `validation` or `permission`. Update the existing test that expected invalid structure input to become generic `unsupported`; it must now assert the specific validation issue.

- [ ] **Step 5: Run the entire workbook core**

Run:

~~~bash
corepack pnpm exec vitest run src/core/workbook src/lib/workbook.test.ts src/lib/formulaReferences.test.ts src/lib/formulaEngine.test.ts
~~~

Expected: PASS, including the seeded 100-run property test.

- [ ] **Step 6: Commit the pressure gate**

~~~bash
git add src/core/workbook/worksheetStructure.property.test.ts src/core/workbook/WorkbookSession.test.ts src/core/workbook/commands.test.ts src/core/workbook/worksheetStructure.ts
git commit -m "test(workbook): pressure structural table invariants"
~~~

### Task 5: Expose Left and Right Column Insertion in Every Spreadsheet Surface

**Files:**
- Create: `src/components/ColumnHeaderContextMenu.tsx`
- Create: `src/components/ColumnHeaderContextMenu.test.tsx`
- Modify: `src/components/Toolbar.tsx`
- Modify: `src/components/CellContextMenu.tsx`
- Modify: `src/components/Grid.tsx`
- Modify: `src/components/Grid.test.tsx`
- Modify: `src/react/viewport/types.ts`
- Modify: `src/react/viewport/GridViewport.tsx`
- Modify: `src/react/viewport/GridViewport.test.tsx`
- Modify: `src/App.tsx`
- Modify: `src/App.test.tsx`

**Interfaces:**
- Consumes: `columns.insert` from Tasks 1–4 and the existing accessible `SplitButton`.
- Produces: `handleInsertColumns("left" | "right")`, cell/header/ribbon actions, and `onColumnHeaderContextMenu` through the shared viewport.

- [ ] **Step 1: Add failing App tests for direction and count**

Extend the structural toolbar/context tests with these expectations:

~~~ts
it("inserts the selected width to the right from the ribbon split menu", async () => {
  const user = userEvent.setup();
  render(<Spreadsheet storage={false} />);
  selectRange("A1", "B1");
  await user.click(screen.getByRole("button", { name: "Insert columns options" }));
  await user.click(screen.getByRole("menuitem", { name: "Insert column right" }));
  expect(screen.getByLabelText("Status")).toHaveTextContent("Inserted 2 columns right");
  expect(screen.getByLabelText("Name box")).toHaveValue("C1:D1");
});

it("expands the selected structured table at its right boundary", async () => {
  const user = userEvent.setup();
  render(<Spreadsheet defaultWorkbook={structuredTableWorkbook()} storage={false} />);
  await user.click(screen.getByRole("gridcell", { name: "B2 10" }));
  await user.click(screen.getByRole("button", { name: "Insert columns options" }));
  await user.click(screen.getByRole("menuitem", { name: "Insert column right" }));
  expect(screen.getByRole("gridcell", { name: "C1 Column3" })).toBeVisible();
  expect(screen.getByLabelText("Status")).toHaveTextContent("Inserted 1 column right");
});
~~~

Also cover cell-menu right insertion, append after the final worksheet column, a protected sheet rejection, and undo as one action.

- [ ] **Step 2: Add failing viewport and Grid header-context tests**

Add `onColumnHeaderContextMenu?(column, event)` to the test harness and assert:

~~~ts
fireEvent.contextMenu(screen.getByRole("columnheader", { name: "Column B" }), {
  clientX: 240,
  clientY: 60
});
expect(contexts).toEqual([{ column: 1, x: 240, y: 60 }]);

fireEvent.keyDown(screen.getByRole("columnheader", { name: "Column B" }), {
  key: "F10",
  shiftKey: true
});
expect(contexts).toHaveLength(2);
~~~

In `Grid.test.tsx`, start with a full `A:C` selection, right-click `B`, and assert the selection remains `A:C`; right-click `D` and assert it becomes the full `D` column.

- [ ] **Step 3: Run UI tests and verify red**

Run:

~~~bash
corepack pnpm exec vitest run src/App.test.tsx src/components/Grid.test.tsx src/react/viewport/GridViewport.test.tsx
~~~

Expected: FAIL because the right action and header callback do not exist.

- [ ] **Step 4: Extend the shared viewport and add the header menu**

Add this callback to `GridViewportProps` and wire it directly to each column-header element:

~~~ts
onColumnHeaderContextMenu?(
  column: GridViewportColumn,
  event: ReactMouseEvent<HTMLDivElement>
): void;
~~~

Add this spreadsheet-level callback to `GridProps`:

~~~ts
onColumnHeaderContextMenu?: (event: {
  column: number;
  x: number;
  y: number;
  opener: HTMLElement;
}) => void;
~~~

`Grid` preserves an existing whole-column multi-selection when the clicked index lies inside it; otherwise it selects rows `0..sheet.rowCount - 1` for the clicked column. Both native contextmenu and `Shift+F10`/`ContextMenu` invoke the same callback.

Create `ColumnHeaderContextMenu` with this exact prop surface:

~~~ts
export type ColumnHeaderContextMenuProps = Readonly<{
  label: string;
  x: number;
  y: number;
  opener: HTMLElement;
  onClose(): void;
  onInsertLeft(): void;
  onInsertRight(): void;
  onDelete(): void;
}>;
~~~

Render `role="menu"` with label `Column <label> context menu`, clamp its fixed position to an 8px viewport margin using the same measured-layout pattern as `CellContextMenu`, focus the first item, close on Escape/outside pointer, and restore focus through the supplied `opener`. Store that opener in App's column-menu state from `event.currentTarget` for both mouse and keyboard opening.

- [ ] **Step 5: Centralize the directional App handler**

Replace `handleInsertColumns()` with:

~~~ts
function handleInsertColumns(direction: "left" | "right") {
  if (!ensureSheetStructureEditable()) return;
  const normalized = normalizeRange(selection);
  const count = normalized.end.column - normalized.start.column + 1;
  const index = direction === "left"
    ? normalized.start.column
    : normalized.end.column + 1;
  const expandTableIds = boundaryTablesForColumnInsertion(
    workbook,
    activeSheet.id,
    normalized,
    index
  );
  dispatchCommand({
    type: "transaction",
    commands: [
      {
        type: "columns.insert",
        sheetId: activeSheet.id,
        index,
        count,
        ...(expandTableIds.length === 0 ? {} : { expandTableIds })
      },
      {
        type: "selection.set",
        selection: {
          start: { row: normalized.start.row, column: index },
          end: { row: normalized.end.row, column: index + count - 1 }
        }
      }
    ]
  }, `Inserted ${pluralize(count, "column")} ${direction}`);
}
~~~

`boundaryTablesForColumnInsertion` includes only tables on the active sheet whose row ranges intersect the selection, whose column ranges intersect the selection, and whose `range.start.column` or `range.end.column + 1` equals `index`.

- [ ] **Step 6: Wire ribbon and cell/header menus**

Extend `SplitButton` with optional `primaryAriaLabel` so the visible label can remain **Insert columns** while its primary action is announced as **Insert column left**. Use:

~~~tsx
<SplitButton
  label="Insert columns"
  primaryAriaLabel="Insert column left"
  icon={<Columns3 />}
  onPrimary={props.onInsertColumnsLeft}
  items={[
    { label: "Insert column left", onSelect: props.onInsertColumnsLeft },
    { label: "Insert column right", onSelect: props.onInsertColumnsRight }
  ]}
/>
~~~

Split the cell-menu insert prop into `onInsertColumnLeft` and `onInsertColumnRight`, and render the new header menu from App state alongside `CellContextMenu`.

- [ ] **Step 7: Run all focused spreadsheet interaction tests**

Run:

~~~bash
corepack pnpm exec vitest run src/App.test.tsx src/components/ColumnHeaderContextMenu.test.tsx src/components/Grid.test.tsx src/react/viewport/GridViewport.test.tsx
~~~

Expected: PASS for ribbon, cell, mouse header, keyboard header, single/multi-column, table boundary, protected, and undo flows.

- [ ] **Step 8: Commit directional column UX**

~~~bash
git add src/App.tsx src/App.test.tsx src/components/Toolbar.tsx src/components/CellContextMenu.tsx src/components/ColumnHeaderContextMenu.tsx src/components/ColumnHeaderContextMenu.test.tsx src/components/Grid.tsx src/components/Grid.test.tsx src/react/viewport/types.ts src/react/viewport/GridViewport.tsx src/react/viewport/GridViewport.test.tsx
git commit -m "feat(ui): insert worksheet columns left or right"
~~~

### Task 6: Make Add Sheet Visible, Focused, and Correctly Sized

**Files:**
- Create: `src/components/SheetTabs.test.tsx`
- Create: `src/standalone.css`
- Create: `src/demo/SpreadsheetEmbeddingDemo.tsx`
- Modify: `src/components/SheetTabs.tsx`
- Modify: `src/react/viewport/types.ts`
- Modify: `src/react/viewport/GridViewport.tsx`
- Modify: `src/react/viewport/GridViewport.test.tsx`
- Modify: `src/components/Grid.tsx`
- Modify: `src/components/Grid.test.tsx`
- Modify: `src/App.tsx`
- Modify: `src/App.test.tsx`
- Modify: `src/App.css`
- Modify: `src/main.tsx`
- Test: `src/Spreadsheet.controlled.test.tsx`
- Test: `src/styles/styleIsolation.test.tsx`

**Interfaces:**
- Consumes: active-sheet transaction and shared viewport API.
- Produces: `GridViewportApi.focusCell`, `GridScrollApi.focusCell`, active tab scrolling, standalone-only viewport sizing, and `/embed` verification fixture.

- [ ] **Step 1: Write failing active-tab and focused-cell tests**

Create `SheetTabs.test.tsx` with a `scrollIntoView` spy:

~~~ts
it("scrolls only the active sheet tab with nearest alignment", () => {
  const scrollIntoView = vi.fn();
  Object.defineProperty(HTMLElement.prototype, "scrollIntoView", {
    configurable: true,
    value: scrollIntoView
  });
  const { rerender } = render(
    <SheetTabs sheets={sheets} activeSheetId="sheet-1" onSelect={vi.fn()} onAdd={vi.fn()} />
  );
  scrollIntoView.mockClear();
  rerender(<SheetTabs sheets={sheets} activeSheetId="sheet-2" onSelect={vi.fn()} onAdd={vi.fn()} />);
  expect(scrollIntoView).toHaveBeenCalledTimes(1);
  expect(scrollIntoView).toHaveBeenCalledWith({ block: "nearest", inline: "nearest" });
});
~~~

Update the existing Add sheet test to assert `Sheet2` is selected, `A1` has focus, and status is exactly `Added Sheet2`.

- [ ] **Step 2: Add failing API-bridge tests**

Extend `GridViewportApi` test coverage with:

~~~ts
let api: GridViewportApi | null = null;
render(<StatefulViewport onRegisterApi={(next) => { api = next; }} />);
act(() => api?.focusCell("row/ada", "salary"));
expect(screen.getByRole("gridcell", { name: "Ada Salary" })).toHaveFocus();
~~~

In `Grid.test.tsx`, capture `GridScrollApi`, call `focusCell(0, 1)`, and assert `B1` is visible and focused.

- [ ] **Step 3: Run the tests and verify red**

Run:

~~~bash
corepack pnpm exec vitest run src/components/SheetTabs.test.tsx src/react/viewport/GridViewport.test.tsx src/components/Grid.test.tsx src/App.test.tsx -t "sheet|focus|active tab"
~~~

Expected: FAIL because focus and tab scrolling APIs are absent and status is generic.

- [ ] **Step 4: Implement committed-render focus and active-tab scroll**

Extend both API layers:

~~~ts
export type GridViewportApi = {
  ensureCellVisible(rowId: string, columnId: string): void;
  focusCell(rowId: string, columnId: string): void;
};

export type GridScrollApi = {
  ensureCellVisible(row: number, column: number): void;
  focusCell(row: number, column: number): void;
};
~~~

`GridViewport.focusCell` ensures the cell is rendered, then focuses its `gridCellDomId` with `preventScroll`; if virtualization requires a render, schedule one `requestAnimationFrame` retry. `Grid` maps numeric coordinates through `sheetRowId` and `sheetColumnId`.

In App, set a `pendingGridFocusSheetIdRef` only after `sheet.add` commits. A layout effect keyed by `activeSheet.id` calls `gridApiRef.current?.focusCell(0, 0)` when the pending ID matches, then clears the ref. Read the actual active sheet name from `session.getSnapshot()` and set status `Added <name>`.

In `SheetTabs`, keep a ref map keyed by sheet ID and call `scrollIntoView({ block: "nearest", inline: "nearest" })` from a layout effect when `activeSheetId` changes.

- [ ] **Step 5: Isolate standalone sizing from embedded sizing**

Create `src/standalone.css` with exactly:

~~~css
html,
body,
#root {
  width: 100%;
  height: 100%;
  margin: 0;
}

body {
  overflow: hidden;
}

.js-spreadsheet-standalone {
  min-height: 0 !important;
}
~~~

Render standalone as:

~~~tsx
export default function App() {
  return <Spreadsheet className="js-spreadsheet-standalone" style={{ height: "100dvh", minHeight: 0 }} />;
}
~~~

Import `standalone.css` only in the root workbook branch of `main.tsx`; do not add global selectors to `App.css` or the published stylesheet. Change the narrow `.spreadsheet-surface` height from `100vh` to `100%`.

Add `/embed` routing to `main.tsx` that loads `SpreadsheetEmbeddingDemo` plus `App.css`, but not `standalone.css`. The demo renders `<Spreadsheet storage={false} style={{ height: 480 }} />` inside a `520px` host and exposes `data-testid="embedding-host"`.

- [ ] **Step 6: Run sizing/isolation tests**

Run:

~~~bash
corepack pnpm exec vitest run src/components/SheetTabs.test.tsx src/react/viewport/GridViewport.test.tsx src/components/Grid.test.tsx src/App.test.tsx src/Spreadsheet.controlled.test.tsx src/styles/styleIsolation.test.tsx
~~~

Expected: PASS. Published CSS still contains no global `html`, `body`, or `#root` selectors; embedded `style` remains host-controlled.

- [ ] **Step 7: Commit sheet feedback and standalone sizing**

~~~bash
git add src/components/SheetTabs.tsx src/components/SheetTabs.test.tsx src/react/viewport/types.ts src/react/viewport/GridViewport.tsx src/react/viewport/GridViewport.test.tsx src/components/Grid.tsx src/components/Grid.test.tsx src/App.tsx src/App.test.tsx src/App.css src/standalone.css src/demo/SpreadsheetEmbeddingDemo.tsx src/main.tsx src/Spreadsheet.controlled.test.tsx src/styles/styleIsolation.test.tsx
git commit -m "fix(ui): keep newly added sheets visible and focused"
~~~

### Task 7: Add Typed Google Configuration and Connector Primitives

**Files:**
- Create: `src/lib/googleErrors.ts`
- Create: `src/lib/googleConfiguration.ts`
- Create: `src/lib/googleConfiguration.test.ts`
- Create: `src/lib/googleAuth.test.ts`
- Modify: `src/core/workbook/services.ts`
- Modify: `src/lib/googleAuth.ts`
- Modify: `src/lib/googleSheets.ts`
- Modify: `src/lib/googleSheets.test.ts`
- Modify: `src/entry/google.ts`

**Interfaces:**
- Consumes: existing `TokenProvider`, GIS token flow, and Sheets API importer.
- Produces: nested Google service types, origin/client-ID validation, prepared browser token providers, and typed sanitized connector errors.

- [ ] **Step 1: Write failing configuration tests**

Create `googleConfiguration.test.ts` with this matrix:

~~~ts
describe("Google configuration", () => {
  it.each([
    ["https://sheets.example.com", "eligible"],
    ["https://sheets.example.com:8443", "eligible"],
    ["http://192.168.6.232:4173", "blocked"],
    ["https://192.168.6.232", "blocked"],
    ["https://[2001:db8::1]", "blocked"],
    ["file:///tmp/index.html", "blocked"],
    ["null", "blocked"],
    ["http://127.0.0.1:4173", "eligible"],
    ["http://[::1]:4173", "eligible"]
  ] as const)("assesses %s as %s", (origin, status) => {
    expect(assessGoogleOAuthOrigin(origin).status).toBe(status);
  });

  it("normalizes only public web OAuth client ids", () => {
    expect(validateGoogleClientId(" 123-abc.apps.googleusercontent.com ")).toEqual({
      valid: true,
      value: "123-abc.apps.googleusercontent.com"
    });
    expect(validateGoogleClientId("secret-value")).toMatchObject({ valid: false });
  });

  it("resolves direct provider, managed id, stored id, then factories", () => {
    const direct = provider("direct");
    expect(resolveGoogleAuthSource({ tokenProvider: direct }, null, legacyFactory, builtInFactory))
      .toMatchObject({ kind: "provider", provider: direct });
    expect(resolveGoogleAuthSource({ clientId: CLIENT_ID, tokenProviderFactory: nestedFactory }, { value: "stored.apps.googleusercontent.com", source: "stored" }, legacyFactory, builtInFactory))
      .toMatchObject({ kind: "client", clientId: CLIENT_ID, source: "managed", factory: nestedFactory });
    expect(resolveGoogleAuthSource({}, { value: "stored.apps.googleusercontent.com", source: "stored" }, legacyFactory, builtInFactory))
      .toMatchObject({ kind: "client", clientId: "stored.apps.googleusercontent.com", source: "stored", factory: legacyFactory });
  });
});
~~~

Treat the current raw IP as blocked but do not emit a localhost recommendation in any user-facing message.

- [ ] **Step 2: Write failing typed auth/API tests**

In `googleAuth.test.ts`, stub `window.google.accounts.oauth2.initTokenClient` and assert:

~~~ts
const provider = createBrowserTokenProvider(CLIENT_ID);
await provider.prepare();
const tokenPromise = provider.getAccessToken([SHEETS_READONLY_SCOPE]);
expect(requestAccessToken).toHaveBeenCalledTimes(1);
callback({ access_token: "token" });
await expect(tokenPromise).resolves.toBe("token");
await expect(provider.getAccessToken([SHEETS_READONLY_SCOPE])).resolves.toBe("token");
expect(requestAccessToken).toHaveBeenCalledTimes(1);
~~~

Cover `popup_failed_to_open` → `popup_blocked`, `popup_closed` → `popup_closed`, access denial, script load failure, initialization failure, and invalid client. In `googleSheets.test.ts`, assert exact codes for invalid URL, 401/403 access denied, API-disabled 403 body, 404, 429, network rejection, and unknown status; assert messages contain neither bearer token nor response body.

- [ ] **Step 3: Run connector tests and verify red**

Run:

~~~bash
corepack pnpm exec vitest run src/lib/googleConfiguration.test.ts src/lib/googleAuth.test.ts src/lib/googleSheets.test.ts
~~~

Expected: FAIL because the configuration/error modules and prepared provider do not exist.

- [ ] **Step 4: Add the public service contracts**

In `services.ts`, define:

~~~ts
export interface TokenProvider {
  prepare?(): void | Promise<void>;
  getAccessToken(scopes: readonly string[]): Promise<string>;
}

export type GoogleClientIdStorage = Readonly<{
  load(): string | null | Promise<string | null>;
  save(clientId: string): void | Promise<void>;
  clear(): void | Promise<void>;
}>;

export type GoogleSheetsServiceConfiguration = Readonly<{
  clientId?: string;
  tokenProvider?: TokenProvider;
  tokenProviderFactory?: (clientId: string) => TokenProvider;
  clientIdStorage?: GoogleClientIdStorage | false;
}>;
~~~

Add `googleSheets?: GoogleSheetsServiceConfiguration` to `SpreadsheetServices`. Retain `googleTokenProviderFactory` with `/** @deprecated Use googleSheets.tokenProviderFactory. */` for one compatibility release.

- [ ] **Step 5: Implement typed errors and pure configuration**

`googleErrors.ts` exports the spec's complete union and class:

~~~ts
export type GoogleSheetsErrorCode =
  | "configuration_missing"
  | "incompatible_origin"
  | "origin_mismatch"
  | "invalid_client"
  | "invalid_sheet_url"
  | "gis_load_failed"
  | "popup_blocked"
  | "popup_closed"
  | "access_denied"
  | "api_not_enabled"
  | "sheet_not_found"
  | "rate_limited"
  | "network_failed"
  | "unknown";

export class GoogleSheetsError extends Error {
  constructor(
    readonly code: GoogleSheetsErrorCode,
    message: string,
    readonly recoverable: boolean,
    options?: ErrorOptions
  ) {
    super(message, options);
    this.name = "GoogleSheetsError";
  }
}

export function toGoogleSheetsError(error: unknown): GoogleSheetsError;
~~~

`toGoogleSheetsError` returns an existing typed error unchanged and otherwise creates `unknown` with a fixed safe message; it never incorporates arbitrary error text.

`googleConfiguration.ts` exports:

~~~ts
export type GoogleOAuthOriginAssessment =
  | { status: "eligible"; origin: string; registration: "unverified" }
  | { status: "blocked"; origin: string; reason: "opaque" | "file" | "insecure" | "ip_literal" };

export type GoogleClientIdValidation =
  | { valid: true; value: string }
  | { valid: false; message: string };

export type GoogleAuthSource =
  | { kind: "provider"; provider: TokenProvider }
  | {
      kind: "client";
      clientId: string;
      source: "managed" | "stored" | "session";
      factory: (clientId: string) => TokenProvider;
    }
  | { kind: "missing" };

export type EditableGoogleClientId = Readonly<{
  value: string;
  source: "stored" | "session";
}>;
~~~

Implement `validateGoogleClientId`, `assessGoogleOAuthOrigin`, and `resolveGoogleAuthSource(configuration, editableClientId, deprecatedFactory, builtInFactory)` with the tested precedence. Only a direct provider or non-built-in host factory bypasses built-in origin eligibility.

- [ ] **Step 6: Make the browser provider prepare before the import gesture**

Re-export the core `TokenProvider` type from `googleAuth.ts` and return this stronger type:

~~~ts
export type BrowserTokenProvider = TokenProvider & {
  prepare(): Promise<void>;
};

export function createBrowserTokenProvider(clientId: string): BrowserTokenProvider;
~~~

`prepare()` loads and validates GIS without opening a popup. After preparation, `getAccessToken` must call `initTokenClient(...).requestAccessToken()` before its first asynchronous boundary when no valid cached token exists. A programmatic caller that did not prepare remains backward compatible by awaiting `prepare()` internally, while the standalone controller always prepares first to preserve user activation. Map GIS callback/error-callback values to `GoogleSheetsError` codes and retain the existing 5-minute expiry safety window.

- [ ] **Step 7: Type the Sheets API failures**

Wrap token/network failures with `toGoogleSheetsError`, throw `invalid_sheet_url` before auth, and map HTTP responses to the tested codes. For 403, parse only `error.status` and known `error.errors[].reason` strings to distinguish API disabled from access denied, then discard the body.

Export errors/config/auth/import helpers from `src/entry/google.ts`.

- [ ] **Step 8: Run and commit connector primitives**

Run:

~~~bash
corepack pnpm exec vitest run src/lib/googleConfiguration.test.ts src/lib/googleAuth.test.ts src/lib/googleSheets.test.ts
~~~

Expected: PASS with typed, sanitized errors and synchronous post-prepare token requests.

~~~bash
git add src/core/workbook/services.ts src/lib/googleErrors.ts src/lib/googleConfiguration.ts src/lib/googleConfiguration.test.ts src/lib/googleAuth.ts src/lib/googleAuth.test.ts src/lib/googleSheets.ts src/lib/googleSheets.test.ts src/entry/google.ts
git commit -m "feat(google): add typed import configuration and errors"
~~~

### Task 8: Build the Google Setup Dialog and Per-Instance Controller

**Files:**
- Create: `src/react/browserGoogleClientIdStorage.ts`
- Create: `src/react/browserGoogleClientIdStorage.test.ts`
- Create: `src/react/useGoogleSheetsImport.ts`
- Create: `src/react/useGoogleSheetsImport.test.tsx`
- Create: `src/components/GoogleSheetsImportDialog.tsx`
- Create: `src/components/GoogleSheetsImportDialog.test.tsx`
- Modify: `src/test/setup.ts`
- Modify: `src/App.tsx`
- Modify: `src/App.test.tsx`
- Modify: `src/App.css`
- Modify: `src/components/Toolbar.tsx`
- Modify: `src/Spreadsheet.controlled.test.tsx`
- Modify: `src/react/Spreadsheet.types.test.tsx`
- Modify: `src/react/index.ts`
- Modify: `src/entry/core.ts`
- Modify: `src/entry/react.ts`
- Modify: `src/entry/entrypoints.test.ts`
- Modify: `src/index.ts`

**Interfaces:**
- Consumes: Task 7 configuration/errors/provider/import APIs and existing `session.replaceWorkbook`.
- Produces: standalone client-ID storage, `useGoogleSheetsImport`, `GoogleSheetsImportDialog`, nested public service types, and the **Import Google Sheet** File action.

- [ ] **Step 1: Write failing storage and controller tests**

Use the versioned key and assert storage contains only the normalized client ID:

~~~ts
export const GOOGLE_CLIENT_ID_STORAGE_KEY = "javascript-spreadsheet.google-client-id.v1";

it("loads, saves, and clears only the public client id", async () => {
  const storage = new MemoryStorage();
  const adapter = createBrowserGoogleClientIdStorage(storage);
  await adapter.save("123-abc.apps.googleusercontent.com");
  expect(storage.getItem(GOOGLE_CLIENT_ID_STORAGE_KEY)).toBe("123-abc.apps.googleusercontent.com");
  await expect(adapter.load()).resolves.toBe("123-abc.apps.googleusercontent.com");
  await adapter.clear();
  await expect(adapter.load()).resolves.toBeNull();
});
~~~

Hook tests must cover direct provider bypass, managed ID over stored ID, nested factory over deprecated factory, built-in origin block, storage load/save/clear failure, per-instance provider cache, preparation, single-flight submit, retry after failure, close invalidation, and late-result suppression.

- [ ] **Step 2: Write failing dialog accessibility tests**

Add native dialog shims in `src/test/setup.ts`, then assert:

~~~ts
if (!("showModal" in HTMLDialogElement.prototype)) {
  Object.defineProperty(HTMLDialogElement.prototype, "showModal", {
    configurable: true,
    value(this: HTMLDialogElement) {
      this.setAttribute("open", "");
    }
  });
}
if (!("close" in HTMLDialogElement.prototype)) {
  Object.defineProperty(HTMLDialogElement.prototype, "close", {
    configurable: true,
    value(this: HTMLDialogElement) {
      this.removeAttribute("open");
      this.dispatchEvent(new Event("close"));
    }
  });
}
~~~

~~~ts
render(<GoogleSheetsImportDialog controller={setupController()} opener={openerRef} />);
const dialog = screen.getByRole("dialog", { name: "Import Google Sheet" });
expect(dialog).toHaveAttribute("aria-modal", "true");
expect(screen.getByLabelText("Google OAuth client ID")).toHaveFocus();
await user.click(screen.getByRole("button", { name: "Save and continue" }));
expect(saveClientId).toHaveBeenCalledWith("123-abc.apps.googleusercontent.com");
~~~

Cover initial focus for missing/ready/blocked states, exact origin display/copy, eligible-unverified wording, no localhost recommendation for the LAN IP, Change/Forget, first-invalid-field focus, `aria-busy`, disabled duplicate submit, inline `role="alert"`, Escape/Cancel, and opener focus restoration.

- [ ] **Step 3: Run storage/controller/dialog tests and verify red**

Run:

~~~bash
corepack pnpm exec vitest run src/react/browserGoogleClientIdStorage.test.ts src/react/useGoogleSheetsImport.test.tsx src/components/GoogleSheetsImportDialog.test.tsx
~~~

Expected: FAIL because the three modules do not exist.

- [ ] **Step 4: Implement standalone storage and controller contracts**

Create this controller surface:

~~~ts
export type GoogleSheetsImportPhase =
  | "closed"
  | "loading"
  | "setup"
  | "blocked"
  | "preparing"
  | "ready"
  | "authorizing"
  | "importing"
  | "error";

export type GoogleSheetsImportController = Readonly<{
  open: boolean;
  phase: GoogleSheetsImportPhase;
  origin: string;
  originAssessment: GoogleOAuthOriginAssessment;
  clientIdSource: "managed" | "stored" | "session" | "missing";
  clientIdDraft: string;
  sheetDraft: string;
  error?: GoogleSheetsError;
  setClientIdDraft(value: string): void;
  setSheetDraft(value: string): void;
  openDialog(): void;
  closeDialog(): void;
  saveClientId(): Promise<void>;
  forgetClientId(): Promise<void>;
  changeClientId(): void;
  importSheet(): void;
}>;

export function useGoogleSheetsImport(options: Readonly<{
  configuration?: GoogleSheetsServiceConfiguration;
  deprecatedTokenProviderFactory?: (clientId: string) => TokenProvider;
  origin: string;
  onImported(result: GoogleSheetsImportResult): void | Promise<void>;
  onError?(error: GoogleSheetsError): void;
}>): GoogleSheetsImportController;
~~~

The hook loads storage once per instance, caches providers by direct-provider/factory identity plus client ID, preloads providers with `prepare`, and increments an attempt generation on submit and close. `importSheet()` must validate `sheetDraft`, synchronously call `importWorkbookFromGoogleSheets(sheetDraft, provider)` in the button gesture, then attach handlers; do not await storage first. While pending it ignores duplicate calls. A matching generation alone may invoke `onImported`. After successful `onImported`, close the dialog, clear only the sheet draft/error, and retain configured client-ID state.

- [ ] **Step 5: Implement the native modal presentation**

`GoogleSheetsImportDialog` accepts only `{ controller, opener }`. Keep it mounted per Spreadsheet instance and use an effect to call `showModal()`/`close()` based on `controller.open`. Render:

- title **Import Google Sheet** and descriptive replacement warning;
- exact origin with Copy;
- setup checklist and client-ID field when missing/session-editing;
- read-only source plus Change/Forget only for stored IDs;
- blocked or eligible-unverified origin status;
- sheet URL/ID field in ready/error states;
- **Save and continue**, **Import and replace workbook**, Retry, Cancel, and Close as state permits;
- `aria-busy` and live progress for preparing/authorizing/importing;
- internally scrolling content with the existing neutral theme tokens.

Use the native cancel event to call `closeDialog`; focus missing config, ready URL, or blocked Close on open, and restore `opener.current` on close.

- [ ] **Step 6: Integrate with Spreadsheet and standalone App**

Rename the toolbar action to **Import Google Sheet** and make it call `controller.openDialog()`. Keep the action in File. Remove `window.prompt`, direct env lookup, and old promise chain from `SpreadsheetWorkbook`.

On controller success, run:

~~~ts
const result = session.replaceWorkbook(imported.workbook, {
  history: "preserve",
  origin: "import"
});
if (result.status === "committed") {
  resetAfterWorkbookReplacement(`Imported ${imported.spreadsheetTitle}`);
}
~~~

Map controller failures to the existing host callback without exposing causes:

~~~ts
onError={(error) => invokeHostCallback(onError, {
  code: `service.google.${error.code}`,
  message: error.message,
  recoverable: error.recoverable
})}
~~~

Keep the dialog open with its inline error; `onError` is additional host observability, not a replacement for user feedback.

Standalone `App` supplies:

~~~tsx
const standaloneGoogleSheets = {
  clientId: import.meta.env.VITE_GOOGLE_CLIENT_ID as string | undefined,
  clientIdStorage: getDefaultBrowserGoogleClientIdStorage()
} satisfies GoogleSheetsServiceConfiguration;

return (
  <Spreadsheet
    className="js-spreadsheet-standalone"
    style={{ height: "100dvh", minHeight: 0 }}
    services={{ googleSheets: standaloneGoogleSheets }}
  />
);
~~~

Do not merge this storage with workbook storage. Embedded instances persist nothing unless their host passes `clientIdStorage`.

- [ ] **Step 7: Update public type and compatibility exports**

Export `GoogleClientIdStorage`, `GoogleSheetsServiceConfiguration`, and `TokenProvider` from core/react/root type surfaces. Export config/error helpers only from `js-spreadsheet/connectors/google`. Keep dialog/controller internal. Extend entrypoint tests to import the optional Google connector while proving `js-spreadsheet/core` evaluates in Node without DOM access.

Replace the old controlled factory test with direct-provider, nested-factory, deprecated-factory, and per-instance-cache tests. Keep `features.googleSheets=false` hiding the renamed action.

- [ ] **Step 8: Run the full Google/UI test slice**

Run:

~~~bash
corepack pnpm exec vitest run src/lib/googleConfiguration.test.ts src/lib/googleAuth.test.ts src/lib/googleSheets.test.ts src/react/browserGoogleClientIdStorage.test.ts src/react/useGoogleSheetsImport.test.tsx src/components/GoogleSheetsImportDialog.test.tsx src/App.test.tsx src/Spreadsheet.controlled.test.tsx src/react/Spreadsheet.types.test.tsx src/entry/entrypoints.test.ts
~~~

Expected: PASS for missing config, blocked origin, managed/stored/session lifecycle, valid import, failure preservation, cancellation, focus, and compatibility.

- [ ] **Step 9: Commit the setup/import experience**

~~~bash
git add src/core/workbook/services.ts src/react/browserGoogleClientIdStorage.ts src/react/browserGoogleClientIdStorage.test.ts src/react/useGoogleSheetsImport.ts src/react/useGoogleSheetsImport.test.tsx src/components/GoogleSheetsImportDialog.tsx src/components/GoogleSheetsImportDialog.test.tsx src/test/setup.ts src/App.tsx src/App.test.tsx src/App.css src/components/Toolbar.tsx src/Spreadsheet.controlled.test.tsx src/react/Spreadsheet.types.test.tsx src/react/index.ts src/entry/core.ts src/entry/react.ts src/entry/google.ts src/entry/entrypoints.test.ts src/index.ts
git commit -m "feat(google): guide standalone sheet import setup"
~~~

### Task 9: Fix DataTable Top-Layer Menus and First-Click Tree Toggles

**Files:**
- Modify: `src/react/DataTable.tsx`
- Modify: `src/react/DataTableColumnMenu.tsx`
- Modify: `src/react/DataTableCell.tsx`
- Modify: `src/react/DataTable.test.tsx`
- Modify: `src/styles/data-table.css`
- Modify: `src/styles/styleIsolation.test.tsx`
- Modify: `src/test/setup.ts`

**Interfaces:**
- Consumes: current DataTable column trigger state and tree-row dispatch.
- Produces: native auto-popover column menus with fixed clamped placement and one-click tree toggles that do not enter grid selection.

- [ ] **Step 1: Write failing menu and tree interaction tests**

Add deterministic `showPopover` and `hidePopover` shims to test setup:

~~~ts
if (!("showPopover" in HTMLElement.prototype)) {
  Object.defineProperty(HTMLElement.prototype, "showPopover", {
    configurable: true,
    value(this: HTMLElement) {
      this.setAttribute("data-popover-open", "true");
    }
  });
}
if (!("hidePopover" in HTMLElement.prototype)) {
  Object.defineProperty(HTMLElement.prototype, "hidePopover", {
    configurable: true,
    value(this: HTMLElement) {
      this.removeAttribute("data-popover-open");
    }
  });
}
~~~

In `DataTable.test.tsx`, assert:

~~~ts
const trigger = screen.getByRole("button", { name: "Column options for Name" });
await user.click(trigger);
const menu = screen.getByRole("menu", { name: "Name column menu" });
expect(menu).toHaveAttribute("popover", "auto");
expect(menu.closest(".js-spreadsheet-data-table")).not.toBeNull();
expect(showPopover).toHaveBeenCalledTimes(1);
fireEvent.keyDown(menu, { key: "Escape" });
expect(trigger).toHaveFocus();
~~~

Mock the trigger at each viewport edge and assert fixed `top/left` remain within an 8px margin. For the tree regression:

~~~ts
const ref = createRef<DataTableHandle>();
renderTable({ ref, getSubRows: (row) => row.children, defaultState: { expandedRowIds: ["e1"] } });
const selectionBefore = ref.current?.getSelection();
const toggle = screen.getByRole("button", { name: "Collapse Ada" });
fireEvent.pointerDown(toggle);
fireEvent.click(toggle);
expect(screen.queryByRole("gridcell", { name: "e1-child Name" })).not.toBeInTheDocument();
expect(ref.current?.getSelection()).toEqual(selectionBefore);
~~~

- [ ] **Step 2: Run DataTable tests and verify red**

Run:

~~~bash
corepack pnpm exec vitest run src/react/DataTable.test.tsx src/styles/styleIsolation.test.tsx
~~~

Expected: FAIL because the menu is an absolutely positioned descendant and pointerdown reaches the grid.

- [ ] **Step 3: Implement the native auto popover**

Change DataTable open state to retain both column ID and trigger element:

~~~ts
type OpenColumnMenu = Readonly<{
  columnId: string;
  trigger: HTMLButtonElement;
}>;
~~~

Pass `anchor`, `onClose`, and `onAnnouncement` to `DataTableColumnMenu`. Its root becomes:

~~~tsx
<div
  ref={menuRef}
  popover="auto"
  className="js-spreadsheet-data-table__column-menu"
  role="menu"
  aria-label={`${label} column menu`}
  style={{ position: "fixed", top: position.top, left: position.left }}
>
~~~

On mount, measure anchor/menu, clamp to 8px viewport edges, call `showPopover()`, and focus the first enabled control. Listen for native `toggle`, capture-phase viewport scroll, resize, column hiding, and trigger disconnect; close once, call `hidePopover()` when necessary, clear DataTable state, and restore trigger focus unless it disconnected.

Update CSS to remove `inset`, container-query positioning, and absolute placement. Set `margin: 0`, `width: min(248px, calc(100vw - 16px))`, and `max-height: min(520px, calc(100dvh - 16px))`; keep DataTable root `overflow: hidden`.

- [ ] **Step 4: Stop tree toggle pointerdown before the grid sees it**

Add:

~~~tsx
onPointerDown={(event) => {
  event.preventDefault();
  event.stopPropagation();
}}
onMouseDown={(event) => {
  event.preventDefault();
  event.stopPropagation();
}}
~~~

Keep click's existing `stopPropagation` and dispatch exactly once. Increase the tree toggle's minimum width/height to `30px` without changing row height.

- [ ] **Step 5: Run and commit DataTable regressions**

Run:

~~~bash
corepack pnpm exec vitest run src/react/DataTable.test.tsx src/react/DataTable.styles.test.tsx src/styles/styleIsolation.test.tsx
~~~

Expected: PASS for edge placement, scope, Escape/outside close, focus restoration, trigger removal, one-click collapse/expand, and no selection side effects.

~~~bash
git add src/react/DataTable.tsx src/react/DataTableColumnMenu.tsx src/react/DataTableCell.tsx src/react/DataTable.test.tsx src/styles/data-table.css src/styles/styleIsolation.test.tsx src/test/setup.ts
git commit -m "fix(datatable): keep menus visible and toggles immediate"
~~~

### Task 10: Add External-Server Browser Coverage and Update Documentation

**Files:**
- Create: `tests/datatable.spec.ts`
- Modify: `tests/spreadsheet.spec.ts`
- Modify: `playwright.config.ts`
- Modify: `README.md`
- Modify: `docs/google-sheets-connector.md`
- Modify: `docs/getting-started.md`
- Modify: `docs/embedding.md`
- Modify: `docs/features.md`

**Interfaces:**
- Consumes: all implemented user journeys and the existing preview on port `4173`.
- Produces: browser regressions that can reuse one external server and accurate clone/embed/Google documentation.

- [ ] **Step 1: Make Playwright external-server aware**

Change configuration to:

~~~ts
const externalBaseURL = process.env.E2E_BASE_URL?.trim();

export default defineConfig({
  testDir: "./tests",
  timeout: 30_000,
  use: {
    baseURL: externalBaseURL || "http://127.0.0.1:5173",
    trace: "on-first-retry"
  },
  ...(externalBaseURL
    ? {}
    : {
        webServer: {
          command: "pnpm dev",
          url: "http://127.0.0.1:5173",
          reuseExistingServer: !process.env.CI
        }
      }),
  projects: [{ name: "chromium", use: { ...devices["Desktop Chrome"] } }]
});
~~~

Run:

~~~bash
corepack pnpm exec playwright test --list
~~~

Expected: the suite lists successfully. With `E2E_BASE_URL` set, process inspection later must show no listener on 5173.

- [ ] **Step 2: Add spreadsheet browser regressions**

Add tests with these observable contracts:

~~~ts
test("keeps a newly added sheet visible and focused in the viewport", async ({ page }) => {
  await page.setViewportSize({ width: 1280, height: 800 });
  await page.goto("/");
  await page.getByRole("button", { name: "Add sheet", exact: true }).click();
  const tab = page.getByRole("tab", { name: "Sheet2", exact: true });
  await expect(tab).toBeVisible();
  await expect(tab).toHaveAttribute("aria-selected", "true");
  await expect(page.getByRole("gridcell", { name: "A1", exact: true })).toBeFocused();
  await expect(page.getByLabel("Status")).toContainText("Added Sheet2");
  const bottom = await page.getByRole("tablist", { name: "Sheet tabs" }).evaluate((node) => node.getBoundingClientRect().bottom);
  expect(bottom).toBeLessThanOrEqual(800);
});

test("keeps an embedded spreadsheet inside its fixed host", async ({ page }) => {
  await page.setViewportSize({ width: 1280, height: 900 });
  await page.goto("/embed");
  const host = page.getByTestId("embedding-host");
  const spreadsheet = page.locator('[data-js-spreadsheet-root="workbook"]');
  expect((await spreadsheet.boundingBox())!.height).toBeLessThanOrEqual((await host.boundingBox())!.height);
});

test("offers table-aware left and right insertion from ribbon and column headers", async ({ page }) => {
  await page.goto("/");
  await page.getByRole("gridcell", { name: "A1", exact: true }).click();
  await page.getByRole("button", { name: "Insert columns options" }).click();
  await page.getByRole("menuitem", { name: "Insert column right" }).click();
  await expect(page.getByLabel("Status")).toContainText("Inserted 1 column right");
  await page.getByRole("columnheader", { name: "Column B" }).click({ button: "right" });
  await page.getByRole("menu", { name: "Column B context menu" }).getByRole("menuitem", { name: "Insert column left" }).click();
  await expect(page.getByLabel("Status")).toContainText("Inserted 1 column left");
});

test("explains Google setup on the LAN IP instead of appearing inert", async ({ page }) => {
  await page.goto("/");
  await openRibbonTab(page, "File");
  await page.getByRole("button", { name: "Import Google Sheet" }).click();
  const dialog = page.getByRole("dialog", { name: "Import Google Sheet" });
  await expect(dialog).toContainText("http://192.168.6.232:4173");
  await expect(dialog).toContainText(/HTTPS DNS origin/i);
  await expect(dialog).not.toContainText(/localhost/i);
});
~~~

Also add narrow `390×844` coverage for Add sheet/tab bounds, header keyboard menu via `Shift+F10`, and dialog internal scrolling at 200% zoom.

- [ ] **Step 3: Add DataTable browser regressions**

Create `tests/datatable.spec.ts`:

~~~ts
test("keeps the rightmost column menu inside desktop and narrow viewports", async ({ page }) => {
  for (const viewport of [{ width: 1280, height: 800 }, { width: 390, height: 844 }]) {
    await page.setViewportSize(viewport);
    await page.goto("/datatable");
    const grid = page.getByRole("grid", { name: "Employee directory" });
    await grid.evaluate((node) => { node.scrollLeft = node.scrollWidth; });
    await page.getByRole("button", { name: "Column options for Active" }).click();
    const box = await page.getByRole("menu", { name: "Active column menu" }).boundingBox();
    expect(box).not.toBeNull();
    expect(box!.x).toBeGreaterThanOrEqual(8);
    expect(box!.y).toBeGreaterThanOrEqual(8);
    expect(box!.x + box!.width).toBeLessThanOrEqual(viewport.width - 8);
    expect(box!.y + box!.height).toBeLessThanOrEqual(viewport.height - 8);
    await page.keyboard.press("Escape");
    await expect(page.getByRole("button", { name: "Column options for Active" })).toBeFocused();
  }
});

test("collapses a tree row on the first narrow-screen click without scrolling", async ({ page }) => {
  await page.setViewportSize({ width: 390, height: 844 });
  await page.goto("/datatable");
  const grid = page.getByRole("grid", { name: "Employee directory" });
  const before = await grid.evaluate((node) => node.scrollLeft);
  await page.getByRole("button", { name: "Collapse Engineering team" }).click();
  await expect(page.getByRole("gridcell", { name: /e-ada Name/ })).toHaveCount(0);
  expect(await grid.evaluate((node) => node.scrollLeft)).toBe(before);
  await page.getByRole("button", { name: "Expand Engineering team" }).click();
  await expect(page.getByRole("gridcell", { name: /e-ada Name/ })).toBeVisible();
});
~~~

- [ ] **Step 4: Rewrite documentation around the actual product behavior**

Make all five documentation surfaces consistent:

- File → **Import Google Sheet** is a one-time, read-only workbook replacement.
- A fresh clone opens the setup dialog when configuration is absent.
- Built-in browser OAuth needs an eligible, exactly registered origin; the current LAN IP is intentionally diagnosed as incompatible.
- Standalone stores only a public client ID under `javascript-spreadsheet.google-client-id.v1`.
- Embedded apps pass `services.googleSheets.tokenProvider`, `clientId`, `tokenProviderFactory`, and optional `clientIdStorage`; embedded storage is opt-in.
- Embedded spreadsheet hosts provide at least `420px` block size and remain responsible for their own component height.
- Document GIS/Sheets CSP origins and iframe popup permission.
- Document `features.googleSheets=false`.
- Document the one-server browser command exactly as shown below.
- Remove every stale `Data → Link Google Sheet`, env-only, and continuous-link claim from README and docs.

- [ ] **Step 5: Build the new worktree and replace—not duplicate—the preview server**

Run the production build first:

~~~bash
corepack pnpm run build
~~~

Expected: `dist/app` is current. Stop the existing process bound to port 4173, then launch exactly one preview from this worktree with:

~~~bash
corepack pnpm exec vite preview --host 0.0.0.0 --port 4173 --strictPort
~~~

Use the execution environment's long-running process/session mechanism. Do not start port 5173 and do not run two 4173 listeners.

- [ ] **Step 6: Run browser tests against the IP server**

Run:

~~~bash
E2E_BASE_URL=http://192.168.6.232:4173 corepack pnpm run test:e2e
~~~

Expected: PASS. Process inspection shows one listener on `0.0.0.0:4173` and none on 5173.

Verify that explicitly:

~~~bash
ss -ltnp | rg ':4173|:5173'
~~~

Expected: exactly one `0.0.0.0:4173` listener and no `:5173` line.

- [ ] **Step 7: Commit browser coverage and docs**

~~~bash
git add playwright.config.ts tests/spreadsheet.spec.ts tests/datatable.spec.ts README.md docs/google-sheets-connector.md docs/getting-started.md docs/embedding.md docs/features.md
git commit -m "test: cover spreadsheet setup and structure workflows"
~~~

### Task 11: Run the Release Gate, Open the PR, and Resolve Review

**Files:**
- Verify: every file changed in Tasks 1–10
- Modify only if a check or actionable review identifies a concrete defect

**Interfaces:**
- Consumes: the complete branch and one-server preview.
- Produces: a reviewed pull request ready for squash merge into `main`.

- [ ] **Step 1: Run the complete deterministic gate**

Run:

~~~bash
corepack pnpm run typecheck
corepack pnpm test
corepack pnpm run build
corepack pnpm run build:lib
E2E_BASE_URL=http://192.168.6.232:4173 corepack pnpm run test:e2e
~~~

Expected: every command exits 0. If a failure occurs, add a focused regression test, make the smallest correction, rerun the failing slice, and commit that correction separately before rerunning this entire gate.

- [ ] **Step 2: Verify package boundaries and the final diff**

Run:

~~~bash
git diff --check origin/main...HEAD
git status --short --branch
git diff --stat origin/main...HEAD
git log --oneline origin/main..HEAD
~~~

Expected: no whitespace errors, no uncommitted files, no secrets/tokens/client IDs, and only the approved specification, plan, implementation, tests, and docs.

- [ ] **Step 3: Push and create the new pull request**

Run:

~~~bash
git push -u origin feat/spreadsheet-ux-followup
gh pr create --base main --head feat/spreadsheet-ux-followup --title "feat: complete spreadsheet setup and structure editing" --body $'## Summary\n- add guided Google Sheets import setup for standalone and embedded apps\n- keep new sheets visible and focused and add left/right table-aware column insertion\n- fix DataTable menu clipping and first-click tree toggles\n\n## Verification\n- pnpm run typecheck\n- pnpm test\n- pnpm run build\n- pnpm run build:lib\n- E2E_BASE_URL=http://192.168.6.232:4173 pnpm run test:e2e'
~~~

Expected: a new PR URL whose diff is based only on `origin/main`, not the already merged DataTable branch.

- [ ] **Step 4: Monitor CI and collect actionable review comments**

Run:

~~~bash
gh pr checks --watch
gh pr view --comments
~~~

Inspect unresolved review threads through the GitHub review API or connected review tooling. Classify each comment as actionable, already addressed, incorrect with evidence, or out of scope. Do not make speculative changes for non-actionable suggestions.

- [ ] **Step 5: Fix actionable findings with focused tests and commits**

For each valid finding: reproduce it, add the narrowest failing test, implement the correction, run the focused test plus affected suite, commit with `fix: <specific behavior>`, push, and wait for checks again. Repeat until checks are green and no actionable thread remains.

- [ ] **Step 6: Perform the final squash merge**

Only after CI and actionable review are clear, run:

~~~bash
gh pr merge --squash --delete-branch
~~~

Expected: PR state `MERGED`, `origin/main` contains the squash commit, and the single preview can then be rebuilt from main if the user wants to keep it running.

## Requirement-to-Task Map

| Approved requirement | Owning task(s) |
| --- | --- |
| Strict table-safe structural core | 1–4 |
| Left/right ribbon, cell, and header actions | 5 |
| Immediate Add sheet, focused A1, visible tab | 6 |
| Standalone viewport vs embedded host sizing | 6, 10 |
| Typed Google config/errors and browser auth | 7 |
| Guided setup dialog, storage, host injection | 8 |
| DataTable menu clipping and first-click toggle | 9 |
| One external port-4173 browser suite | 10 |
| Accurate clone/embed documentation | 10 |
| Full pressure gate, PR review, squash merge | 11 |
