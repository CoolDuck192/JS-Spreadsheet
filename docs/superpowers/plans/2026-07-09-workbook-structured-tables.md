# Workbook Structured Tables Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add Excel-style structured tables to the workbook model, expose each table as a live `DataTable` session, provide a contextual spreadsheet Table tab, and round-trip real native Excel tables without losing formulas, filters, typed values, or logical row identity.

**Architecture:** `WorkbookModel` advances to version 2 and stores workbook-wide table metadata while cells, formats, validation, comments, hyperlinks, and formulas remain in `SheetModel`. Pure structured-table reducers run only through the atomic `WorkbookSession` command path; `WorkbookTableSession` resolves stable IDs to the current worksheet on every snapshot and never copies canonical row values. ExcelJS owns native Excel table relationships, ranges, names, styles, and totals, while a narrow OOXML supplement preserves calculated-column and filter XML that ExcelJS 4.4 does not fully round-trip.

**Tech Stack:** TypeScript 6, React 19, Vitest 4, Testing Library, Playwright, HyperFormula 3.3, ExcelJS 4.4, fast-check, fflate, @xmldom/xmldom, Vite 8.

## Global Constraints

- The release remains React-only; Laravel, Web Components, vanilla mounting, Vue, and other framework bindings remain excluded.
- The workbook table is a live worksheet projection. Never introduce a second canonical row-value collection.
- All table mutations are atomic `WorkbookCommand` reductions and create at most one history entry.
- Table and column IDs are non-reused UUID-like IDs in app-native state.
- Row IDs survive insert, delete, physical sort, filter, resize, and undo while the workbook session exists.
- XLSX does not persist opaque app IDs. A configured key column preserves row identity across re-import; otherwise row IDs regenerate.
- Workbook physical sorting uses evaluated typed values and never raw formula strings or formatted display strings.
- Structural formula changes use the token-aware reference service. Regex replacement of A1-looking text is prohibited.
- Table filtering changes visibility only; it never copies, deletes, or page-locally substitutes source rows.
- Table names use Excel-compatible syntax and workbook-wide case-insensitive uniqueness so app-native tables never become unexportable.
- CSV is a data-only export. It never emits executable formulas, and potentially executable text is escaped by the policy in Task 8.
- XLSX bytes are untrusted input. Archive, relationship, and XML safety checks run before ExcelJS or the OOXML supplement parses content.
- Styles stay scoped below spreadsheet or DataTable roots and use the approved white, slate, and muted-green tokens.
- No table structural control is exposed until the prerequisite correctness gate in Task 0 passes.
- Prefix every project command in this plan with `corepack pnpm`.
- Do not publish a partial workbook-table feature.

## Required Upstream Interfaces

This plan starts after the foundation and local DataTable plans have created these exact modules:

- `src/core/commands/types.ts` exports `CommandEnvelope`, `CommandResult`, and `TableIssue`.
- `src/core/workbook/WorkbookSession.ts` exports `WorkbookSession`, `WorkbookSnapshot`, `WorkbookCommandResult`, `WorkbookDiagnosticEvent`, and `createWorkbookSession`.
- `src/core/workbook/commands.ts` exports `WorkbookCommand` and `reduceWorkbookCommand`.
- `src/table/core/types.ts` exports `ColumnDef`, `TableSession`, `TableViewSnapshot`, `TableIntent`, `TableCellSnapshot`, `TableSelection`, and the DOM-free `ExportArtifact` (`bytes: Uint8Array`, `mediaType: string`, `fileName: string`).
- `src/table/core/query.ts` exports `FilterExpression` and `QueryScalar`.
- `src/react/DataTable.tsx` exports the session-backed React `DataTable`; this is the canonical internal source module, while the final public React import is `js-spreadsheet/react` (and the root React alias).
- `src/App.tsx` exports the controlled/session-backed `Spreadsheet` while preserving standalone behavior.

These source paths are implementation boundaries, not competing public package paths. The packaging plan exposes headless workbook/table sessions and types from `js-spreadsheet/core`, React components/hooks/props from `js-spreadsheet/react`, and the same React surface from the root `js-spreadsheet` alias. `src/index.ts` remains only a source-compatibility barrel until those final entrypoints are built.

The workbook session must expose the foundation contract:

~~~ts
export interface WorkbookSession {
  getSnapshot(): WorkbookSnapshot;
  getCellEvaluation(sheetId: string, address: string): ComputedCellValue;
  subscribe(listener: () => void): () => void;
  subscribeDiagnostics(listener: (event: WorkbookDiagnosticEvent) => void): () => void;
  dispatch(command: WorkbookCommand | CommandEnvelope<WorkbookCommand>): WorkbookCommandResult;
  replaceWorkbook(workbook: WorkbookModel, options?: { history?: "reset" | "preserve"; origin?: "external" | "import" }): WorkbookCommandResult;
  destroy(): void;
}
~~~

Structured-table work extends this interface only with `table(tableId)` in Task 5. It must preserve `subscribeDiagnostics` and `replaceWorkbook` unchanged, including diagnostic cleanup on destroy and the foundation's replacement/history semantics.

`WorkbookTableSession.getCell` reads stored/formula/format metadata from the current `WorkbookSnapshot.workbook` and uses `getCellEvaluation` only for the typed computed value. It derives display text through the existing display-format service.

The structured-table subsystem extends `WorkbookCommand` with this exact union:

~~~ts
export type StructuredTableCommand =
  | { type: "table.create"; sheetId: string; range: CellRange; name?: string; headerRow?: boolean; totalsRow?: boolean; style?: TableStyle }
  | { type: "table.rename"; tableId: string; name: string }
  | { type: "table.renameColumn"; tableId: string; columnId: string; name: string }
  | { type: "table.resize"; tableId: string; range: CellRange }
  | { type: "table.setHeaderRow"; tableId: string; enabled: boolean }
  | { type: "table.setTotalsRow"; tableId: string; enabled: boolean }
  | { type: "table.setTotalsFunction"; tableId: string; columnId: string; aggregate: TableAggregate }
  | { type: "table.setStyle"; tableId: string; style: TableStyle }
  | { type: "table.setKeyColumn"; tableId: string; columnId?: string }
  | { type: "table.setCalculatedColumn"; tableId: string; columnId: string; formula?: string }
  | { type: "table.setFilter"; tableId: string; filter?: FilterExpression }
  | { type: "table.sort"; tableId: string; sorting: readonly TableSort[] }
  | { type: "table.insertRows"; tableId: string; count: number; beforeRowId?: string; afterRowId?: string }
  | { type: "table.deleteRows"; tableId: string; rowIds: readonly string[] }
  | { type: "table.editCells"; tableId: string; edits: readonly { rowId: string; columnId: string; rawText: string }[] }
  | { type: "table.convertToRange"; tableId: string };
~~~

At most one insert anchor may be present; no anchor appends. Every command after creation addresses tables, columns, and rows by stable ID rather than a visible index. `table.setTotalsFunction` requires an enabled totals row unless `aggregate` is `"none"`; it writes/removes the matching totals formula and updates metadata in one reduction.

## File Structure

New files and their single responsibilities:

- `src/core/ids.ts` — random app IDs and deterministic key-derived XLSX row IDs.
- `src/core/workbook/tableNames.ts` — Excel-compatible table-name parsing, validation, and normalized uniqueness keys.
- `src/core/workbook/migrateWorkbook.ts` — validate persisted workbook versions and migrate version 1 to version 2.
- `src/core/workbook/structuredTables.ts` — table lookup, range math, metadata validation, and non-row table mutations.
- `src/core/workbook/structuredTableRows.ts` — body-row insertion, deletion, permutation, calculated-column filling, and cell-plane movement.
- `src/core/workbook/structuredTableFilter.ts` — typed `FilterExpression` evaluation for workbook rows.
- `src/table/workbook/WorkbookTableSession.ts` — live `WorkbookSession` to `TableSession` projection.
- `src/components/SpreadsheetTableTab.tsx` — contextual ribbon controls.
- `src/components/StructuredTableFilterPanel.tsx` — typed filter editor for workbook columns.
- `src/components/CalculatedColumnPanel.tsx` — calculated-column formula editor.
- `src/react/workbook/WorkbookTableView.tsx` — adaptive workbook-table panel composed from `DataTable`.
- `src/lib/xlsxTables.ts` — ExcelJS table-model conversion.
- `src/lib/structuredFormula.ts` — token-aware Excel structured-reference to internal A1 translation and reverse export translation.
- `src/lib/xlsxTableXml.ts` — narrowly scoped OOXML read/write supplement.
- `src/lib/xlsxSecurity.ts` — preflight ZIP limits, safe XML parsing, and relationship validation for untrusted XLSX bytes.
- `src/test/fixtures/xlsx/generated-sales-structured-table.xlsx` — deterministic ExcelJS sales-table fixture.
- `src/test/fixtures/xlsx/exceljs-issue-1669.xlsx` — pinned upstream native-table/filter fixture.
- `src/test/fixtures/xlsx/generated-sales-structured-table.expected.json` — sales-fixture expectations.
- `src/test/fixtures/xlsx/README.md` — fixture provenance and reproducible contents.

Existing files modified:

- `src/types.ts` — version 2 workbook and structured-table types.
- `src/lib/workbook.ts` — blank workbook and sheet lifecycle integration.
- `src/lib/persistence.ts` — call the versioned migrator.
- `src/lib/googleSheets.ts` — create version 2 workbooks with no tables.
- `src/lib/xlsx.ts` — orchestrate table import/export.
- `src/core/workbook/commands.ts` — add table command union and reducer routing.
- `src/core/workbook/WorkbookSession.ts` — supply IDs/evaluation and vend child table sessions.
- `src/components/Toolbar.tsx` — dynamic contextual Table tab and Insert Table action.
- `src/components/Grid.tsx` — table style/context and structured-filter visibility.
- `src/App.tsx` — create/open/export/convert table flows.
- `src/App.css` — scoped neutral table styles.
- `src/index.ts` — temporary source-compatibility exports; final package ownership is split between `js-spreadsheet/core` and `js-spreadsheet/react`.
- `package.json` and `pnpm-lock.yaml` — add test/XML dependencies.

---

### Task 0: Prove the Correctness Prerequisites

**Files:**
- Verify: `src/core/workbook/WorkbookSession.test.ts`
- Verify: `src/core/workbook/commands.test.ts`
- Verify: `src/lib/formulaReferences.test.ts`
- Verify: `src/lib/formulaEngine.test.ts`
- Verify: `src/lib/workbook.test.ts`
- Verify: `src/lib/validation.test.ts`

**Interfaces:**
- Consumes: atomic `WorkbookSession.dispatch`, `WorkbookSession.getCellEvaluation`, token-aware formula transforms, typed cell parsing, and typed range sort.
- Produces: a hard green gate. This task changes no code and creates no commit.

- [ ] **Step 1: Run the atomic dispatch tests**

Run:

~~~bash
corepack pnpm exec vitest run src/core/workbook/WorkbookSession.test.ts src/core/workbook/commands.test.ts
~~~

Expected: PASS, including back-to-back dispatches reducing against the latest revision and rejected commands preserving workbook object, revision, and history.

- [ ] **Step 2: Run formula transformation and evaluation tests**

Run:

~~~bash
corepack pnpm exec vitest run src/lib/formulaReferences.test.ts src/lib/formulaEngine.test.ts
~~~

Expected: PASS for quoted `"A1"` text, `LOG10`, scientific notation `1E10`, absolute references, ranges, sheet-qualified references, cross-sheet dependents, and typed formula results.

- [ ] **Step 3: Run typed sorting and validation tests**

Run:

~~~bash
corepack pnpm exec vitest run src/lib/workbook.test.ts src/lib/validation.test.ts
~~~

Expected: PASS for numeric/date/boolean sorting through evaluated values and validation through evaluated candidates.

- [ ] **Step 4: Record the gate result**

Do not edit a file. Continue only if Steps 1–3 all pass. If a prerequisite fails, return to its owning foundation task and make that prerequisite green before creating structured-table metadata.

### Task 1: Migrate WorkbookModel to Version 2 and Add ID/Name Services

**Files:**
- Create: `src/core/ids.ts`
- Create: `src/core/ids.test.ts`
- Create: `src/core/workbook/tableNames.ts`
- Create: `src/core/workbook/tableNames.test.ts`
- Create: `src/core/workbook/migrateWorkbook.ts`
- Create: `src/core/workbook/migrateWorkbook.test.ts`
- Modify: `src/types.ts`
- Modify: `src/lib/workbook.ts`
- Modify: `src/lib/persistence.ts`
- Modify: `src/lib/googleSheets.ts`
- Modify: `src/lib/xlsx.ts`
- Modify: model literals in `src/App.test.tsx`, `src/components/Grid.test.tsx`, `src/lib/formulaEngine.perf.test.ts`, `src/lib/persistence.test.ts`, `src/lib/workbook.test.ts`, and `src/lib/googleSheets.test.ts`

**Interfaces:**
- Consumes: `CellRange`, `CellContent`, and `FilterExpression`.
- Produces: `WorkbookModel.version === 2`, `WorkbookModel.tables`, `IdGenerator`, `createRandomId`, `createStableKeyRowId`, `validateExcelTableName`, `normalizeExcelTableNameKey`, and `migrateWorkbookModel`.

- [ ] **Step 1: Write failing ID and Excel table-name tests**

Add tests with these exact assertions:

~~~ts
import { describe, expect, it } from "vitest";
import { createRandomId, createStableKeyRowId } from "./ids";

describe("workbook IDs", () => {
  it("creates kind-prefixed non-empty random IDs", () => {
    expect(createRandomId("table")).toMatch(/^table-[0-9a-f-]+$/);
    expect(createRandomId("table-column")).toMatch(/^table-column-[0-9a-f-]+$/);
    expect(createRandomId("table-row")).toMatch(/^table-row-[0-9a-f-]+$/);
  });

  it("derives stable typed row IDs from XLSX keys", async () => {
    const first = await createStableKeyRowId("Employees", "Employee ID", 101);
    const repeat = await createStableKeyRowId("employees", "employee id", 101);
    const stringKey = await createStableKeyRowId("Employees", "Employee ID", "101");

    expect(first).toBe(repeat);
    expect(first).not.toBe(stringKey);
    expect(first).toMatch(/^table-row-key-[0-9a-f]{32}$/);
  });
});
~~~

In `tableNames.test.ts`, cover every valid/invalid boundary from Task 2: `_Sales`, `\Sales`, Unicode-letter names, 255/256 code points, whitespace, illegal characters, `R`/`C`, in-grid A1 (`A1`, `XFD1048576`), out-of-grid lookalikes (`XFE1`, `A1048577`), R1C1, NFKC-equivalent names, and locale-independent case-insensitive keys.

- [ ] **Step 2: Run the ID/name tests and verify red**

Run:

~~~bash
corepack pnpm exec vitest run src/core/ids.test.ts src/core/workbook/tableNames.test.ts
~~~

Expected: FAIL because the ID/name modules do not exist.

- [ ] **Step 3: Implement the ID and table-name services**

Create these exact exports in `src/core/ids.ts`:

~~~ts
import type { CellContent } from "../types";

export type IdKind = "table" | "table-column" | "table-row";
export type IdGenerator = (kind: IdKind) => string;

export function createRandomId(kind: IdKind): string;
export async function createStableKeyRowId(
  tableName: string,
  columnName: string,
  value: CellContent
): Promise<string>;
~~~

`createRandomId` must use `globalThis.crypto.randomUUID()` when available and `crypto.getRandomValues` to build RFC-4122 bytes otherwise. It must not use a module counter. `createStableKeyRowId` must normalize table/column names with `normalize("NFKC").trim().toLowerCase()` (never locale-dependent casing), encode the key type separately from the value, hash with SHA-256, and return the first 16 bytes as 32 lowercase hexadecimal characters. Add composed/decomposed Unicode, Turkish-I, numeric-vs-text key, and repeated-call tests proving the same inputs produce the same ID regardless of host locale.

Create `tableNames.ts` with the `ExcelTableNameValidation`, `validateExcelTableName`, and `normalizeExcelTableNameKey` signatures specified in Task 2. Implement the exact Excel syntax, length, cell-reference, and normalization rules there so migration can consume them before `structuredTables.ts` exists.

- [ ] **Step 4: Write failing migration tests**

Cover:

~~~ts
it("migrates a valid version 1 workbook to version 2", () => {
  const migrated = migrateWorkbookModel(versionOneFixture);
  expect(migrated).toMatchObject({ version: 2, tables: [] });
});

it("preserves app-native structured IDs in version 2", () => {
  const migrated = migrateWorkbookModel(versionTwoFixture);
  expect(migrated?.tables[0]).toMatchObject({
    id: "table-fixed",
    columns: [{ id: "column-fixed" }],
    rowIds: ["row-fixed"]
  });
});

it.each([null, {}, { version: 99 }, malformedVersionTwoFixture])(
  "rejects invalid persisted input",
  value => expect(migrateWorkbookModel(value)).toBeNull()
);
~~~

- [ ] **Step 5: Run migration tests and verify red**

Run:

~~~bash
corepack pnpm exec vitest run src/core/workbook/migrateWorkbook.test.ts
~~~

Expected: FAIL because the migrator and version 2 model are absent.

- [ ] **Step 6: Define the version 2 table model**

Import `TableDataType` from `src/table/core/types.ts` and `TableSort` plus `FilterExpression` from `src/table/core/query.ts` as type-only imports. Do not redefine those shared contracts. Add:

~~~ts
export type TableAggregate =
  | "none"
  | "sum"
  | "average"
  | "count"
  | "countNumbers"
  | "min"
  | "max"
  | "standardDeviation"
  | "variance";

export type TableStyle = {
  theme: string;
  showFirstColumn?: boolean;
  showLastColumn?: boolean;
  showRowStripes?: boolean;
  showColumnStripes?: boolean;
};

export type StructuredTableColumn = {
  id: string;
  name: string;
  sheetColumn: number;
  dataType?: TableDataType;
  calculatedFormula?: string;
  totalsFunction?: TableAggregate;
  totalsLabel?: string;
};

export type StructuredTable = {
  id: string;
  name: string;
  sheetId: string;
  range: CellRange;
  headerRow: boolean;
  totalsRow: boolean;
  columns: readonly StructuredTableColumn[];
  rowIds: readonly string[];
  keyColumnId?: string;
  style?: TableStyle;
  sort?: readonly TableSort[];
  filter?: FilterExpression;
};
~~~

Change `WorkbookModel.version` to literal `2` and add:

~~~ts
tables: StructuredTable[];
~~~

Use `FilterExpression` from `src/table/core/query.ts`. Do not duplicate the query AST in `src/types.ts`.

- [ ] **Step 7: Implement the versioned migrator**

Create:

~~~ts
export function migrateWorkbookModel(value: unknown): WorkbookModel | null;
~~~

The migrator must:

1. Validate the existing sheet fields currently checked in `src/lib/persistence.ts`.
2. Convert version 1 to version 2 and add `tables: []`.
3. Validate every version 2 table range, string ID, string name, boolean flag, column, and row ID.
4. Require workbook-wide unique table IDs and Excel-valid, case-insensitively unique table names. Use the shared `validateExcelTableName`/normalized-name helper created earlier in this task; do not let persistence apply a looser rule than create, rename, import, or export.
5. Require in-bounds, non-overlapping ranges on existing sheets; `columns.length` equal to range width; every `sheetColumn` aligned in order to that range; `rowIds.length` equal to body height after header/totals rows; and unique column/row IDs.
6. Require `keyColumnId`, sort column IDs, aggregate/totals column references, and every filter AST column ID to belong to the same table. A column cannot have both `totalsFunction` and `totalsLabel`; a label is valid only when `totalsRow` is enabled and must equal the stored text in that physical totals cell.
7. Reject tables intersecting persisted merges or impossible header/totals geometry rather than allowing adapter indexing aliases.
8. Preserve valid version 2 IDs exactly.
9. Apply existing sheet default migrations for formats, dimensions, hidden state, filters, charts, merges, and protection.

Replace the private persistence type guard/migrator with a call to `migrateWorkbookModel`. Invalid storage continues to return `createBlankWorkbook()`.

- [ ] **Step 8: Update every workbook constructor**

Make `createBlankWorkbook()`, Google Sheets import, CSV sheet replacement, and XLSX import produce:

~~~ts
{
  version: 2,
  activeSheetId,
  sheets,
  namedRanges,
  tables: []
}
~~~

Update typed test literals from `version: 1` to `version: 2` and add `tables: []`. Keep version 1 literals only in migration tests and persistence backward-compatibility tests.

- [ ] **Step 9: Run migration and affected regression tests**

Run:

~~~bash
corepack pnpm exec vitest run src/core/ids.test.ts src/core/workbook/tableNames.test.ts src/core/workbook/migrateWorkbook.test.ts src/lib/persistence.test.ts src/lib/googleSheets.test.ts src/lib/workbook.test.ts src/lib/xlsx.test.ts
~~~

Expected: PASS. Existing workbook behavior remains unchanged and all current constructors return version 2.

- [ ] **Step 10: Type-check and commit**

Run:

~~~bash
corepack pnpm exec tsc -b --pretty false
git add src/core/ids.ts src/core/ids.test.ts src/core/workbook/tableNames.ts src/core/workbook/tableNames.test.ts src/core/workbook/migrateWorkbook.ts src/core/workbook/migrateWorkbook.test.ts src/types.ts src/lib/workbook.ts src/lib/persistence.ts src/lib/googleSheets.ts src/lib/xlsx.ts src/App.test.tsx src/components/Grid.test.tsx src/lib/formulaEngine.perf.test.ts src/lib/persistence.test.ts src/lib/workbook.test.ts src/lib/googleSheets.test.ts
git commit -m "feat(workbook): migrate schema for structured tables"
~~~

Expected: TypeScript exits 0 and the commit contains schema/migration changes only.

### Task 2: Add Structured-Table Metadata, Validation, and Range Operations

**Files:**
- Create: `src/core/workbook/structuredTables.ts`
- Create: `src/core/workbook/structuredTables.test.ts`
- Modify: `src/lib/workbook.ts`

**Interfaces:**
- Consumes: `WorkbookModel`, `WorkbookCommand`, `IdGenerator`, merged-cell and protection queries.
- Produces: table lookup/range helpers, Excel-compatible table-name validation, and `reduceStructuredTableCommand` for non-row operations.

- [ ] **Step 1: Write failing range and creation tests**

Test these exact public functions:

~~~ts
export function getStructuredTable(
  workbook: WorkbookModel,
  tableId: string
): StructuredTable | null;

export function getStructuredTableAtCell(
  workbook: WorkbookModel,
  sheetId: string,
  coord: CellCoord
): StructuredTable | null;

export function getStructuredTableForSelection(
  workbook: WorkbookModel,
  sheetId: string,
  selection: CellRange
): StructuredTable | null;

export function getStructuredTableBodyRange(
  table: StructuredTable
): CellRange | null;
~~~

Create an `A1:C4` table with headers and assert three columns, three row IDs, full range `A1:C4`, body range `A2:C4`, and stable lookup by cell/selection.

- [ ] **Step 2: Write failing validation tests**

Assert atomic rejection for:

- Case-insensitive duplicate table names.
- Invalid Excel table names: blank/whitespace-only; more than 255 Unicode code points; a first character other than a Unicode letter, `_`, or `\`; a later character other than a Unicode letter/number, `_`, or `.`; reserved single-letter `R`/`C`; A1 references such as `A1` and `XFD1048576`; and R1C1 references such as `R1C1`, all checked case-insensitively after NFKC normalization. Include valid Unicode-letter, `_Sales`, and `\Sales` cases. A name that merely looks like a cell outside Excel's `XFD1048576` grid remains valid.
- Blank or duplicate header cells.
- Overlap with another table.
- Any merged cell intersecting the proposed range.
- Protected sheet or locked cell.
- Range outside sheet bounds.
- Zero-width range.
- Resize expansion absorbs adjacent populated cells into the table while still rejecting another table, a merge, or protected cells.

Assert exact issue codes `TABLE_NAME_INVALID`, `TABLE_NAME_CONFLICT`, `TABLE_HEADER_INVALID`, `TABLE_RANGE_OVERLAP`, `TABLE_MERGE_CONFLICT`, `TABLE_PROTECTED`, and `TABLE_RANGE_BLOCKED`. Prove create and rename share the same name validator and that every rejection leaves the workbook byte-for-byte unchanged.

- [ ] **Step 3: Run the metadata tests and verify red**

Run:

~~~bash
corepack pnpm exec vitest run src/core/workbook/structuredTables.test.ts
~~~

Expected: FAIL because structured-table helpers are absent.

- [ ] **Step 4: Implement table lookup and body-range math**

Consume the shared name helpers created in Task 1; migration, commands, XLSX import, and XLSX export all use these same exports:

~~~ts
export type ExcelTableNameValidation =
  | { valid: true; normalizedKey: string }
  | { valid: false; reason: "empty" | "tooLong" | "invalidCharacter" | "reserved" | "cellReference" };

export function validateExcelTableName(name: string): ExcelTableNameValidation;
export function normalizeExcelTableNameKey(name: string): string;
~~~

Validation trims no meaningful characters: leading/trailing whitespace is invalid rather than silently rewritten. Measure the 255-character limit in Unicode code points, normalize with NFKC, and use locale-independent case folding for uniqueness. The first character is a Unicode letter, `_`, or `\`; later characters are Unicode letters/numbers, `_`, or `.`. Reject `R`, `C`, in-grid A1 references, and R1C1 references case-insensitively. The A1 check must parse column/row bounds (`A1` through `XFD1048576`) rather than reject every letter-number suffix. Do not implement this validation with browser locale APIs or duplicate it in `structuredTables.ts`.

Implement body rows with:

~~~ts
const firstBodyRow = table.range.start.row + (table.headerRow ? 1 : 0);
const lastBodyRow = table.range.end.row - (table.totalsRow ? 1 : 0);
~~~

Return `null` when `firstBodyRow > lastBodyRow`. Selection resolution must prefer the table containing `selection.start`; if none contains it, return the first intersecting table in workbook order.

- [ ] **Step 5: Implement creation and deterministic defaults**

Add:

~~~ts
export type StructuredTableCommandServices = {
  createId: IdGenerator;
  getCellEvaluation(sheetId: string, address: string): ComputedCellValue;
};

export type StructuredTableReduction =
  | { status: "committed"; workbook: WorkbookModel }
  | { status: "unchanged"; workbook: WorkbookModel }
  | { status: "rejected"; workbook: WorkbookModel; issues: readonly TableIssue[] };

export function reduceStructuredTableCommand(
  workbook: WorkbookModel,
  command: StructuredTableCommand,
  services: StructuredTableCommandServices
): StructuredTableReduction;
~~~

Creation must:

1. Generate a workbook-unique `Table1`, `Table2` name when omitted.
2. Validate supplied/generated names with `validateExcelTableName` and compare `normalizedKey` workbook-wide.
3. Read nonblank unique header names when `headerRow` is true.
4. Generate `Column1`, `Column2` names when `headerRow` is false.
5. Generate table, column, and body-row IDs through `services.createId`.
6. Use a default style with a light neutral theme and row stripes.
7. Append one immutable table definition to `workbook.tables` only after all validation passes.

- [ ] **Step 6: Implement rename, column rename, style, key column, and resize**

Preserve IDs for every physical row/column that survives. New trailing rows/columns receive fresh IDs. Shrinking drops removed IDs. `table.setKeyColumn` validates that the requested column belongs to the table but does not rewrite current row IDs.

Range resize in this release keeps the same top-left cell. Reject a different top-left with `TABLE_RANGE_BLOCKED` so row/column identity is not guessed. Expanding over ordinary populated cells absorbs those cells as new table rows/columns and assigns fresh logical IDs; overlap, merges, and protection remain blockers.

- [ ] **Step 7: Implement header/totals toggles and conversion**

Header toggle behavior:

- False to true shifts body and totals cells down one row within the footprint and grows the range.
- False to true writes the existing stable `StructuredTableColumn.name` values into the new header cells after shifting; it never derives replacement names from the first body row.
- True to false shifts body and totals cells up one row and shrinks the range.
- True to false preserves column names in metadata while removing the physical header cells, so a later re-enable restores the same names.
- Body row IDs stay unchanged.

Totals toggle behavior:

- False to true adds one totals row after the body.
- True to false clears the totals row, clears each column's standard `totalsFunction` and `totalsLabel`, and shrinks the range.
- A direct text edit in the first totals cell records that column's `totalsLabel`; a direct formula/value edit or a standard aggregate clears `totalsLabel`. `totalsFunction` and `totalsLabel` are mutually exclusive.

Conversion behavior:

- Materialize the current table style into sparse cell formats.
- Preserve cells, formulas, validation, comments, and hyperlinks.
- Remove only the table definition and its filter projection.

- [ ] **Step 8: Integrate table metadata with sheet lifecycle**

Modify workbook sheet operations so:

- Deleting a sheet deletes its tables.
- Renaming/moving a sheet preserves table IDs.
- Duplicating a sheet generates fresh table, column, and row IDs and unique table names.

Pass an optional `IdGenerator` into sheet duplication; the atomic command path supplies the session generator and direct callers use `createRandomId`.

- [ ] **Step 9: Run metadata tests**

Run:

~~~bash
corepack pnpm exec vitest run src/core/workbook/structuredTables.test.ts src/lib/workbook.test.ts
~~~

Expected: PASS for creation, validation, ID preservation, toggle behavior, conversion, and sheet lifecycle.

- [ ] **Step 10: Type-check and commit**

Run:

~~~bash
corepack pnpm exec tsc -b --pretty false
git add src/core/workbook/structuredTables.ts src/core/workbook/structuredTables.test.ts src/lib/workbook.ts
git commit -m "feat(workbook): add structured table metadata"
~~~

Expected: TypeScript exits 0 and the commit contains pure model behavior only.

### Task 3: Implement Table Rows, Calculated Columns, Typed Sorting, and Filtering

**Files:**
- Create: `src/core/workbook/structuredTableRows.ts`
- Create: `src/core/workbook/structuredTableRows.test.ts`
- Create: `src/core/workbook/structuredTableFilter.ts`
- Create: `src/core/workbook/structuredTableFilter.test.ts`
- Modify: `src/core/workbook/structuredTables.ts`
- Modify: `src/core/workbook/commands.ts`
- Modify: `src/core/workbook/commands.test.ts`
- Modify: `src/lib/formulaReferences.ts`
- Modify: `src/lib/formulaReferences.test.ts`
- Modify: `src/lib/formulaEngine.ts`
- Modify: `src/lib/formulaEngine.test.ts`
- Modify: `src/lib/workbook.ts`
- Modify: `src/lib/workbook.test.ts`

**Interfaces:**
- Consumes: `WorkbookSession.getCellEvaluation`, the canonical `src/lib/formulaReferences.ts` token-aware service, the canonical `src/lib/formulaEngine.ts` projection, `FilterExpression`, generic `WorkbookCommand` routing, and table metadata helpers.
- Produces: range-local row insertion/deletion/permutation, rectangular formula rewriting, filtered totals evaluation, stable typed sort, row visibility, and table-aware generic structure commands.

- [ ] **Step 1: Write failing row identity tests**

Cover:

~~~ts
it("inserts before a stable row ID and preserves existing IDs", () => {
  const result = insertStructuredTableRows(workbook, "table-1", {
    beforeRowId: "row-2",
    count: 2
  }, services);
  expect(table(result.workbook).rowIds).toEqual([
    "row-1", "generated-row-1", "generated-row-2", "row-2"
  ]);
});

it("deletes non-contiguous stable row IDs atomically", () => {
  const result = deleteStructuredTableRows(
    workbook,
    "table-1",
    ["row-1", "row-3"],
    services
  );
  expect(table(result.workbook).rowIds).toEqual(["row-2"]);
});
~~~

Add one rejection case for each workbook-wide ID/name, range overlap/bounds, column width/alignment, body-row count, key-column, sort/filter reference, merge, and header/totals invariant listed below; assert malformed storage falls back without partial table recovery.

Also assert that values, formats, validation, comments, and hyperlinks move together and unrelated cells to the right remain unchanged.

- [ ] **Step 2: Write failing calculated-column and formula tests**

Put tokenizer/rewrite unit cases in `src/lib/formulaReferences.test.ts` and workbook row-edit integration/undo cases in `src/core/workbook/structuredTableRows.test.ts`. Do not hide the new shared rewrite contract inside only the table reducer test.

Set the first body formula to `=C2*D2+$A$1`, apply it as a calculated column, insert and sort rows, then assert each destination formula retains row-relative `C/D` references and absolute `$A$1`.

Include formulas containing quoted `"A1"`, `LOG10(A2)`, `1E10`, and `'Rates 2026'!B2` to prove the token-aware service is used.

Add range-local dependency cases outside the table and on another sheet: `SUM(Sheet1!A2:A4)` expands when a table row is inserted; a single-cell reference below the insertion moves only when its column is inside the table footprint; a reference in an unaffected column does not move; deletion shrinks ranges or yields the existing reference-error form when every referenced row is removed. Undo restores every dependent formula byte-for-byte.

Add a table-driven semantics suite for every reference shape the rectangular rewrite supports:

- A same-sheet rectangle wholly outside the edited columns is byte-identical; one wholly inside follows the insertion/deletion row rules; one that crosses the left or right table boundary is split into a parenthesized reference union whose outside slice is unchanged and whose inside slice is rewritten.
- Each member of an existing union is transformed independently in source order and the result is flattened without reordering. A user-authored union member deleted in full becomes `#REF!`; an internally created slice deleted while another slice of that same original rectangle survives is omitted. If no slice survives, the original area becomes `#REF!`.
- Whole-column references remain byte-identical because a rectangular row edit neither changes their column membership nor sheet extent. Whole-row references are expanded against the worksheet's declared row/column bounds into at most three bounded rectangular slices, then processed with the same split rules; this avoids pretending that only part of a whole row moved as a whole-sheet edit.
- Workbook named-range definitions are rewritten once with the same service before dependent formulas publish; formula tokens that use the name remain unchanged. Reject the complete command with `TABLE_FORMULA_REFERENCE_UNSUPPORTED` when an intersecting name is dynamic, external, 3-D, or cannot be represented by the supported union grammar.
- `$` markers are preserved on each dimension, but they do not freeze coordinates during structural edits; mixed references such as `$B2`, `B$2`, and `$B$2` receive the same structural coordinate change as `B2` while retaining their original markers.
- An unqualified reference targets the formula's own sheet. A qualified reference is rewritten only when its normalized target sheet is the edited sheet, even when the formula lives elsewhere; references to other sheets stay byte-identical and retain their original quoting/escaping.

For insertion at row `p` by `n`, an affected row interval entirely above `p` is unchanged, one starting at/after `p` shifts by `+n`, and one spanning `p` expands its end by `n`. For deletion of inclusive rows `p..q`, an interval above is unchanged, one below shifts by `-(q-p+1)`, a partial intersection removes those rows and closes the gap, and an interval with no surviving target becomes `#REF!`. Tests cover boundary insertions, deletion at each edge, nested formulas, union precedence, whole-row/column references, named ranges, mixed absolutes, cross-sheet dependents, and exact undo restoration.

- [ ] **Step 3: Write failing totals-function tests**

Put aggregate/visibility projection and cross-sheet recalculation cases in `src/lib/formulaEngine.test.ts`; keep command-level metadata/formula atomicity cases in `src/core/workbook/structuredTableRows.test.ts`.

Enable the totals row and set every `TableAggregate` variant: `none`, `sum`, `average`, `count`, `countNumbers`, `min`, `max`, `standardDeviation`, and `variance`. Assert metadata and the totals cell update together, the displayed value comes from typed evaluation, row insert/delete/sort refreshes the result, `"none"` clears the formula, and undo restores both formula and metadata. Reject a non-`none` aggregate while totals are disabled. Directly editing a standard totals cell converts it to a custom formula/value and clears `totalsFunction`; direct text in the first totals column records `totalsLabel`.

Build independent row masks and prove a body row contributes only when all three sources say visible: it passes the structured-table `FilterExpression`, it passes every legacy `SheetFilter` whose range covers that row, and it is not explicitly set in `SheetModel.hiddenRows`. Toggle each source alone and in combination; structured and legacy filter masks remain projections and are never copied into `hiddenRows`. Clearing one mask must not override either of the other two. A cross-sheet formula referencing the totals cell must recalculate after every visibility change.

For every aggregate, cover mixed numeric/date/text/boolean/blank/error input, all rows excluded, one visible row, and two or more visible rows. Match Excel range semantics: `sum` returns zero with no numeric inputs; `average` returns `#DIV/0!` with none; `count` counts every nonblank value including errors; `countNumbers` counts numeric/date values only; `min`/`max` return zero with no numeric inputs; sample `standardDeviation`/`variance` require at least two numeric values and otherwise return `#DIV/0!`. `sum`/`average`/`min`/`max`/`standardDeviation`/`variance` ignore text/booleans/blanks, treat dates by their numeric serial, and propagate the first visible typed error in sheet-row order; `count` counts that error and `countNumbers` ignores it. Assert single-value standard deviation/variance errors, empty-mask results, deterministic first-error propagation, and clearing all masks restores the unfiltered typed results.

- [ ] **Step 4: Write failing typed sort tests**

Cover:

- Formula results `2` and `10` sort numerically.
- Dates sort chronologically.
- Booleans sort by boolean value.
- Stable multi-column sorting.
- Errors remain stable after ordinary values.
- Null and empty string sort last.
- Zero and false are not blank.
- `rowIds` receive the exact same permutation as their physical rows.

- [ ] **Step 5: Write failing filter tests**

Use the shared query AST to cover comparison, set, range, blank, logical, and not expressions. Assert `isBlank` means null or empty string only. Header and totals rows remain visible, and clearing a filter changes no cell or row ID.

- [ ] **Step 6: Run the new tests and verify red**

Run:

~~~bash
corepack pnpm exec vitest run src/core/workbook/structuredTableRows.test.ts src/core/workbook/structuredTableFilter.test.ts src/lib/formulaReferences.test.ts src/lib/formulaEngine.test.ts src/core/workbook/commands.test.ts src/lib/workbook.test.ts
~~~

Expected: FAIL because row/filter modules, rectangular reference rewriting, totals projection, and table-aware generic command guards do not exist.

- [ ] **Step 7: Implement range-local cell-plane movement**

Create focused helpers:

~~~ts
export type TableCellPlanes = Pick<
  SheetModel,
  "cells" | "formats" | "validations" | "comments" | "hyperlinks"
>;

export function insertStructuredTableRows(
  workbook: WorkbookModel,
  tableId: string,
  input: { count: number; beforeRowId?: string; afterRowId?: string },
  services: StructuredTableCommandServices
): StructuredTableReduction;

export function deleteStructuredTableRows(
  workbook: WorkbookModel,
  tableId: string,
  rowIds: readonly string[],
  services: StructuredTableCommandServices
): StructuredTableReduction;
~~~

Shift only columns from `table.range.start.column` through `table.range.end.column`. Reject expansion if the new bottom footprint contains unrelated data, a merge, or another table. Clear vacated sparse-map keys rather than writing null entries.

Extend the canonical token-aware reference service in `src/lib/formulaReferences.ts` with:

~~~ts
export type RectangularRowEditContext = {
  formulaSheetId: string;
  editedSheetId: string;
  tableColumnStart: number;
  tableColumnEnd: number;
  row: number;
  count: number;
  operation: "insert" | "delete";
  sheetBounds: { rowCount: number; columnCount: number };
};

export type FormulaRewriteResult =
  | { ok: true; formula: string }
  | { ok: false; issue: TableIssue };

export function rewriteFormulaForRectangularRowEdit(
  formula: string,
  context: RectangularRowEditContext
): FormulaRewriteResult;
~~~

Run it over formulas on every sheet and over workbook named-range definitions in the same atomic reduction, before publishing any workbook. Implement the Step 2 reference-shape table literally: partition partially intersecting rectangles by the table's column boundaries, apply row interval insertion/deletion only to the inside partition, render a parenthesized union when partitions no longer share a rectangle, and preserve `$` plus sheet-qualifier tokens. Whole-column references remain unchanged; whole-row references first expand to bounded left/inside/right rectangles using `sheetBounds`. Existing unions transform member-by-member without sorting or opportunistic coalescing. Reject unsupported intersecting dynamic/external/3-D names or references atomically with `TABLE_FORMULA_REFERENCE_UNSUPPORTED`; never partially rewrite a workbook. Do not model a table-row edit as a whole-sheet row edit and do not use regex replacement.

- [ ] **Step 8: Implement calculated columns and totals formulas**

Add:

~~~ts
export function setStructuredTableCalculatedColumn(
  workbook: WorkbookModel,
  tableId: string,
  columnId: string,
  formula: string | undefined,
  services: StructuredTableCommandServices
): StructuredTableReduction;
~~~

Require formulas to begin with `=`. Store an anchor formula for the first body row and fill every body row through token-aware translation. When the formula is present, body cells in that column are read-only to ordinary edits; clearing it restores editability.

Implement `setStructuredTableTotalsFunction(workbook, tableId, columnId, aggregate, services)`. Map standard aggregates to Excel-compatible totals formulas, using the 100-series `SUBTOTAL` function numbers where available so explicitly hidden rows are excluded, and explicit Excel-compatible sample `VAR`/`STDEV` forms where needed. Store `totalsFunction`, clear any `totalsLabel`, and keep the totals cell/formula atomic with metadata.

Extend the canonical FormulaEngine projection in `src/lib/formulaEngine.ts` with structured-total overrides. After syncing ordinary workbook changes, compute each table's effective visible body-row mask as the intersection of (a) its `FilterExpression`, (b) all applicable legacy `SheetFilter` results using the existing typed filter service, and (c) the inverse of explicit `SheetModel.hiddenRows`. Aggregate only those rows with the exact empty/single/error/type semantics specified in Step 3, then feed the resulting typed literal or typed error into HyperFormula for the physical totals cell while the canonical workbook retains the Excel-compatible formula. Run dependent recalculation after installing overrides, so formulas on any sheet that reference a totals cell observe the filtered result. Recompute overrides after body edits, formula changes, sort, insert/delete, structured-filter or legacy-filter changes, row hide/unhide, undo/redo, and workbook replacement. Track overrides by table/column ID and remove them on totals disable, conversion, table/sheet deletion, or session destroy. Tests in `src/lib/formulaEngine.test.ts` assert filter projections never mutate `SheetModel.hiddenRows`, dependent recalculation observes the override, and the canonical/exported totals formula remains Excel compatible.

- [ ] **Step 9: Implement stable typed row permutation**

Add:

~~~ts
export function sortStructuredTableRows(
  workbook: WorkbookModel,
  tableId: string,
  sorting: readonly TableSort[],
  services: StructuredTableCommandServices
): StructuredTableReduction;
~~~

Build a stable permutation from pre-command `getCellEvaluation` results, apply it once to all cell planes and `rowIds`, regenerate calculated formulas at target rows, and translate other moved relative formulas by destination row minus source row. Do not modify formulas outside the table for a sort.

- [ ] **Step 10: Implement typed table filtering**

Create:

~~~ts
export function matchesStructuredTableFilter(
  expression: FilterExpression,
  getValue: (columnId: string) => QueryScalar
): boolean;

export function isStructuredTableRowVisible(
  workbook: WorkbookModel,
  table: StructuredTable,
  sheetRow: number,
  getEvaluation: (sheetId: string, address: string) => ComputedCellValue
): boolean;
~~~

Use tagged `QueryScalar` comparisons. Never convert values through formatted display strings.

- [ ] **Step 11: Guard generic workbook structure commands**

Update the canonical generic reducer in `src/core/workbook/commands.ts` and add corresponding rejection/routing assertions to `src/core/workbook/commands.test.ts`. Keep low-level range movement assertions in `src/lib/workbook.test.ts`. For row/column insert/delete/sort command handling:

- Operations before a table shift its range and absolute `sheetColumn` values.
- Operations inside a body create/remove the matching logical IDs.
- Header/totals destructive intersections are rejected.
- Partial range sorts intersecting a table are rejected with `TABLE_PARTIAL_STRUCTURAL_EDIT`.
- An exact table sort routes to `sortStructuredTableRows`.

- [ ] **Step 12: Run row, filter, formula, and workbook regressions**

Run:

~~~bash
corepack pnpm exec vitest run src/core/workbook/structuredTableRows.test.ts src/core/workbook/structuredTableFilter.test.ts src/lib/formulaReferences.test.ts src/lib/formulaEngine.test.ts src/core/workbook/commands.test.ts src/lib/workbook.test.ts
~~~

Expected: PASS with row-ID and formula invariants intact.

- [ ] **Step 13: Type-check and commit**

Run:

~~~bash
corepack pnpm exec tsc -b --pretty false
git add src/core/workbook/structuredTableRows.ts src/core/workbook/structuredTableRows.test.ts src/core/workbook/structuredTableFilter.ts src/core/workbook/structuredTableFilter.test.ts src/core/workbook/structuredTables.ts src/core/workbook/commands.ts src/core/workbook/commands.test.ts src/lib/formulaReferences.ts src/lib/formulaReferences.test.ts src/lib/formulaEngine.ts src/lib/formulaEngine.test.ts src/lib/workbook.ts src/lib/workbook.test.ts
git commit -m "feat(workbook): add atomic table row operations"
~~~

Expected: TypeScript exits 0 and the commit contains no React or XLSX changes.

### Task 4: Route Structured Tables Through Atomic Workbook Commands

**Files:**
- Modify: `src/core/workbook/commands.ts`
- Modify: `src/core/workbook/commands.test.ts`
- Modify: `src/core/workbook/WorkbookSession.ts`
- Modify: `src/core/workbook/WorkbookSession.test.ts`
- Modify: `src/core/workbook/history.ts`
- Modify: `src/core/workbook/history.test.ts`

**Interfaces:**
- Consumes: `reduceStructuredTableCommand`, `IdGenerator`, current `ComputedCellValue`.
- Produces: the complete `WorkbookCommand` table union and atomic history behavior.

- [ ] **Step 1: Write failing table command tests**

Add tests that dispatch:

~~~ts
session.dispatch({
  type: "table.create",
  sheetId,
  range: range("A1", "C4"),
  name: "Employees",
  headerRow: true,
  totalsRow: false
});
~~~

Assert one revision increment, one history entry, and exactly one subscriber notification.

Add a rejected batch edit where the second edit violates validation and assert neither edit commits.

- [ ] **Step 2: Write failing undo/redo identity tests**

Create, insert, sort, filter, and delete table rows. Undo each command and assert exact serialized workbook equality, including table, column, and row IDs. Redo and assert the same post-command IDs return.

Add a failing deterministic history-weight test. Starting from two workbooks with the same cell planes, make the second contain a structured table and assert its estimated weight increases for the table definition itself, every `StructuredTableColumn`, every `rowId`, each sort item/filter AST node, every style field, key/aggregate/calculated-formula field, and other persisted table metadata. Change only sort/filter/style metadata and assert the estimator changes without requiring a cell edit. Commit enough weighted table snapshots to cross the 2,000,000 logical-entry limit; assert oldest snapshots trim, retained weight never exceeds the configured maximum, and undo/redo remains exact for every retained snapshot.

- [ ] **Step 3: Run command tests and verify red**

Run:

~~~bash
corepack pnpm exec vitest run src/core/workbook/commands.test.ts src/core/workbook/WorkbookSession.test.ts src/core/workbook/history.test.ts
~~~

Expected: FAIL because the table command union is not routed and structured-table metadata is not included in history weight.

- [ ] **Step 4: Add the complete table command union**

Add the command variants from the Public model and command additions section. Use stable table/column/row IDs for every command after creation. `table.insertRows` rejects zero/negative count and rejects commands with both anchors.

- [ ] **Step 5: Route table reductions with current services**

In `reduceWorkbookCommand`:

1. Detect command types beginning with `table.` through an exhaustive type guard.
2. Supply the session `createId` service.
3. Supply evaluation from the pre-command workbook snapshot.
4. Return rejected issues without replacing history.
5. Commit the returned workbook exactly once.

Extend `estimateWorkbookWeight` in `history.ts` at the same time. The estimator must count table definitions, column definitions, `rowIds`, sort descriptors, every filter AST node/value, style fields, key-column linkage, calculated formulas, totals functions/labels, and all other persisted structured-table metadata. Count logical entries deterministically; do not use JSON byte length, heap sampling, or a constant per table. Preserve the foundation limits of 100 snapshots and 2,000,000 logical entries.

Use issue codes:

~~~ts
type StructuredTableIssueCode =
  | "TABLE_NOT_FOUND"
  | "TABLE_NAME_INVALID"
  | "TABLE_NAME_CONFLICT"
  | "TABLE_HEADER_INVALID"
  | "TABLE_RANGE_OVERLAP"
  | "TABLE_RANGE_BLOCKED"
  | "TABLE_MERGE_CONFLICT"
  | "TABLE_PROTECTED"
  | "TABLE_ROW_NOT_FOUND"
  | "TABLE_COLUMN_NOT_FOUND"
  | "TABLE_CALCULATED_COLUMN_READ_ONLY"
  | "TABLE_TOTALS_ROW_REQUIRED"
  | "TABLE_PARTIAL_STRUCTURAL_EDIT"
  | "TABLE_FORMULA_REFERENCE_UNSUPPORTED";
~~~

- [ ] **Step 6: Make direct workbook cell edits table-aware**

In command reduction:

- Header cell edits route to `table.renameColumn`.
- Calculated body cells reject ordinary `cell.set`.
- Totals cell edits set the cell formula/value and clear the column's standard totals function when it becomes custom. Text in the first totals column records `totalsLabel`; every other custom edit clears any prior label.
- Merge commands intersecting a table reject atomically.

- [ ] **Step 7: Run atomic command tests**

Run:

~~~bash
corepack pnpm exec vitest run src/core/workbook/commands.test.ts src/core/workbook/WorkbookSession.test.ts src/core/workbook/history.test.ts src/core/workbook/structuredTables.test.ts src/core/workbook/structuredTableRows.test.ts
~~~

Expected: PASS with one history entry and publication per committed table command.

- [ ] **Step 8: Type-check and commit**

Run:

~~~bash
corepack pnpm exec tsc -b --pretty false
git add src/core/workbook/commands.ts src/core/workbook/commands.test.ts src/core/workbook/WorkbookSession.ts src/core/workbook/WorkbookSession.test.ts src/core/workbook/history.ts src/core/workbook/history.test.ts
git commit -m "feat(workbook): dispatch table commands atomically"
~~~

Expected: TypeScript exits 0 and the command/session integration is isolated in one commit.

### Task 5: Implement the Live WorkbookTableSession Projection

**Files:**
- Create: `src/table/workbook/WorkbookTableSession.ts`
- Create: `src/table/workbook/WorkbookTableSession.test.ts`
- Create: `src/table/workbook/WorkbookTableSession.contract.test.ts`
- Modify: `src/core/workbook/WorkbookSession.ts`
- Modify: `src/core/workbook/WorkbookSession.test.ts`

**Interfaces:**
- Consumes: `TableSession`, `QueryRow`, `TableViewSnapshot`, `ColumnDef`, `TableIntent`, `ExportArtifact`, `WorkbookSession`, and stable structured-table IDs.
- Produces: `WorkbookTableSession`, `WorkbookTableRow`, `createWorkbookTableSession`, and `WorkbookSession.table(tableId)`.

- [ ] **Step 1: Write failing live-projection tests**

Add:

~~~ts
import { createWorkbookTableSession } from "./WorkbookTableSession";

it("reads cells lazily from the latest workbook snapshot", () => {
  const tableSession = createWorkbookTableSession(workbookSession, "table-1");
  const before = tableSession.getSnapshot();

  workbookSession.dispatch({
    type: "cell.set",
    sheetId: "sheet-1",
    address: "B2",
    input: "updated"
  });

  const after = tableSession.getSnapshot();
  expect(after.revision).not.toBe(before.revision);
  expect(after.getCell("row-1", "column-2").storedValue).toBe("updated");
});
~~~

Also assert that the snapshot has no copied value matrix and that `getCell(rowId, columnId)` is the sole cell accessor.

- [ ] **Step 2: Write failing intent-mapping tests**

The final workbook adapter advertises these exact capabilities:

~~~ts
const workbookTableCapabilities: TableCapabilities = {
  sort: { executor: "client", scope: "completeDataset" },
  filter: { executor: "client", scope: "completeDataset" },
  group: false,
  aggregate: false,
  pagination: false,
  edit: { executor: "client", scope: "completeDataset" },
  bulkEdit: { executor: "client", scope: "completeDataset" },
  metadata: { executor: "client", scope: "completeDataset" },
  validation: { executor: "client", scope: "completeDataset" },
  formula: "fullLocalDataset",
  subscription: true,
  undo: { executor: "client", scope: "completeDataset" },
  export: { executor: "client", scope: "completeDataset" }
};
~~~

During Task 5 keep `export: false`; Task 8 enables it only after native export is implemented and tested. Derive `operationStates` through the shared resolver so grouping, aggregation, and pagination controls are disabled with reasons before click.

Cover these mappings:

| Table intent | Workbook command |
|---|---|
| `edit-cells` | one `table.editCells` batch |
| `clear-cells` | one `table.editCells` batch with blank raw text |
| `set-selection` | worksheet selection command resolved from IDs |
| `set-sorting` | `table.sort` |
| `set-filter` | `table.setFilter` |
| `update-cell-metadata` | one atomic transaction of range format/validation, cell comment, read-only, or formula cell commands |
| `resize-column` | `columns.resize` for the stable table column's physical sheet column |
| `set-column-visibility` | `columns.hidden.set` for that physical sheet column |
| `set-column-order` / `set-column-pinning` | workbook-table-view local state only; worksheet column order is unchanged |
| `insert-rows` | `table.insertRows` |
| `delete-rows` | `table.deleteRows` |
| `undo` | `history.undo` |
| `redo` | `history.redo` |
| `refresh` | current committed snapshot with no source mutation |

Assert unsupported grouping, aggregate, and pagination intents return `rejected/unsupported` rather than processing worksheet rows. Metadata patches validate all stable IDs and build one transaction before dispatch; a mixed valid/invalid batch changes nothing. Formula metadata is translated to a formula `table.editCells` command, while calculated-column definitions continue through `table.setCalculatedColumn`.

- [ ] **Step 3: Write failing identity and lifecycle tests**

Assert:

- Sorting changes each affected `QueryRow` location while retaining its `id`.
- Filtering changes visible `rows` but not `totalRowCount`.
- Undo restores prior row locations and values.
- Two subscribers receive one notification for one workbook commit.
- Destroying a child stops child notifications but leaves the parent alive.
- Destroying the parent destroys all memoized child sessions.
- Converting the table returns an empty snapshot with `TABLE_NOT_FOUND` rather than stale rows.

In `WorkbookTableSession.contract.test.ts`, import `defineTableSessionContract` from `src/table/core/session.contract.ts` and register a workbook-table harness. Supply operations for every advertised capability and explicit unsupported grouping/aggregate/pagination cases. Export is omitted/disabled in the Task 5 harness and added to the same contract in Task 8 when native export is enabled.

- [ ] **Step 4: Run adapter tests and verify red**

Run:

~~~bash
corepack pnpm exec vitest run src/table/workbook/WorkbookTableSession.test.ts src/table/workbook/WorkbookTableSession.contract.test.ts
~~~

Expected: FAIL because the workbook adapter does not exist.

- [ ] **Step 5: Define the workbook locator row and adapter**

Create:

~~~ts
export type WorkbookTableRow = Readonly<{
  id: string;
  tableId: string;
  sheetId: string;
  sheetRow: number;
}>;

export interface WorkbookTableSession
  extends TableSession<WorkbookTableRow> {
  readonly tableId: string;
}

export function createWorkbookTableSession(
  workbookSession: WorkbookSession,
  tableId: string
): WorkbookTableSession;
~~~

`TableViewSnapshot.rows` must contain the shared `QueryRow<WorkbookTableRow>` wrappers. `WorkbookTableRow` is only an identity/location record; do not attach a cell-value object to it.

- [ ] **Step 6: Build snapshots from current table metadata**

For every `getSnapshot()`:

1. Resolve the table by ID from the current parent workbook snapshot.
2. Compute current body sheet rows from header/totals offsets.
3. Pair those sheet rows with `table.rowIds`.
4. Apply the persisted table filter to visible rows.
5. Set `totalRowCount` to `{ kind: "known", value: table.rowIds.length }`, `rowCount` to the filtered projected row count, and `completeness` to `"completeDataset"`.
6. Generate `ColumnDef<WorkbookTableRow>` entries using stable table column IDs.
7. Implement `getCell(rowId, columnId)` by resolving a current address, reading stored/formula/format metadata from `WorkbookSnapshot.workbook`, calling `getCellEvaluation` for the typed computed value, and formatting display text through `formatDisplayValue`.
8. Translate workbook selection to stable table cell IDs when it intersects the table.

Memoize only immutable column definitions by table column identity. Do not memoize worksheet values outside a parent snapshot revision.

- [ ] **Step 7: Implement intent dispatch and result translation**

Dispatch one workbook command per table intent. Preserve the originating command ID when the shared command envelope supplies one. Map workbook validation, permission, unsupported, and conflict-free committed results directly into shared `CommandResult` values.

For `insert-rows`, enforce at most one of `beforeRowId` and `afterRowId`; no anchor appends. For `delete-rows`, reject any row ID absent from the current table rather than deleting a positional substitute.

- [ ] **Step 8: Add WorkbookSession.table**

Extend the existing foundation interface; do not replace it with a table-only declaration or remove diagnostics/replacement behavior:

~~~ts
export interface WorkbookSession {
  getSnapshot(): WorkbookSnapshot;
  getCellEvaluation(sheetId: string, address: string): ComputedCellValue;
  subscribe(listener: () => void): () => void;
  subscribeDiagnostics(listener: (event: WorkbookDiagnosticEvent) => void): () => void;
  dispatch(command: WorkbookCommand | CommandEnvelope<WorkbookCommand>): WorkbookCommandResult;
  replaceWorkbook(workbook: WorkbookModel, options?: { history?: "reset" | "preserve"; origin?: "external" | "import" }): WorkbookCommandResult;
  table(tableId: string): WorkbookTableSession;
  destroy(): void;
}
~~~

Store a private `Map<string, WorkbookTableSession>` inside each workbook session. Create on first request, return the same child on subsequent requests, and destroy the map's values during parent destruction. Import the workbook adapter value from `src/table/workbook/WorkbookTableSession.ts`; the adapter imports `WorkbookSession` as a type-only import to avoid a runtime cycle.

Add a regression assertion that adding/using `table()` does not change `subscribeDiagnostics`, diagnostic unsubscribe/destroy cleanup, or `replaceWorkbook` reset/preserve semantics. A replaced workbook must update existing child table sessions through the normal parent publication path rather than recreating or orphaning them.

- [ ] **Step 9: Run adapter and parent-session tests**

Run:

~~~bash
corepack pnpm exec vitest run src/table/workbook/WorkbookTableSession.test.ts src/table/workbook/WorkbookTableSession.contract.test.ts src/core/workbook/WorkbookSession.test.ts
~~~

Expected: PASS for live values, stable identity, intent routing, ordinary/diagnostic subscriptions, replacement semantics, destruction, and unsupported capabilities.

- [ ] **Step 10: Type-check and commit**

Run:

~~~bash
corepack pnpm exec tsc -b --pretty false
git add src/table/workbook/WorkbookTableSession.ts src/table/workbook/WorkbookTableSession.test.ts src/table/workbook/WorkbookTableSession.contract.test.ts src/core/workbook/WorkbookSession.ts src/core/workbook/WorkbookSession.test.ts
git commit -m "feat(table): project workbook tables as live sessions"
~~~

Expected: TypeScript exits 0 and no React UI is included in this commit.

### Task 6: Add the Contextual Spreadsheet Table Tab

**Files:**
- Create: `src/components/SpreadsheetTableTab.tsx`
- Create: `src/components/SpreadsheetTableTab.test.tsx`
- Create: `src/components/StructuredTableFilterPanel.tsx`
- Create: `src/components/StructuredTableFilterPanel.test.tsx`
- Create: `src/components/CalculatedColumnPanel.tsx`
- Create: `src/components/CalculatedColumnPanel.test.tsx`
- Modify: `src/components/Toolbar.tsx`
- Modify: `src/components/Grid.tsx`
- Modify: `src/App.tsx`
- Modify: `src/App.css`
- Modify: `src/App.test.tsx`

**Interfaces:**
- Consumes: controlled/session-backed `Spreadsheet` in `src/App.tsx`, table commands, table lookup, and shared query AST.
- Produces: Insert Table action, contextual Table ribbon, styled grid projection, typed filter/calculated-column panels, and open/export/convert callbacks.

- [ ] **Step 1: Write failing contextual-tab tests**

Add Testing Library flows:

~~~ts
it("shows the contextual Table tab only while selection intersects a table", async () => {
  const user = userEvent.setup();
  render(<Spreadsheet defaultWorkbook={tableWorkbook} storage={false} />);

  await user.click(screen.getByRole("gridcell", { name: /A2/ }));
  expect(screen.getByRole("tab", { name: "Table" })).toBeInTheDocument();

  await user.click(screen.getByRole("gridcell", { name: /H20/ }));
  expect(screen.queryByRole("tab", { name: "Table" })).not.toBeInTheDocument();
});
~~~

Also test that the active cell's table wins when a large selection intersects two tables.

- [ ] **Step 2: Write failing Insert Table and command-control tests**

Cover:

- Insert ribbon → Table creates from current selection.
- Table tab contains accessible name, range, header, totals-row, per-column totals-function, style, key-column, calculated-column, filter, export, convert, and open-view controls.
- Name/range Enter commits one command.
- Invalid drafts retain focus and announce the issue.
- Conversion preserves values and removes the contextual tab.
- Arrow/Home/End ribbon navigation includes the contextual tab only when present.

- [ ] **Step 3: Write failing panel tests**

`StructuredTableFilterPanel` must build these exact AST categories: comparison, set, range, blank, logical, and not. `CalculatedColumnPanel` must reject text without a leading equals sign and dispatch `table.setCalculatedColumn` for the active stable column ID.

- [ ] **Step 4: Run component tests and verify red**

Run:

~~~bash
corepack pnpm exec vitest run src/components/SpreadsheetTableTab.test.tsx src/components/StructuredTableFilterPanel.test.tsx src/components/CalculatedColumnPanel.test.tsx src/App.test.tsx
~~~

Expected: FAIL because contextual components and Insert Table do not exist.

- [ ] **Step 5: Implement the contextual tab props**

Create:

~~~ts
export type SpreadsheetTableTabProps = {
  table: StructuredTable;
  activeColumnId?: string;
  issues: readonly TableIssue[];
  onRename(name: string): void;
  onResize(range: CellRange): void;
  onHeaderRow(enabled: boolean): void;
  onTotalsRow(enabled: boolean): void;
  onTotalsFunction(columnId: string, aggregate: TableAggregate): void;
  onStyle(style: TableStyle): void;
  onKeyColumn(columnId?: string): void;
  onCalculatedColumn(columnId: string, formula?: string): void;
  onFilter(filter?: FilterExpression): void;
  onExport(): void;
  onConvertToRange(): void;
  onOpenTableView(): void;
};
~~~

Use controlled local drafts for name/range/formula text, reset drafts when `table.id` or committed metadata changes, and expose field-level errors through `aria-describedby` and an assertive issue summary.

Render Export in Task 6 but keep it disabled with the adapter's export-capability reason. Task 8 enables and wires the control after the native exporter passes; do not ship an inert or worksheet-only substitute.

- [ ] **Step 6: Make Toolbar tabs dynamic**

Keep base tabs in a constant and derive:

~~~ts
const tabs = tableContext
  ? [...BASE_RIBBON_TABS, { id: "table" as const, label: "Table" }]
  : BASE_RIBBON_TABS;
~~~

If `activeTab === "table"` and table context disappears, set the active tab to `"home"` and focus Home. Render `SpreadsheetTableTab` only from the Table panel. Add an Insert Table button to the Insert panel.

- [ ] **Step 7: Wire the spreadsheet shell to commands**

In `src/App.tsx`:

1. Resolve `activeTable` through `getStructuredTableForSelection`.
2. Resolve `activeColumnId` from `selection.start.column`.
3. Dispatch all tab mutations through the current `WorkbookSession`.
4. Dispatch totals-function changes as `table.setTotalsFunction` for the active stable column ID and display the current typed total.
5. Create a table with current normalized selection and defaults.
6. Keep the selection inside the created table so the contextual tab appears.
7. Store `openTableId` for Task 7.
8. Report command issues through the existing status surface and the Table tab.

Do not return to direct `setHistory` mutations for these flows.

- [ ] **Step 8: Add table-aware grid context**

Add a grid callback:

~~~ts
getStructuredTableCell?: (
  address: string
) => {
  tableId: string;
  columnId: string;
  rowId?: string;
  role: "header" | "body" | "totals";
  style?: TableStyle;
} | null;
~~~

Use it to apply scoped role classes and table header/filter affordances. Combine structured table visibility with legacy sheet filters and explicit hidden rows. Do not write table filter results into `SheetModel.hiddenRows`.

- [ ] **Step 9: Add scoped neutral styles**

Add only root-scoped selectors such as:

~~~css
.js-spreadsheet-root.js-spreadsheet-workbook .structured-table-cell--header {
  background: var(--spreadsheet-slate-50);
  color: var(--spreadsheet-slate-900);
}

.js-spreadsheet-root.js-spreadsheet-workbook .structured-table-cell--body[data-striped="true"] {
  background: var(--spreadsheet-slate-25);
}

.js-spreadsheet-root.js-spreadsheet-workbook .ribbon-tab--contextual[aria-selected="true"] {
  border-color: var(--spreadsheet-selection-green);
}
~~~

Reuse exported spreadsheet variables. Do not introduce unscoped element selectors.

- [ ] **Step 10: Run component and regression tests**

Run:

~~~bash
corepack pnpm exec vitest run src/components/SpreadsheetTableTab.test.tsx src/components/StructuredTableFilterPanel.test.tsx src/components/CalculatedColumnPanel.test.tsx src/App.test.tsx src/components/Grid.test.tsx
~~~

Expected: PASS for contextual behavior, command routing, grid styling, filtering, and existing ribbon/grid behavior.

- [ ] **Step 11: Type-check and commit**

Run:

~~~bash
corepack pnpm exec tsc -b --pretty false
git add src/components/SpreadsheetTableTab.tsx src/components/SpreadsheetTableTab.test.tsx src/components/StructuredTableFilterPanel.tsx src/components/StructuredTableFilterPanel.test.tsx src/components/CalculatedColumnPanel.tsx src/components/CalculatedColumnPanel.test.tsx src/components/Toolbar.tsx src/components/Grid.tsx src/App.tsx src/App.css src/App.test.tsx
git commit -m "feat(spreadsheet): add contextual table tools"
~~~

Expected: TypeScript exits 0 and the contextual Table UI is complete.

### Task 7: Open Workbook Tables in the Adaptive DataTable View

**Files:**
- Create: `src/react/workbook/WorkbookTableView.tsx`
- Create: `src/react/workbook/WorkbookTableView.test.tsx`
- Modify: `src/App.tsx`
- Modify: `src/App.css`
- Modify: `src/index.ts`

**Interfaces:**
- Consumes: `DataTable` and `WorkbookTableSession`.
- Produces: `WorkbookTableView`, live dual-view behavior, Open in Spreadsheet navigation, and source exports that the packaging plan assigns to the final core/react entrypoints.

- [ ] **Step 1: Write failing live dual-view tests**

Add:

~~~ts
it("reflects edits between DataTable and Spreadsheet without copying rows", async () => {
  const session = createWorkbookSession({ defaultWorkbook: tableWorkbook });
  render(
    <>
      <Spreadsheet session={session} />
      <WorkbookTableView session={session.table("table-1")} onClose={() => {}} />
    </>
  );

  await editTableCell("row-1", "column-2", "Updated");
  expect(screen.getByRole("gridcell", { name: /B2 Updated/ })).toBeInTheDocument();
});
~~~

Also test spreadsheet edits appearing in an already-mounted table view, two simultaneous table views, and isolated separate workbook sessions.

- [ ] **Step 2: Write failing Open in Spreadsheet tests**

Select a table cell by stable IDs, activate `Open in Spreadsheet`, and assert the adaptive view closes, the mapped worksheet cell becomes selected, and the grid scroll API receives its current row/column.

- [ ] **Step 3: Run the React tests and verify red**

Run:

~~~bash
corepack pnpm exec vitest run src/react/workbook/WorkbookTableView.test.tsx
~~~

Expected: FAIL because `WorkbookTableView` does not exist.

- [ ] **Step 4: Implement WorkbookTableView**

Create:

~~~ts
export type WorkbookTableViewProps = {
  session: WorkbookTableSession;
  onClose(): void;
  onOpenInSpreadsheet?(cell?: TableCellRef): void;
};

export function WorkbookTableView(
  props: WorkbookTableViewProps
): React.ReactElement;
~~~

Render a labeled panel/dialog containing `DataTable session={session}`, a compact title/status row, Close, and Open in Spreadsheet. Do not derive a `rows` prop from the workbook.

- [ ] **Step 5: Integrate the view in Spreadsheet**

When `openTableId` is set, obtain `session.table(openTableId)` and mount `WorkbookTableView`. On Open in Spreadsheet:

1. Resolve stable row/column IDs through the latest table metadata.
2. Dispatch worksheet selection.
3. Close the adaptive view.
4. Call the existing grid scroll API with current physical coordinates.
5. Restore grid focus.

If the table was converted/deleted while open, show the adapter's `TABLE_NOT_FOUND` state and keep Close functional.

- [ ] **Step 6: Export the source surface with final package ownership**

Add source-compatibility named exports from `src/index.ts`, but document and test the final package split explicitly:

- `WorkbookTableSession`, `WorkbookTableRow`, and `createWorkbookTableSession` are headless and must be re-exported by `src/entry/core.ts`, producing public imports from `js-spreadsheet/core` only.
- `WorkbookTableView` and `WorkbookTableViewProps` are React-owned and must be re-exported by `src/entry/react.ts`, producing public imports from `js-spreadsheet/react` and the root `js-spreadsheet` React alias.
- Do not expose `WorkbookTableView` from `js-spreadsheet/core`, and do not make the core entry import `src/index.ts`, React, CSS, or browser globals.

The later packaging plan creates the entry files and consumer tests; `src/index.ts` is not described as the final package barrel in this plan.

- [ ] **Step 7: Run React, adapter, and spreadsheet tests**

Run:

~~~bash
corepack pnpm exec vitest run src/react/workbook/WorkbookTableView.test.tsx src/table/workbook/WorkbookTableSession.test.ts src/App.test.tsx
~~~

Expected: PASS for live edits, selection mapping, multiple views, isolation, and missing-table recovery.

- [ ] **Step 8: Type-check and commit**

Run:

~~~bash
corepack pnpm exec tsc -b --pretty false
git add src/react/workbook/WorkbookTableView.tsx src/react/workbook/WorkbookTableView.test.tsx src/App.tsx src/App.css src/index.ts
git commit -m "feat(react): open workbook tables in DataTable"
~~~

Expected: TypeScript exits 0 and the commit contains the adaptive workbook projection only.

### Task 8: Import and Export Native Excel Structured Tables

**Files:**
- Create: `src/lib/xlsxTables.ts`
- Create: `src/lib/xlsxTables.test.ts`
- Create: `src/lib/structuredFormula.ts`
- Create: `src/lib/structuredFormula.test.ts`
- Create: `src/lib/xlsxTableXml.ts`
- Create: `src/lib/xlsxTableXml.test.ts`
- Create: `src/lib/xlsxSecurity.ts`
- Create: `src/lib/xlsxSecurity.test.ts`
- Create: `src/lib/xlsx.real-fixture.test.ts`
- Create: `scripts/create-sales-table-fixture.mjs`
- Create: `src/test/fixtures/xlsx/generated-sales-structured-table.xlsx`
- Create: `src/test/fixtures/xlsx/exceljs-issue-1669.xlsx`
- Create: `src/test/fixtures/xlsx/generated-sales-structured-table.expected.json`
- Create: `src/test/fixtures/xlsx/README.md`
- Modify: `src/lib/xlsx.ts`
- Modify: `src/lib/xlsx.test.ts`
- Modify: `src/table/workbook/WorkbookTableSession.ts`
- Modify: `src/table/workbook/WorkbookTableSession.test.ts`
- Modify: `src/table/workbook/WorkbookTableSession.contract.test.ts`
- Modify: `src/components/SpreadsheetTableTab.test.tsx`
- Modify: `src/App.tsx`
- Modify: `src/App.test.tsx`
- Modify: `package.json`
- Modify: `pnpm-lock.yaml`

**Interfaces:**
- Consumes: untrusted XLSX bytes, ExcelJS worksheets/tables, workbook table metadata, `IdGenerator`, `createStableKeyRowId`, and typed XLSX cell semantics.
- Produces: archive/XML preflight validation, native table import/export, `XlsxImportOptions`, `exportStructuredTableToXlsx`, CSV injection-safe serialization, and preservation of unsupported table XML fields.

- [ ] **Step 1: Add browser-safe ZIP/XML dependencies**

Run:

~~~bash
corepack pnpm add fflate @xmldom/xmldom
~~~

Expected: `package.json` and `pnpm-lock.yaml` include both runtime dependencies.

- [ ] **Step 2: Create the deterministic sales-table fixture**

Create `scripts/create-sales-table-fixture.mjs` with ExcelJS 4.4 and generate `src/test/fixtures/xlsx/generated-sales-structured-table.xlsx`. The script must create `SalesTable` with:

- Worksheet `Sales`.
- Range `A1:F6`.
- Headers `Order ID`, `Order Date`, `Region`, `Units`, `Unit Price`, `Amount`.
- Four rows:
  - `SO-1001`, `2026-01-15`, `East`, `2`, `12.5`, calculated amount `25`.
  - `SO-1002`, `2026-01-16`, `West`, `10`, `3`, calculated amount `30`.
  - `SO-1003`, `2026-02-01`, `North`, `1`, `100`, calculated amount `100`.
  - `SO-1004`, `2026-02-05`, `East`, `4`, `7.5`, calculated amount `30`.
- `Amount` calculated-column formula equal to Units multiplied by Unit Price.
- Totals row summing `Amount`.
- Light muted table style with row stripes.

Set workbook created/modified dates and ZIP entry mtimes to fixed UTC values. ExcelJS cannot author every filter/calculated-column XML variant, so the script then uses `fflate` plus `@xmldom/xmldom` to add the exact active Region filter and calculated-column metadata that a native table contains. This fixture is deterministic generated coverage, not the independently authored compatibility fixture in Step 3. In `README.md` record the generator command, ExcelJS version, fixed fixture date, exact matrix, and SHA-256. Run:

~~~bash
corepack pnpm exec node scripts/create-sales-table-fixture.mjs
corepack pnpm exec node --input-type=module -e "import { readFile } from 'node:fs/promises'; import { createHash } from 'node:crypto'; const bytes = await readFile('src/test/fixtures/xlsx/generated-sales-structured-table.xlsx'); process.stdout.write(createHash('sha256').update(bytes).digest('hex') + '\n');"
~~~

Expected: rerunning the generator produces the recorded hash and the hash command prints 64 lowercase hexadecimal characters.

- [ ] **Step 3: Add a pinned upstream native-filter fixture**

Download ExcelJS's issue-1669 fixture at immutable commit `5bed18b45e824f409b08456b59b87430ded023ab`:

~~~bash
corepack pnpm exec node --input-type=module -e "import { writeFile } from 'node:fs/promises'; const url = 'https://raw.githubusercontent.com/exceljs/exceljs/5bed18b45e824f409b08456b59b87430ded023ab/spec/integration/data/test-issue-1669.xlsx'; const response = await fetch(url); if (!response.ok) throw new Error('fixture download failed: ' + response.status); await writeFile('src/test/fixtures/xlsx/exceljs-issue-1669.xlsx', new Uint8Array(await response.arrayBuffer()));"
~~~

Verify:

~~~bash
corepack pnpm exec node --input-type=module -e "import { readFile } from 'node:fs/promises'; import { createHash } from 'node:crypto'; const bytes = await readFile('src/test/fixtures/xlsx/exceljs-issue-1669.xlsx'); const hash = createHash('sha256').update(bytes).digest('hex'); if (hash !== 'e01c5be103274ae596e9d72ef26b873b87c647269ec1411f1e53ea2b93c25761') throw new Error('unexpected fixture hash: ' + hash);"
~~~

Expected: the verification exits 0. Inspect `docProps/app.xml` and assert `<Application>Microsoft Excel</Application>` plus `AppVersion` are present, making this the independently authored Microsoft Excel compatibility fixture. Record the pinned URL, commit, SHA-256, embedded application metadata, ExcelJS MIT license, and its custom/value-filter purpose in the fixture README.

- [ ] **Step 4: Add exact fixture expectations**

Create `generated-sales-structured-table.expected.json`:

~~~json
{
  "sheetName": "Sales",
  "tableName": "SalesTable",
  "range": "A1:F6",
  "headerRow": true,
  "totalsRow": true,
  "columns": [
    "Order ID",
    "Order Date",
    "Region",
    "Units",
    "Unit Price",
    "Amount"
  ],
  "bodyRowCount": 4,
  "keyColumn": "Order ID",
  "calculatedColumn": "Amount",
  "filteredOutOrderIds": ["SO-1003"]
}
~~~

- [ ] **Step 5: Write failing structured-reference translation tests**

Create `src/lib/structuredFormula.test.ts` before implementation. Cover `[@Units]`, `[@[Unit Price]]`, table-qualified references, headers containing `]`/`#`/apostrophes, `#This Row`, `#Headers`, `#Totals`, `#Data`, `#All`, ranges between table columns, mixed outside-table A1 references, quoted text, and function names containing digits. Assert import produces an A1 formula anchored to the first body row and export reconstructs an equivalent structured formula. Unsupported/ambiguous structured syntax must return a typed issue and block export rather than being silently dropped.

Define:

~~~ts
export function structuredFormulaToA1(
  formula: string,
  table: StructuredTable,
  anchorBodyRow: number
): { ok: true; formula: string } | { ok: false; issue: TableIssue };

export function a1FormulaToStructured(
  formula: string,
  table: StructuredTable,
  anchorBodyRow: number
): { ok: true; formula: string } | { ok: false; issue: TableIssue };
~~~

Use the foundation tokenizer/reference service; do not parse either formula form with a single regular expression.

- [ ] **Step 6: Write failing native-table import tests**

Load `generated-sales-structured-table.xlsx` and assert:

- One `StructuredTable` named `SalesTable` exists.
- Range/header/totals/style/columns match the expectation JSON.
- Four body row IDs exist.
- Expanded worksheet formulas exist in every Amount body cell.
- `calculatedFormula` is anchored to the first body row.
- Order Date retains native date semantics from the typed XLSX prerequisite.
- Imported filter becomes `FilterExpression` and does not permanently populate filter-hidden rows in `SheetModel.hiddenRows`.
- Standard totals functions, a custom totals formula, a literal custom totals value, and `totalsRowLabel` import into the correct physical totals cells. Standard functions populate `totalsFunction`; a label populates `totalsLabel`; custom formula/value cells leave `totalsFunction` unset.

Separately load the pinned Microsoft Excel issue-1669 fixture and assert two tables import: `Table1` and `Table2`, each at `A1:B6`, with no totals row. `Table1` preserves `TableStyleMedium2` and a custom `Column1 != 4` filter. `Table2` preserves `TableStyleLight9`, headers `DK`/`Name`, and the three-value `DK` filter. Import-export-import this real fixture and assert both filters, names, ranges, headers, and styles remain equivalent. Do not apply the SalesTable/calculated/date assertions to this fixture.

- [ ] **Step 7: Write failing identity tests**

Import reordered table bytes twice with:

~~~ts
const options: XlsxImportOptions = {
  tableKeys: {
    SalesTable: { columnName: "Order ID" }
  }
};
~~~

Assert each Order ID receives the same row ID after row reorder. Assert table and column IDs differ between imports. Import without `tableKeys` and assert all row IDs regenerate. Assert blank/duplicate configured keys reject the complete import.

- [ ] **Step 8: Write failing native export tests**

Export an app-created table and reopen bytes with ExcelJS. Assert:

~~~ts
expect(excelWorkbook.getWorksheet("Sales")?.getTables()).toHaveLength(1);
expect(excelWorkbook.getWorksheet("Sales")?.getTable("SalesTable").name).toBe("SalesTable");
~~~

Inspect raw ZIP entries and assert:

- A worksheet table relationship exists.
- `xl/tables/table1.xml` exists.
- Table XML contains `tableStyleInfo`, standard totals metadata, `totalsRowFunction="custom"`, `totalsRowFormula`, `totalsRowLabel`, calculated-column formula, and filter criteria.
- No app table, column, row ID, or `keyColumnId` string appears anywhere in XML.

Round-trip a totals row containing, in separate columns, a standard sum, custom formula, literal custom value, and label. Assert ExcelJS-visible cell values/formulas, raw OOXML attributes/elements, import-export-import equivalence, and JSON-persist/reload-export equivalence. The OOXML supplement must distinguish a label from an arbitrary custom string by `totalsLabel`; it must never infer or overwrite a custom formula/value as a standard function.

In `xlsxSecurity.test.ts`, construct minimal archives and assert typed, all-or-nothing rejection before ExcelJS is invoked for: input over 64 MiB; more than 4,096 ZIP entries; a single declared uncompressed entry over 32 MiB; declared total uncompressed size over 256 MiB; compression ratio over 100:1; encrypted entries; duplicate or traversal entry names; table XML over 4 MiB; `DOCTYPE` or `ENTITY` in any parsed XML; excessive XML depth/elements/attributes; malformed XML; duplicate relationship IDs; external table relationships; a relationship target escaping `xl/`; and missing/wrong-content-type table targets. Include a highly compressed repeated-data ZIP bomb and a forged central-directory-size case without allocating the claimed output. Assert the rejection issue is one of `XLSX_ARCHIVE_LIMIT`, `XLSX_XML_UNSAFE`, or `XLSX_RELATIONSHIP_INVALID`, no partial workbook is returned, and the ExcelJS loader spy has zero calls.

Before implementation, extend `WorkbookTableSession.test.ts`, its shared contract harness, `SpreadsheetTableTab.test.tsx`, and `App.test.tsx` with failing tests for both CSV scopes, injection escaping, native XLSX, disabled pre-wiring state, external-dependency rejection, and the DOM-free artifact boundary. Core/session assertions require `Uint8Array` bytes plus exact media type/file name and prove no `Blob`/DOM global is touched; React assertions require the sole Blob conversion, object-URL cleanup, and download-failure reporting.

- [ ] **Step 9: Run XLSX tests and verify red**

Run:

~~~bash
corepack pnpm exec vitest run src/lib/structuredFormula.test.ts src/lib/xlsxTables.test.ts src/lib/xlsxTableXml.test.ts src/lib/xlsxSecurity.test.ts src/lib/xlsx.real-fixture.test.ts src/table/workbook/WorkbookTableSession.test.ts src/table/workbook/WorkbookTableSession.contract.test.ts src/components/SpreadsheetTableTab.test.tsx src/App.test.tsx
~~~

Expected: FAIL because table adapters, secure XML support, export artifacts, CSV serialization, and React wiring do not exist.

- [ ] **Step 10: Implement formula translation and ExcelJS table conversion**

Create:

~~~ts
export type XlsxImportOptions = {
  idGenerator?: IdGenerator;
  tableKeys?: Readonly<
    Record<string, { columnName: string }>
  >;
};

export async function importStructuredTablesFromWorksheet(
  worksheet: ExcelJS.Worksheet,
  sheetId: string,
  xmlMetadata: readonly NativeTableXmlMetadata[],
  options: XlsxImportOptions
): Promise<StructuredTable[]>;

export function addStructuredTablesToWorksheet(
  workbook: WorkbookModel,
  sheet: SheetModel,
  worksheet: ExcelJS.Worksheet
): void;
~~~

Implement `structuredFormula.ts` first. Import prefers worksheet-expanded A1 body formulas when present and checks them against the table XML formula; otherwise translate XML structured references to the first-body-row A1 anchor. Export translates the internal A1 anchor back to an equivalent structured-reference expression for table XML. Import through a local `listWorksheetTables(worksheet)` helper that runtime-validates and narrows `worksheet.getTables()` to `ExcelJS.Table[]`; ExcelJS 4.4's declaration incorrectly exposes `[Table, void][]` even though runtime returns table objects. Export through `worksheet.addTable()` after ordinary cells/formats have been written, and use `worksheet.getTable(name)` for named assertions. Pass worksheet cell values into table rows so formulas, booleans, numbers, and dates retain ExcelJS native values.

Validate every imported/exported native name with `validateExcelTableName`; reject invalid or normalized case-insensitive duplicate names with a typed issue instead of sanitizing them into a different workbook.

- [ ] **Step 11: Implement XLSX identity rules**

For each imported table:

1. Generate fresh table and column IDs.
2. Match configured table and column names case-insensitively.
3. Reject blank/duplicate configured keys.
4. Derive stable row IDs with `createStableKeyRowId` when configured.
5. Generate fresh random row IDs otherwise.
6. Set `keyColumnId` to the newly generated matching column ID.

Never write these IDs into workbook properties, hidden columns, comments, or XML extension fields.

- [ ] **Step 12: Implement the narrow OOXML supplement**

Create:

~~~ts
export type NativeTableXmlMetadata = {
  name: string;
  calculatedColumns: Readonly<Record<string, string>>;
  filter?: FilterExpression;
  totals: Readonly<Record<string, {
    function?: string;
    formula?: string;
    label?: string;
  }>>;
};

export function readNativeTableXml(
  data: Uint8Array
): readonly NativeTableXmlMetadata[];

export function patchNativeTableXml(
  data: Uint8Array,
  tables: readonly StructuredTable[]
): Uint8Array;
~~~

Use `fflate` to read/write ZIP entries and `@xmldom/xmldom` to parse/serialize XML. Match tables by their Excel name. Read/write only `calculatedColumnFormula` and supported `filterColumn` children that ExcelJS 4.4 drops. Preserve every unrelated ZIP entry and XML node. Do not parse XML with regular expressions.

Also round-trip `totalsRowFunction`, `totalsRowFormula`, and `totalsRowLabel`. A standard function maps to `totalsFunction`; `totalsRowFormula` is translated through `structuredFormula.ts` and stored in the physical totals cell; a custom literal remains the worksheet cell value; and `totalsRowLabel` maps to both the physical text and `totalsLabel`. On export, write mutually exclusive standard function, custom formula/value, or label metadata without changing the physical cell.

Create `xlsxSecurity.ts` and call it before either ExcelJS or `fflate` inflates/parses untrusted input:

~~~ts
export type XlsxSecurityLimits = {
  maxArchiveBytes: number;
  maxEntries: number;
  maxEntryBytes: number;
  maxTotalUncompressedBytes: number;
  maxCompressionRatio: number;
  maxTableXmlBytes: number;
  maxXmlDepth: number;
  maxXmlElements: number;
  maxXmlAttributes: number;
};

export function validateXlsxArchive(
  data: Uint8Array,
  limits?: Partial<XlsxSecurityLimits>
): { ok: true } | { ok: false; issue: TableIssue };
~~~

Defaults are the Step 8 values: 64 MiB input, 4,096 entries, 32 MiB per entry, 256 MiB declared total, 100:1 per-entry ratio, 4 MiB per table XML, XML depth 128, 100,000 elements, and 200,000 attributes. Copy the caller's bytes before validation so later mutation cannot bypass the checked archive. Parse ZIP central-directory metadata without inflating payloads; reject ZIP64/encryption when sizes cannot be proven within limits, invalid/overlapping offsets, duplicate or traversal paths, inconsistent local/central headers, and size arithmetic overflow. Then inflate only validated required entries with bounded output. Before DOM parsing, reject case-insensitive `<!DOCTYPE`/`<!ENTITY`, malformed documents, and configured depth/element/attribute limits. Resolve worksheet relationship targets as normalized package paths, allow only internal table relationships below `xl/tables/`, and verify unique IDs, existing targets, and table content types. Any failure rejects the complete import and skips ExcelJS. `patchNativeTableXml` applies the same limits to its input and output and never returns a partial ZIP.

- [ ] **Step 13: Integrate table XML with workbook import/export**

Import order:

1. Validate ZIP metadata, required content types/relationships, and bounded XML with `validateXlsxArchive`.
2. Read only the validated raw table XML.
3. Load the validated workbook with ExcelJS.
4. Convert worksheets to `SheetModel`.
5. Convert ExcelJS tables and merge raw metadata.
6. Normalize filter-hidden body rows into the table filter projection.
7. Return version 2 `WorkbookModel`.

Export order:

1. Write worksheets/cells/formats.
2. Add native ExcelJS tables.
3. Serialize ExcelJS bytes.
4. Patch unsupported table XML.
5. Return the patched `Uint8Array`.

Add:

~~~ts
export async function exportStructuredTableToXlsx(
  workbook: WorkbookModel,
  tableId: string
): Promise<Uint8Array>;
~~~

This helper creates a one-sheet workbook projection containing the selected table as a native Excel table while preserving typed cells and formulas that are self-contained on that sheet. Before writing, use the formula dependency tokenizer to detect references to other sheets, workbook named ranges, or cells outside the exported projection. Reject with `TABLE_EXPORT_EXTERNAL_DEPENDENCY` and direct callers to full-workbook export; never emit silently broken `#REF!` formulas or copy unrelated workbook data into a table-only export.

Wire `WorkbookTableSession.export(options): Promise<ExportArtifact>` and the contextual Table-tab export now, not in Task 6. Enable the adapter's export capability after wiring. `format: "xlsx", scope: "completeDataset"` calls `exportStructuredTableToXlsx` and returns `{ bytes, mediaType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName: "<valid-table-name>.xlsx" }`, preserving the active table filter as native metadata while retaining all body rows. Reject `xlsx/currentView` with an explicit unsupported issue rather than silently discarding hidden rows or table semantics.

`format: "csv"` supports `currentView` or `completeDataset` using typed display serialization and returns `{ bytes: TextEncoder().encode(csv), mediaType: "text/csv;charset=utf-8", fileName: "<valid-table-name>.csv" }`. Apply an always-on data-only formula-injection policy. Actual formulas export their evaluated display value, never formula source. For a text value whose first character is tab/CR/LF or whose first non-whitespace character is `=`, `+`, `-`, or `@`, prefix one apostrophe before RFC 4180 quoting; an already apostrophe-prefixed value is not double-escaped. Apply this only to text/formula-result strings, not typed negative numbers or dates. There is no unsafe opt-out in this release. Test `=1+1`, ` +cmd`, `-2+3`, `@SUM`, tab/CR/LF prefixes, already-safe apostrophes, commas, quotes, newlines, Unicode whitespace, negative numeric `-42`, both scopes, and byte-exact UTF-8/CRLF output.

Core/session tests assert both formats return a real `Uint8Array`, exact media type, sanitized file name, and stable bytes; they must also assert `Blob`, `document`, `URL.createObjectURL`, and other DOM globals are never read. Only the React `App.tsx` handler converts `ExportArtifact.bytes` to `new Blob([bytes], { type: artifact.mediaType })`, passes `artifact.fileName` to the existing download service, and revokes its object URL on success or failure. `WorkbookTableSession`, `xlsxTables`, and shared contract tests remain runnable in a TypeScript/Node environment without DOM libraries.

Make the Step 8 failing adapter, contract, Table-tab, and App tests green; do not add export tests only after this implementation.

- [ ] **Step 14: Run generated and real-fixture XLSX tests**

Run:

~~~bash
corepack pnpm exec vitest run src/lib/structuredFormula.test.ts src/lib/xlsxTables.test.ts src/lib/xlsxTableXml.test.ts src/lib/xlsxSecurity.test.ts src/lib/xlsx.real-fixture.test.ts src/lib/xlsx.test.ts src/table/workbook/WorkbookTableSession.test.ts src/table/workbook/WorkbookTableSession.contract.test.ts src/components/SpreadsheetTableTab.test.tsx src/App.test.tsx
~~~

Expected: PASS for real Excel import, generated native table export, import-export-import behavior, identity rules, shared formulas, native typed values, raw OOXML/security assertions, adapter contracts, Table-tab wiring, injection-safe CSV, and browser-download cleanup behavior.

- [ ] **Step 15: Type-check and commit**

Run:

~~~bash
corepack pnpm exec tsc -b --pretty false
git add package.json pnpm-lock.yaml scripts/create-sales-table-fixture.mjs src/lib/structuredFormula.ts src/lib/structuredFormula.test.ts src/lib/xlsx.ts src/lib/xlsx.test.ts src/lib/xlsxTables.ts src/lib/xlsxTables.test.ts src/lib/xlsxTableXml.ts src/lib/xlsxTableXml.test.ts src/lib/xlsxSecurity.ts src/lib/xlsxSecurity.test.ts src/lib/xlsx.real-fixture.test.ts src/test/fixtures/xlsx src/table/workbook/WorkbookTableSession.ts src/table/workbook/WorkbookTableSession.test.ts src/table/workbook/WorkbookTableSession.contract.test.ts src/components/SpreadsheetTableTab.test.tsx src/App.tsx src/App.test.tsx
git commit -m "feat(xlsx): round-trip native structured tables"
~~~

Expected: TypeScript exits 0 and the committed fixture provenance identifies real Microsoft Excel.

### Task 9: Pressure-Test Workbook Structured Tables

**Files:**
- Create: `src/core/workbook/structuredTables.property.test.ts`
- Create: `src/core/workbook/structuredTables.perf.test.ts`
- Create: `tests/workbook-tables.spec.ts`
- Verify: `package.json`
- Verify: `pnpm-lock.yaml`
- Modify: `playwright.config.ts` only if Firefox and WebKit projects are not already present

**Interfaces:**
- Consumes: the entire workbook-table stack.
- Produces: randomized invariant evidence, memory/performance evidence, browser coverage, and the subsystem release gate.

- [ ] **Step 1: Verify the foundation's pinned fast-check version**

Run:

~~~bash
corepack pnpm exec node -e "const p=require('./package.json'); if (p.devDependencies?.['fast-check'] !== '4.9.0') throw new Error('fast-check must remain exactly pinned to 4.9.0')"
corepack pnpm why fast-check --depth 0
~~~

Expected: both commands exit 0, `package.json` contains exactly `"fast-check": "4.9.0"`, and the frozen lockfile resolves that direct dependency to 4.9.0. Do not run an unpinned add command or modify the dependency here; Foundation Task 7 owns installation.

- [ ] **Step 2: Write randomized invariant tests**

Generate valid sequences of create, resize, insert, delete, edit, calculated-column, sort, filter, undo, and redo commands. After every committed command assert:

~~~ts
const body = getStructuredTableBodyRange(table);
const bodyRowCount = body ? body.end.row - body.start.row + 1 : 0;

expect(table.rowIds).toHaveLength(bodyRowCount);
expect(new Set(table.rowIds).size).toBe(table.rowIds.length);
expect(new Set(table.columns.map(column => column.id)).size)
  .toBe(table.columns.length);
expect(table.columns).toHaveLength(
  table.range.end.column - table.range.start.column + 1
);
~~~

Record commands that committed, undo them in reverse, and assert exact serialized equality with the original workbook.

Run `fc.assert` with `{ seed: 20260709, numRuns: 500, endOnFailure: true }`. When a property fails, include fast-check's seed/path in the thrown output and document the exact replay command in the test so CI counterexamples reproduce locally.

- [ ] **Step 3: Write formula permutation fuzz tests**

Generate formulas containing relative/absolute cells, ranges, quoted strings, `LOG10`, scientific notation, and sheet-qualified references. Generate row permutations and assert:

- Quoted and token text is byte-identical.
- Relative references move by destination-source offset.
- Absolute dimensions remain absolute.
- Cross-sheet references retain their sheet qualifier.
- Applying a permutation followed by its inverse restores original formulas and row IDs.

- [ ] **Step 4: Write bounded performance tests**

Build a 100,000-row, 100-column table with sparse populated cells and assert:

- `WorkbookTableSession.getSnapshot()` does not materialize a value matrix.
- `getCell` reads only requested cells.
- The snapshot contains exactly one lightweight identity/location `QueryRow` per visible logical body row, with no copied cell values; React DOM bounds are asserted separately in the browser test.
- One row edit creates no duplicate 100,000-row canonical collection.
- Repeated physical sorts create snapshots whose weight includes table definitions, columns, all 100,000 `rowIds`, filters, sorting, styles, calculated/totals/name metadata, and the ordinary cell planes.
- During alternating ascending/descending sorts, weight-based trimming occurs while fewer than 100 snapshots have been attempted. Continue past 100 sorts and assert retained history stays within both the foundation's 100-entry and 2,000,000-logical-entry bounds and still supports exact undo/redo for every retained snapshot.

Use `performance.now()` inside this test with warmup and at least five measured iterations. Assert structural laziness and a generous 5-second safety ceiling for the 100,000-row snapshot/edit case so catastrophic regressions fail without pretending this jsdom/Node test is a production frame benchmark. Record raw timings in the assertion message. The packaging pressure plan later owns pinned-host production baselines and tighter comparative thresholds; do not reference a helper or baseline file that does not exist yet.

- [ ] **Step 5: Run property and performance tests**

Run:

~~~bash
corepack pnpm exec vitest run src/core/workbook/structuredTables.property.test.ts src/core/workbook/structuredTables.perf.test.ts
~~~

Expected: PASS for at least 500 randomized command sequences and the pinned 100,000-row performance case.

- [ ] **Step 6: Write the browser workflow**

In `tests/workbook-tables.spec.ts` cover:

1. Enter headers and typed values.
2. Select range and create a table.
3. Verify contextual Table tab.
4. Open adaptive table view.
5. Edit, insert, sort, and filter.
6. Return to spreadsheet and verify the same committed cells.
7. Undo and verify row identity/value restoration.
8. Upload the real Excel fixture.
9. Verify imported table metadata and calculated values.
10. Export, re-import, and verify native table behavior.
11. Complete the flow by keyboard only and verify focus retention.

- [ ] **Step 7: Ensure all three browser projects exist**

If absent, add Chromium, Firefox, and WebKit projects using the matching Playwright desktop devices. Keep the existing network server binding configuration from the overall browser plan.

- [ ] **Step 8: Run targeted browser tests**

Run:

~~~bash
corepack pnpm exec playwright test tests/workbook-tables.spec.ts --project=chromium
corepack pnpm exec playwright test tests/workbook-tables.spec.ts --project=firefox
corepack pnpm exec playwright test tests/workbook-tables.spec.ts --project=webkit
~~~

Expected: PASS in all three engines with no skipped workbook-table steps.

- [ ] **Step 9: Run the complete subsystem gate**

Run:

~~~bash
corepack pnpm test
corepack pnpm build
corepack pnpm exec playwright test
~~~

Expected: all unit, integration, property, fixture, React, type/build, and browser tests pass. No P0/P1 issue, silent page-local fallback, or unverified native XLSX behavior remains.

- [ ] **Step 10: Commit pressure tests**

Run:

~~~bash
git add playwright.config.ts src/core/workbook/structuredTables.property.test.ts src/core/workbook/structuredTables.perf.test.ts tests/workbook-tables.spec.ts
git commit -m "test(table): pressure test workbook table workflows"
~~~

Expected: the final subsystem commit contains tests/configuration only.

## Row Identity Contract

- App-native JSON persistence preserves table, column, and row IDs exactly.
- Insert creates new row IDs; delete removes only the requested stable IDs.
- Physical sort moves row IDs with the same logical rows.
- Filter never changes `rowIds`.
- Resize retains IDs for surviving physical rows and columns.
- Undo/redo restores IDs from workbook history snapshots.
- Table and column IDs always regenerate on XLSX import.
- Row IDs regenerate on XLSX import unless the host supplies a unique key-column configuration.
- A configured key derives opaque stable IDs across repeated import and row reorder.
- Blank or duplicate configured keys reject the entire import.
- XLSX never stores app IDs or key-column metadata.

## Self-Review Checklist

- [ ] Every structured-table requirement in the approved design maps to Tasks 1–9.
- [ ] No UI structural command precedes the Task 0 correctness gate.
- [ ] All later type and method names match the interfaces declared above.
- [ ] Every new file has one stated responsibility and an exact test file.
- [ ] Every task follows red, green, regression verification, type-check, and commit order.
- [ ] App-native identity and XLSX identity have separate explicit tests.
- [ ] Real Excel and generated ExcelJS files both participate in the XLSX gate.
- [ ] Table views resolve current workbook state and never own canonical row values.
- [ ] Formula/sort tests cover raw, evaluated, display, relative, absolute, quoted, scientific, and cross-sheet cases.
- [ ] Final release remains blocked until the complete pressure gate passes.

## Execution Handoff

Plan complete at `docs/superpowers/plans/2026-07-09-workbook-structured-tables.md`.

Recommended execution: use `superpowers:subagent-driven-development`, dispatch one fresh implementation agent per task, and perform specification-compliance then code-quality review before advancing. If execution is moved to a separate session, use `superpowers:executing-plans` and retain the Task 0 and Task 9 gates unchanged.
