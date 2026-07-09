# DataTable Packaging and Pressure Gate Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Produce stable typed package entrypoints and prevent release until consumer, browser, accessibility, race, fuzz, XLSX, performance, and memory gates pass.

**Architecture:** Keep the demo application build and add a multi-entry library build. Package the DOM-free core separately from React and optional connectors, scope all CSS, verify real consumers, then aggregate every subsystem gate into `pnpm run verify:release`.

**Tech Stack:** TypeScript 6, Vite 8 library mode, React 19 peer dependencies, webpack 5 consumer fixture, Vitest 4, Playwright Chromium/Firefox/WebKit, axe-core, fast-check, pnpm 11.

## Global Constraints

- Requires completed foundation, local, workbook-table, and remote-source plans.
- Keep React and ReactDOM as development and peer dependencies, not bundled runtime dependencies.
- Importing `js-spreadsheet/core` must not access DOM, storage, CSS, Google APIs, or React.
- All public CSS custom properties use the single `--js-spreadsheet-*` namespace.
- Keep GPL and HyperFormula licensing visible in package metadata and embedding docs.
- Do not publish from this plan; build and pack locally for verification.
- Keep `private: true` during this PR so no one can accidentally publish the generic package name; a later release decision must choose the registry/name and deliberately remove that guard.
- Performance assertions run against a pinned Node and production build on a documented host profile.
- Release verification fails on any unresolved P0/P1 issue or skipped mandatory suite.

---

### Task 1: Add real library entrypoints and declaration output

**Files:**
- Create: `src/entry/core.ts`
- Create: `src/entry/react.ts`
- Create: `src/entry/google.ts`
- Create: `src/entry/styles.css`
- Create: `src/styles/tokens.css`
- Create: `src/styles/spreadsheet.css`
- Create: `vite.lib.config.ts`
- Create: `tsconfig.build.json`
- Test: `src/entry/entrypoints.test.ts`
- Modify: `vite.config.ts`
- Modify: `package.json`

**Interfaces:**
- Consumes: public sessions, table types, React components, and Google connector APIs from prior plans.
- Produces: `js-spreadsheet/core`, `js-spreadsheet/react`, `js-spreadsheet/connectors/google`, and `js-spreadsheet/styles.css`.

- [ ] **Step 1: Write a failing entrypoint boundary test**

```ts
import { describe, expect, expectTypeOf, it } from "vitest";
import type {
  WorkbookTableRow as CoreWorkbookTableRow,
  WorkbookTableSession as CoreWorkbookTableSession
} from "./core";
import type {
  WorkbookTableRow as ReactWorkbookTableRow,
  WorkbookTableSession as ReactWorkbookTableSession
} from "./react";

describe("public entrypoints", () => {
  it("keeps the core entrypoint free of React and browser imports", async () => {
    const core = await import("./core");
    expect(core.createWorkbookSession).toBeTypeOf("function");
    expect(core.createLocalRecordTableSession).toBeTypeOf("function");
    expect(core.createRemoteTableSource).toBeTypeOf("function");
    expect(core.createRemoteTableSession).toBeTypeOf("function");
    expect(core.createWorkbookTableSession).toBeTypeOf("function");
    expect(core.createBlankWorkbook).toBeTypeOf("function");
    expect("Spreadsheet" in core).toBe(false);
    expect("DataTable" in core).toBe(false);
  });

  it("publishes the workbook adapter from both core and React/root surfaces", async () => {
    const [core, react] = await Promise.all([import("./core"), import("./react")]);
    expect(core.createWorkbookTableSession).toBeTypeOf("function");
    expect(react.createWorkbookTableSession).toBe(core.createWorkbookTableSession);
    expect(react.WorkbookTableView).toBeTypeOf("function");
    expectTypeOf<ReactWorkbookTableRow>().toEqualTypeOf<CoreWorkbookTableRow>();
    expectTypeOf<ReactWorkbookTableSession>().toEqualTypeOf<CoreWorkbookTableSession>();
  });
});
```

- [ ] **Step 2: Run the test and verify failure**

Run: `corepack pnpm exec vitest run src/entry/entrypoints.test.ts`

Expected: FAIL because the entrypoints do not exist.

- [ ] **Step 3: Add explicit entry files**

```ts
// src/entry/core.ts
export type * from "../types";
export { createBlankWorkbook, getActiveSheet, getCellContent, setCellContent } from "../lib/workbook";
export { createFormulaEngine } from "../lib/formulaEngine";
export { importWorkbookFromXlsx, exportWorkbookToXlsx } from "../lib/xlsx";
export * from "../core/commands/types";
export * from "../core/values/parseCellInput";
export * from "../core/workbook/WorkbookSession";
export * from "../table/core";
export * from "../table/local";
export * from "../table/remote";
export {
  createWorkbookTableSession,
  type WorkbookTableRow,
  type WorkbookTableSession
} from "../table/workbook/WorkbookTableSession";

// src/entry/react.ts
export { Spreadsheet as default, Spreadsheet } from "../react/Spreadsheet";
export { DataTable } from "../react/DataTable";
export { WorkbookTableView } from "../react/workbook/WorkbookTableView";
export { useWorkbookSession } from "../react/useWorkbookSession";
export { WorkbookSessionProvider, useWorkbookSessionContext } from "../react/WorkbookSessionContext";
export { useTableSession } from "../react/useTableSession";
export { useTableSnapshot } from "../react/useTableSnapshot";
export { TableSessionProvider, useTableSessionContext } from "../react/TableSessionContext";
export {
  createWorkbookTableSession,
  type WorkbookTableRow,
  type WorkbookTableSession
} from "../table/workbook/WorkbookTableSession";
export type { SpreadsheetProps } from "../react/Spreadsheet";
export type { WorkbookTableViewProps } from "../react/workbook/WorkbookTableView";
export type {
  CellEditorProps,
  CellRenderContext,
  ColumnDef,
  DataTableHandle,
  DataTablePresentationProps,
  DataTableProps,
  HeaderRenderContext
} from "../react/tableTypes";

// src/entry/google.ts
export * from "../lib/googleAuth";
export * from "../lib/googleSheets";
```

The package root resolves to `src/entry/react.ts`, so that entry is also the root public surface. `WorkbookTableSession`, `WorkbookTableRow`, and `WorkbookTableViewProps` must therefore be available from `js-spreadsheet`, from their appropriate `js-spreadsheet/core` or `js-spreadsheet/react` subpath, and from the generated declarations. Export the headless adapter factory/types from core and React/root as shown; export the React view only from React/root. Do not re-export `WorkbookTableView`, React types, CSS, or browser helpers from core. The existing `WorkbookSession.table(tableId)` method remains available through the core `WorkbookSession` export.

- [ ] **Step 4: Create buildable style entry dependencies, then add the library build and exports map**

Configure Vite library mode with the three TypeScript entries, exact file names `core.js`, `react.js`, and `connectors/google.js`, and CSS file `styles.css`. Externalize `react`, `react-dom`, `exceljs`, and `hyperformula`; keep ExcelJS and HyperFormula in `dependencies`, and move React/ReactDOM to both `peerDependencies` and `devDependencies`. Generate declarations into `dist/types` through `tsc -p tsconfig.build.json`. Put library assets in `dist/lib` and the existing SPA in `dist/app`; each build empties only its own directory so neither can erase the other.

Before creating `src/entry/styles.css`, create `src/styles/tokens.css` and `src/styles/spreadsheet.css`; `src/styles/data-table.css` already exists from the local-table plan. `tokens.css` establishes on the two approved component roots the exact `--js-spreadsheet-*` names and defaults enumerated in Task 2 Step 3. `spreadsheet.css` starts with the root-qualified workbook container rules needed by the embeddable surface; Task 2 moves the remaining legacy rules out of `App.css`. This ordering is mandatory: Task 1 must build green on its own and may not import a file deferred to Task 2.

`src/entry/styles.css` then imports `../styles/tokens.css`, `../styles/spreadsheet.css`, and the existing `../styles/data-table.css` in that order. Add it as an explicit Rollup/Vite input and force its emitted asset name to `styles.css`; do not make the DOM-free core import CSS. Set `vite.config.ts` app `build.outDir` to `dist/app`, and set the library `outDir` to `dist/lib` with declaration output remaining in `dist/types`.

The package export shape must be:

```json
{
  "main": "./dist/lib/react.js",
  "module": "./dist/lib/react.js",
  "types": "./dist/types/entry/react.d.ts",
  "exports": {
    ".": { "types": "./dist/types/entry/react.d.ts", "import": "./dist/lib/react.js" },
    "./core": { "types": "./dist/types/entry/core.d.ts", "import": "./dist/lib/core.js" },
    "./react": { "types": "./dist/types/entry/react.d.ts", "import": "./dist/lib/react.js" },
    "./connectors/google": { "types": "./dist/types/entry/google.d.ts", "import": "./dist/lib/connectors/google.js" },
    "./styles.css": "./dist/lib/styles.css"
  },
  "files": ["dist/lib", "dist/types", "README.md", "LICENSE"]
}
```

Put `// @vitest-environment node` at the top of `entrypoints.test.ts`. After building, run a fresh Node process that imports `dist/lib/core.js` with no DOM or storage shims; this is the authoritative browser-global boundary check.

- [ ] **Step 5: Run entrypoint test and both builds**

Run:

```bash
corepack pnpm exec vitest run src/entry/entrypoints.test.ts
corepack pnpm run build:lib
corepack pnpm run build:app
node --input-type=module -e "const core = await import('./dist/lib/core.js'); if (typeof core.createWorkbookSession !== 'function' || typeof core.createWorkbookTableSession !== 'function' || 'WorkbookTableView' in core) process.exit(1)"
```

Expected: PASS; `dist` contains ESM, declarations, and CSS plus the app build in its configured directory.

- [ ] **Step 6: Commit library entrypoints**

```bash
git add src/entry src/styles/tokens.css src/styles/spreadsheet.css vite.lib.config.ts vite.config.ts tsconfig.build.json package.json pnpm-lock.yaml
git commit -m "build: add typed library entrypoints"
```

### Task 2: Scope styles and prove instance isolation

**Files:**
- Modify: `src/styles/tokens.css`
- Modify: `src/styles/spreadsheet.css`
- Modify: `src/styles/data-table.css`
- Create: `src/styles/styleIsolation.test.tsx`
- Modify: `src/App.css`
- Modify: `src/react/Spreadsheet.tsx`
- Modify: `src/react/DataTable.tsx`
- Modify: `package.json`
- Modify: `pnpm-lock.yaml`

**Interfaces:**
- Consumes: approved component roots `.js-spreadsheet-root.js-spreadsheet-workbook` and `.js-spreadsheet-root.js-spreadsheet-data-table`.
- Produces: scoped theme variables and host-safe component CSS.

- [ ] **Step 1: Write the host leakage test**

```tsx
type Person = { id: string; name: string };
const rows: readonly Person[] = [{ id: "1", name: "Ada" }];
const columns: readonly ColumnDef<Person>[] = [{
  id: "name",
  header: "Name",
  accessor: (row) => row.name,
  update: (row, value) => ({ ...row, name: String(value) })
}];
const props = { rows, columns, getRowId: (row: Person) => row.id } as const;

it("does not restyle host controls outside component roots", () => {
  render(<><button data-testid="host">Host</button><DataTable rows={rows} columns={columns} getRowId={(row) => row.id} /></>);
  const host = screen.getByTestId("host");
  expect(host.className).toBe("");
  expect(document.querySelector(".js-spreadsheet-root.js-spreadsheet-data-table")).toBeInTheDocument();
});

it("keeps two table theme overrides isolated", () => {
  render(<><DataTable className="theme-a" {...props} /><DataTable className="theme-b" {...props} /></>);
  expect(document.querySelector(".theme-a")).not.toBe(document.querySelector(".theme-b"));
});
```

Add `postcss` as a direct dev dependency. Parse all three CSS files in the same test and fail any ordinary selector that is not rooted beneath `.js-spreadsheet-root`; allow only keyframe selectors and root-qualified `@media`/`@container` contents. Fail any public custom-property declaration or `var(...)` reference outside the `--js-spreadsheet-*` namespace, including every legacy short-prefix spelling. Also compare `getComputedStyle` for host `button`, `input`, and `select` controls before and after mounting both components.

- [ ] **Step 2: Run and verify failure**

Run: `corepack pnpm exec vitest run src/styles/styleIsolation.test.tsx`

Expected: FAIL while global selectors and root assumptions remain.

- [ ] **Step 3: Move tokens and scope selectors**

Define the shared defaults `--js-spreadsheet-accent: #22795d`, `--js-spreadsheet-accent-strong: #17634a`, `--js-spreadsheet-accent-soft: #eaf7f2`, `--js-spreadsheet-border: #cbd5e1`, `--js-spreadsheet-border-strong: #94a3b8`, `--js-spreadsheet-chrome: #f8fafc`, `--js-spreadsheet-chrome-strong: #eef3f7`, `--js-spreadsheet-text: #172033`, `--js-spreadsheet-muted: #5f6e82`, and `--js-spreadsheet-surface: #ffffff` on each component root. These are the same names already established by the local table plan; migrate every workbook/table declaration and reference to them and introduce no compatibility alias under another prefix. Replace global `body`, `button`, `input`, `select`, and universal selectors with root-qualified selectors. Replace `100vh` assumptions with container sizing for embedded components while preserving standalone App layout through an app-only wrapper. Keep `prefers-reduced-motion`, forced-colors-safe borders, visible focus, and container-query rules from the local table plan.

- [ ] **Step 4: Run CSS isolation and existing UI tests**

Run:

```bash
corepack pnpm exec vitest run src/styles/styleIsolation.test.tsx src/App.test.tsx src/react/DataTable.test.tsx src/react/workbook/WorkbookTableView.test.tsx
corepack pnpm run build:lib
corepack pnpm run build:app
```

Expected: tests and both builds PASS; the emitted `dist/lib/styles.css` contains only `--js-spreadsheet-*` public variables and root-scoped selectors.

- [ ] **Step 5: Commit scoped styles**

```bash
git add src/styles/tokens.css src/styles/spreadsheet.css src/styles/data-table.css src/styles/styleIsolation.test.tsx src/App.css src/react/Spreadsheet.tsx src/react/DataTable.tsx package.json pnpm-lock.yaml
git commit -m "style: scope spreadsheet and data table themes"
```

### Task 3: Test packed consumers in Vite and webpack

**Files:**
- Create: `scripts/test-consumers.mjs`
- Create: `tests/consumers/vite/index.tsx`
- Create: `tests/consumers/vite/index.html`
- Create: `tests/consumers/vite/vite.config.ts`
- Create: `tests/consumers/vite/package.json`
- Create: `tests/consumers/vite/tsconfig.json`
- Create: `tests/consumers/core/index.ts`
- Create: `tests/consumers/core/package.json`
- Create: `tests/consumers/core/tsconfig.json`
- Create: `tests/consumers/webpack/index.tsx`
- Create: `tests/consumers/webpack/index.html`
- Create: `tests/consumers/webpack/webpack.config.cjs`
- Create: `tests/consumers/webpack/package.json`
- Create: `tests/consumers/webpack/tsconfig.json`
- Modify: `package.json`

**Interfaces:**
- Consumes: a freshly rebuilt packed tarball produced by `pnpm pack`.
- Produces: `test:consumers` with real package-name/subpath resolution, browser runtime smoke coverage, and a DOM-free core declaration consumer.

- [ ] **Step 1: Add consumer source files**

Both React consumers must import every public surface through its package name (never a source deep import):

```tsx
import {
  createBlankWorkbook,
  createLocalRecordTableSession,
  createRemoteTableSession,
  createRemoteTableSource,
  createWorkbookSession
} from "js-spreadsheet/core";
import {
  DataTable,
  Spreadsheet,
  TableSessionProvider,
  WorkbookSessionProvider,
  useTableSnapshot,
  type ColumnDef,
  type DataTableHandle,
  type WorkbookTableRow as ReactWorkbookTableRow,
  type WorkbookTableSession as ReactWorkbookTableSession
} from "js-spreadsheet/react";
import {
  DataTable as RootDataTable,
  createWorkbookTableSession as createRootWorkbookTableSession,
  type WorkbookTableRow as RootWorkbookTableRow,
  type WorkbookTableSession as RootWorkbookTableSession
} from "js-spreadsheet";
import { parseSpreadsheetId } from "js-spreadsheet/connectors/google";
import "js-spreadsheet/styles.css";
import { createRoot } from "react-dom/client";
import { useEffect, useRef } from "react";

const rows = [{ id: "1", name: "Ada" }];
const columns = [{
  id: "name",
  header: "Name",
  accessor: (row: { id: string; name: string }) => row.name,
  update: (row: { id: string; name: string }, value: unknown) => ({ ...row, name: String(value) }),
  cell: ({ row, cell }) => <strong data-renderer>{row.original.name}:{cell.displayValue}</strong>,
  editor: ({ rawText, onChange, onCommit }) => <input value={rawText} onChange={(event) => onChange(event.currentTarget.value)} onBlur={onCommit} />
}] satisfies readonly ColumnDef<(typeof rows)[number]>[];

parseSpreadsheetId("https://docs.google.com/spreadsheets/d/example-id/edit");
void createRootWorkbookTableSession;
type WorkbookAdapterSurface =
  | ReactWorkbookTableRow | RootWorkbookTableRow
  | ReactWorkbookTableSession | RootWorkbookTableSession;
void (null as WorkbookAdapterSurface | null);
```

Complete each fixture with a local session, inert remote source/session, and workbook session created from the imported core factories. The remote query returns one deterministic `QueryRow` and a known complete total, then is started/stopped in a committed effect. Render all of the following at runtime: `WorkbookSessionProvider` around `Spreadsheet`, `TableSessionProvider` around a child that calls `useTableSnapshot`, a ref-backed `DataTable` using `useRef<DataTableHandle>(null)`, a session-backed remote `DataTable`, and `RootDataTable`. Keep the renderer/editor callbacks inline and unannotated as above so declaration inference—not local source types—supplies their contexts. After the snapshot is readable and both tables mount, set `data-consumer-ready="true"` on the root and render the snapshot cell value for the smoke assertion.

The dedicated `tests/consumers/core/index.ts` imports `ExportArtifact`, `RemoteTableSource`, `TableAbortSignal`, `TableSession`, `WorkbookTableRow`, `WorkbookTableSession`, and `createWorkbookTableSession` plus representative factories from `js-spreadsheet/core`. It assigns `Awaited<ReturnType<TableSession<Row>["export"]>>` to `ExportArtifact`, verifies the workbook adapter factory/type relationship, reads the structural abort signal in a remote callback type, and uses `Uint8Array`/`ArrayBuffer` only. Its `tsconfig.json` must set `"lib": ["ES2022"]`, `"types": []`, strict mode, and `noEmit`; it must not declare DOM shims or depend on React. This is the authoritative declaration-boundary test for accidental `Blob`, `File`, `AbortSignal`, `HTMLElement`, CSS, or React leakage.

- [ ] **Step 2: Write the pack-and-build script**

The script must first remove `dist/lib` and `dist/types`, run `corepack pnpm run build:lib`, and assert both output trees were recreated; it may only then create a temporary directory and run `corepack pnpm pack --pack-destination <temp>/package`. This prevents stale declarations or CSS from making the packed-consumer gate pass. Inspect the tarball manifest and assert every `main`, `module`, `types`, `files`, and recursive `exports` target exists in the archive. Copy each fixture into the temporary directory and replace its `js-spreadsheet` dependency with the exact tarball path. For every fixture, run `corepack pnpm install --lockfile-only --no-frozen-lockfile`, `corepack pnpm install --frozen-lockfile`, and `corepack pnpm exec tsc --noEmit`; production-build both React fixtures. Inspect `pnpm list react --json --depth Infinity` in the React fixtures and fail if the packed package installs a second React runtime.

Serve each production React build from an ephemeral loopback port with a minimal Node static server, open it with the root Playwright Chromium installation, and require `data-consumer-ready="true"`, the expected snapshot value, and zero console errors, page errors, or failed module/CSS requests. Always close pages, browsers, servers, and remove the complete temporary tree in `finally`. The script exits nonzero on stale/missing build output, any missing manifest target/export/declaration/CSS asset/peer dependency, a type error, duplicate React, bundle failure, or runtime-smoke failure.

Every committed fixture `package.json` contains exactly `"packageManager": "pnpm@11.7.0"` and pins all dependency versions exactly. The Vite fixture depends on Vite and `@vitejs/plugin-react`; the webpack fixture depends on webpack, webpack-cli, ts-loader, css-loader, style-loader, and its HTML build plugin. The core fixture depends only on the tarball and pinned TypeScript. The root workspace remains `packages: ["."]`; do not add fixture projects to the product lockfile.

- [ ] **Step 3: Run and verify initial consumer failure**

Run: `corepack pnpm run test:consumers`

Expected before fixes: FAIL on at least one missing or incorrect packed export.

- [ ] **Step 4: Fix package paths until both consumers pass**

Do not deep-import source files in fixtures. Resolve only through the package export map.

- [ ] **Step 5: Commit consumer verification**

```bash
git add scripts/test-consumers.mjs tests/consumers package.json pnpm-lock.yaml
git commit -m "test: verify packed React consumers"
```

### Task 4: Expand browser and accessibility coverage

**Files:**
- Create: `tests/data-table-accessibility.spec.ts`
- Modify: `playwright.config.ts`
- Modify: `package.json`

**Interfaces:**
- Consumes: local, remote, and workbook demos from prior plans.
- Produces: Chromium, Firefox, and WebKit release-browser coverage.

- [ ] **Step 1: Add all three browser projects**

```ts
projects: [
  { name: "chromium", use: { ...devices["Desktop Chrome"] } },
  { name: "firefox", use: { ...devices["Desktop Firefox"] } },
  { name: "webkit", use: { ...devices["Desktop Safari"] } }
]
```

Keep automation on `baseURL: "http://127.0.0.1:5173"`, but set `webServer.command` to `corepack pnpm run dev -- --port 5173` so Vite binds to `0.0.0.0` through the package script. Manual verification output must discover the machine address (`hostname -I`) and print `http://<machine-ip>:5173`; never hand the user a localhost link.

- [ ] **Step 2: Install the release browser binaries**

Run:

```bash
corepack pnpm exec playwright install chromium firefox webkit
```

Expected: all three browser executables are present in the Playwright cache. In a CI image missing OS libraries, document and run `playwright install --with-deps chromium firefox webkit` only in an authorized provisioning step; do not hide a missing browser by skipping its project.

- [ ] **Step 3: Add the accessibility workflow**

Test keyboard-only header navigation, grid selection, editing, validation, sorting, filtering, conflict resolution, focus after virtualization, and status announcements. Run axe against the local, remote, and workbook views with zero violations.

- [ ] **Step 4: Run the focused cross-browser suite**

Run: `corepack pnpm exec playwright test tests/data-table-accessibility.spec.ts`

Expected: PASS in Chromium, Firefox, and WebKit.

- [ ] **Step 5: Run all browser tests in all projects**

Run: `corepack pnpm run test:e2e:all`

Expected: PASS without relying on a pre-existing local server.

- [ ] **Step 6: Commit browser coverage**

```bash
git add tests/data-table-accessibility.spec.ts playwright.config.ts package.json pnpm-lock.yaml
git commit -m "test: cover data tables across browsers"
```

### Task 5: Enforce performance and memory budgets

**Files:**
- Create: `src/table/performance/localTable.perf.test.ts`
- Create: `src/table/performance/remoteTable.perf.test.ts`
- Create: `src/table/performance/historyMemory.perf.test.ts`
- Create: `tests/performance/data-table.perf.spec.ts`
- Create: `tests/performance/accepted-baseline.json`
- Create: `scripts/run-performance-gate.mjs`
- Modify: `docs/performance.md`
- Modify: `package.json`

**Interfaces:**
- Consumes: production DataTable build and all session types.
- Produces: reproducible `test:performance` with recorded environment and thresholds.

- [ ] **Step 1: Add local and remote scale fixtures**

Generate 100,000 local rows with at least 1,000 columns and a remote source reporting 1,000,000 total rows while retaining at most five 200-row pages. Assert row and column DOM nodes remain bounded to viewport plus overscan and scroll to the last row/column without allocating the full cell matrix.

- [ ] **Step 2: Add operation and memory assertions**

On the pinned benchmark environment, require:

- Local edit-to-paint p95 below 50 ms in the production React build.
- Scroll/render frame p95 below 16.7 ms after warmup on the pinned release host.
- Remote query swap processing below 50 ms, excluding network delay.
- No page-cache growth beyond the declared five-page limit.
- No optimistic overlay after terminal mutation results.
- Bounded workbook history, local inverse history, remote page cache, and remote journal after 1,000 operations.
- After warmup and forced GC, 50 mount/edit/scroll/unmount cycles retain no listeners, observers, timers, sessions, overlays, or pages and grow heap by no more than the accepted 20 MB tolerance.

Run timing in `tests/performance/data-table.perf.spec.ts` against `corepack pnpm run preview -- --port 4173`, which binds `0.0.0.0`; Playwright may connect through loopback. Run structural and heap checks in a child Node process started with `--expose-gc`. Store raw results as JSON in `test-results/performance/` and fail on an absolute threshold or greater than 20% regression from the checked-in accepted baseline metadata. On a host whose CPU/OS differs from baseline, absolute correctness, structural bounds, cleanup, and generous safety ceilings still fail the gate; only the tighter comparative regression assertion is marked non-comparable, never the suite itself.

- [ ] **Step 3: Run the gate on a production build**

Run: `corepack pnpm run test:performance`

Expected: PASS and print local, remote, DOM, and memory metrics.

- [ ] **Step 4: Document the benchmark host and interpretation**

Update `docs/performance.md` with Node version, CPU, browser, dataset shapes, warmup, iterations, thresholds, and the rule that development React numbers are diagnostic rather than release budgets.

- [ ] **Step 5: Commit performance gates**

```bash
git add src/table/performance tests/performance/data-table.perf.spec.ts tests/performance/accepted-baseline.json scripts/run-performance-gate.mjs docs/performance.md package.json
git commit -m "test: enforce data table performance budgets"
```

### Task 6: Assemble the release verification command and documentation

**Files:**
- Create: `scripts/verify-release.mjs`
- Create: `scripts/check-release-review.mjs`
- Modify: `package.json`
- Modify: `README.md`
- Modify: `docs/embedding.md`
- Modify: `docs/architecture.md`
- Modify: `docs/features.md`
- Modify: `docs/performance.md`

**Interfaces:**
- Consumes: every plan's verification command.
- Produces: `verify:release` and complete user/developer documentation.

- [ ] **Step 1: Implement fail-fast release orchestration**

Run, in order: Node/pnpm version check, `git diff --check`, TypeScript/declaration tests, unit tests, fixed-seed property/race tests, app build, library build, fresh-Node core import, packed consumers, real/generated XLSX fixtures, Chromium/Firefox/WebKit E2E, axe/keyboard accessibility, and production performance/memory gates. Print elapsed time and one summary line per stage. Propagate the first nonzero exit code and delete temporary servers/tarballs in `finally`.

- [ ] **Step 2: Add package scripts**

```json
{
  "scripts": {
    "build:app": "tsc -b && vite build",
    "build:lib": "tsc -p tsconfig.build.json && vite build -c vite.lib.config.ts",
    "test:consumers": "node scripts/test-consumers.mjs",
    "test:e2e:all": "playwright test",
    "test:performance": "node scripts/run-performance-gate.mjs",
    "verify:release": "node scripts/verify-release.mjs",
    "verify:release:reviewed": "node scripts/check-release-review.mjs && node scripts/verify-release.mjs"
  }
}
```

- [ ] **Step 3: Update public documentation**

Document local, remote, and workbook examples; source authority; capability scope; conflicts and undo; structured-table identity across XLSX; Table-tab workflows; styles; accessibility; performance; package exports; React-only scope; and GPL/HyperFormula licensing.

- [ ] **Step 4: Run the complete release command from a clean checkout**

Run: `corepack pnpm run verify:release`

Expected: exit 0 with no skipped mandatory suite and no uncommitted generated package artifacts. Before independent review use `verify:release`; after Task 7 creates the review record, the PR gate is `verify:release:reviewed`.

- [ ] **Step 5: Commit release verification and docs**

```bash
git add scripts/verify-release.mjs scripts/check-release-review.mjs package.json README.md docs/embedding.md docs/architecture.md docs/features.md docs/performance.md
git commit -m "docs: finalize data table release workflow"
```

### Task 7: Perform the final branch audit

**Files:**
- Review: all files changed from the feature-branch merge base
- Create: `docs/reviews/2026-07-09-data-table-release.md`

**Interfaces:**
- Consumes: complete implementation and verification output.
- Produces: a review-ready branch for the program plan's PR checkpoint.

- [ ] **Step 1: Run the release gate again without cached test results**

Run: `corepack pnpm run verify:release`

Expected: PASS.

- [ ] **Step 2: Review the complete diff**

Run:

```bash
git diff --check origin/main...HEAD
git diff --stat origin/main...HEAD
git status --short
```

Expected: no whitespace errors, no unexpected generated files, and a clean worktree.

- [ ] **Step 3: Request independent code review**

Invoke `superpowers:requesting-code-review` with the design spec, plan set, merge base, HEAD, release-gate output, and measured performance summary. Also obtain independent final implementation reviews from the available Claude and Auggie CLIs because the user explicitly requested both perspectives. Record each finding, reviewer, severity, disposition, evidence, and fixing commit in `docs/reviews/2026-07-09-data-table-release.md`.

- [ ] **Step 4: Resolve every confirmed P0/P1 finding**

For each accepted finding, add a failing regression test, implement the smallest correction, rerun its focused suite, and commit with a scoped `fix:` message. Reject incorrect findings with current-code evidence.

- [ ] **Step 5: Rerun the full gate after review changes**

Run: `corepack pnpm run verify:release:reviewed`

Expected: PASS and clean worktree. `check-release-review.mjs` fails if the review record is missing, malformed, or contains an open P0/P1. The branch is now eligible for the new pull request described by the master program.

- [ ] **Step 6: Commit the final review record**

```bash
git add docs/reviews/2026-07-09-data-table-release.md
git commit -m "docs: record data table release review"
corepack pnpm run verify:release:reviewed
git status --short
```

Expected: reviewed gate passes and the worktree is clean.
