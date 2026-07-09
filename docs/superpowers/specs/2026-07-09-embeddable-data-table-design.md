# Embeddable React DataTable and Structured Workbook Tables Design

Date: 2026-07-09

Status: Ready for implementation planning after user review

## Decision Summary

Keep the existing spreadsheet and add a React-first `DataTable` surface. The
spreadsheet and table reuse a narrow interaction and command contract, while
their authoritative data sources retain different write, history, formula, and
conflict semantics.

The first public release includes:

- A compact, adaptive React `DataTable` for ordinary application records.
- Local and remote data sources, including dataset-wide server sorting,
  filtering, grouping, aggregation, pagination, mutation, refresh, and
  subscription contracts.
- Excel-style structured tables backed by live workbook ranges.
- A contextual Table tab inside the spreadsheet.
- Scoped neutral styling based on the spreadsheet's existing slate, white, and
  muted-green design tokens.
- A publishable React library with typed ESM exports and consumer-build tests.
- A mandatory pressure-testing gate before release.

Laravel integration, Web Components, vanilla-JavaScript mounting, and non-React
framework bindings are excluded from this release. The headless layer remains
DOM-free so those integrations can be added later without redesigning the core.

## Goals

1. Let a React application embed a capable table with a small configuration
   surface comparable to Frappe DataTable.
2. Preserve controlled state, stable identities, typed definitions, and
   source-owned rendering concepts associated with TanStack Table.
3. Reuse spreadsheet editing behavior without copying workbook table data into
   a second model.
4. Support local records, remote records, and workbook tables without hiding
   differences in authority or execution scope.
5. Preserve existing spreadsheet behavior while extracting reusable boundaries
   incrementally.
6. Prevent shipping silent data corruption, stale remote results, inaccessible
   interactions, or performance regressions.

## Non-Goals

- Implementing a backend. The library defines remote query, mutation,
  conflict, and subscription contracts; host applications implement them.
- Pretending a loaded remote page is the complete dataset.
- Evaluating dataset-wide formulas locally when remote rows are not loaded.
- Maintaining one universal mutation implementation for workbook and record
  sources.
- Perfect visual or behavioral parity with every version of Microsoft Excel.
- Publishing partial table implementations before the complete release gate
  passes.

## Current Constraints

The present application is not yet a reusable table library:

- `App.tsx` owns workbook history, formula synchronization, selection, editing,
  panels, persistence, and most commands.
- `Grid.tsx` is controlled through callbacks but remains tied to `SheetModel`,
  A1 addresses, and `FormulaEngine`.
- `Spreadsheet` takes no props and reads and writes browser `localStorage`.
- The package is private, points at TypeScript source, and has no library-mode
  build or `exports` map.
- Styles affect global elements and assume viewport ownership.
- `WorkbookModel` has no structured-table definitions or stable table row and
  column identities.
- The grid row-virtualizes but does not column-virtualize.

The design deliberately addresses these constraints before treating the table
surface as shippable.

## Architecture

### React presentation layer

Public React views are:

- `DataTable`: adaptive embedded table with a compact toolbar, contextual edit
  controls, optional quick-tools drawer, grid viewport, and status summary.
- `Spreadsheet`: the current full spreadsheet shell, refactored to accept a
  session or controlled workbook props while preserving standalone defaults.
- `SpreadsheetTableTab`: an internal contextual ribbon tab shown when the
  current selection intersects a structured table.

The shared interaction kernel owns only browser and presentation concerns:

- Two-axis virtualization and measurement.
- Focus and keyboard navigation.
- Selection gestures.
- Editor positioning and lifecycle.
- Clipboard parsing and browser clipboard integration.
- Context menus, toolbars, render slots, and accessible DOM.

It does not own canonical row values, workbook state, remote revisions, or
server acknowledgements.

### Shared headless contracts

Workbook and record tables share narrow contracts rather than a universal data
store:

```ts
type CommandEnvelope<TIntent> = {
  id: string;
  intent: TIntent;
  transactionId?: string;
  expectedRevision?: string;
};

type CommandResult<TCurrent = unknown> =
  | { status: "committed"; revision: string }
  | { status: "pending"; operationId: string }
  | {
      status: "rejected";
      reason: "validation" | "permission" | "unsupported";
      issues?: readonly TableIssue[];
    }
  | { status: "conflict"; revision: string; current: TCurrent };

interface TableRowView<TRow> {
  getSnapshot(): TableViewSnapshot<TRow>;
  subscribe(listener: () => void): () => void;
  dispatch(intent: TableIntent): Promise<CommandResult>;
}
```

The command coordinator normalizes UI intent and observes its lifecycle. It is
not the authority over remote data and does not require both source modes to use
the same history or formula implementation.

### RecordTableSession

`RecordTableSession<TRow>` supports local or remote record sources.

Local mode:

- The host supplies rows and a stable `getRowId` function.
- Controlled mode emits row updaters and durable metadata changes to the host.
- Uncontrolled mode keeps rows and table metadata inside the session.
- Local history uses an operation journal with bounded inverse operations.
- Full-local-dataset formulas, sorting, filtering, grouping, and aggregation may
  run locally.

Remote mode:

- The host API remains authoritative.
- A normalized page cache is a disposable read projection.
- Optimistic values are overlays, not mutations of authoritative cached rows.
- Queries are abortable and generation checked.
- Mutations include client mutation IDs and source revision or row version
  tokens.
- Remote undo is exposed only when the source supports version-checked
  compensating mutations.

### WorkbookSession

`WorkbookSession` wraps the current `WorkbookModel`, HyperFormula projection,
selection, persistence, and workbook command behavior behind:

```ts
interface WorkbookSession {
  getSnapshot(): WorkbookSnapshot;
  subscribe(listener: () => void): () => void;
  dispatch(command: WorkbookCommand): WorkbookCommandResult;
  table(tableId: string): WorkbookTableSession;
  destroy(): void;
}
```

Workbook commands reduce atomically against current session state. The initial
migration preserves bounded snapshot history to avoid combining history
replacement with controller extraction. Patch or delta history is introduced
only after behavioral parity and memory benchmarks prove it safe.

`WorkbookTableSession` is a live projection over a structured worksheet range.
It reads current workbook state on every snapshot and dispatches workbook
commands directly. It never maintains a page cache or copied row collection.

### Mode-specific source contracts

```ts
type RecordSource<TRow> =
  | {
      kind: "local";
      rows: readonly TRow[];
      getRowId(row: TRow): string;
      onRowsChange?(updater: RowUpdater<TRow>, context: ChangeContext): void;
    }
  | {
      kind: "remote";
      getRowId(row: TRow): string;
      capabilities: RemoteCapabilities;
      query(request: QueryRequest, context: QueryContext): Promise<QueryResult<TRow>>;
      mutate?(batch: readonly Mutation[], context: MutationContext): Promise<MutationResult<TRow>>;
      subscribe?(listener: SourceListener<TRow>): () => void;
    };
```

The workbook projection intentionally does not implement pagination, remote
cache invalidation, optimistic overlays, or server conflict resolution.

## Structured Workbook Tables

`WorkbookModel` advances to a migrated schema that contains workbook-wide table
definitions. Table names are workbook-unique.

```ts
type StructuredTable = {
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

type StructuredTableColumn = {
  id: string;
  name: string;
  sheetColumn: number;
  dataType?: TableDataType;
  calculatedFormula?: string;
  totalsFunction?: TableAggregate;
};
```

Cells, formulas, formats, comments, validation, and hyperlinks remain solely in
`SheetModel`. The table stores metadata and logical identities only.

Identity rules:

- Table and column IDs are non-reused UUID-like identifiers in app-native state.
- API rows require a stable host-supplied ID.
- Workbook row IDs remain stable through insert, delete, sort, filter, and undo
  while the workbook session exists.
- Excel files do not provide dependable logical row IDs. A configured key
  column preserves identity across re-import; otherwise row IDs regenerate.
- XLSX import generates app IDs for table and column objects. XLSX export
  preserves Excel-supported names and behavior, not opaque app IDs.

Workbook table behavior:

- Creating a table validates headers, name uniqueness, range overlap, merged
  cells, and sheet protection before committing atomically.
- Inserting or deleting a table row updates the worksheet, table range, row-ID
  order, formulas, validation, formats, filters, and undo state together.
- Workbook table sorting physically reorders table body rows, matching Excel
  behavior. It sorts by evaluated typed values, moves row IDs with rows, and
  preserves relative formula semantics.
- Filtering changes visibility without copying or deleting source rows.
- The Table tab supports name, range resize, header/totals controls, style,
  calculated columns, filters, export, conversion to range, and opening the
  adaptive table view.

XLSX import and export must use actual Excel structured tables rather than
treating a table as an AutoFilter range. Tests use files produced by Excel in
addition to files produced by the exporter.

## Public React API

### Simple local records

```tsx
<DataTable
  rows={employees}
  columns={columns}
  getRowId={(row) => row.id}
  onRowsChange={setEmployees}
/>
```

### Advanced local or remote session

```tsx
const source = createRemoteTableSource({
  capabilities,
  query: fetchEmployees,
  mutate: saveEmployeeChanges,
  subscribe: watchEmployeeChanges,
});

const session = useTableSession({ source, columns });

<DataTable session={session} ref={tableRef} />
```

Advanced sessions also accept `document`, `defaultDocument`, and
`onDocumentChange` for durable table metadata that is not part of the host row
objects.

### Shared workbook session

```tsx
const session = useWorkbookSession({
  workbook,
  onWorkbookChange: setWorkbook,
});

<Spreadsheet session={session} />
<DataTable session={session.table("employees-table")} />
```

### Component types

`DataTable` accepts either simple local-record props or a prebuilt session. The
union prevents conflicting `rows`, `source`, and `session` combinations.

`ColumnDef<TRow>` supports:

- Stable column ID and display header.
- Accessor and update functions.
- Data type and typed input parser.
- Cell and header render functions.
- Custom editor component.
- Formula or calculated-column definition.
- Formatting and validation.
- Editable and permission predicates.
- Sorting, filtering, grouping, and aggregation configuration.
- Width, minimum and maximum width, visibility, pinning, and metadata.
- Accessor, display-only, and computed column kinds.

`DataTableHandle` exposes:

```ts
interface DataTableHandle {
  focus(): void;
  refresh(): Promise<void>;
  dispatch(intent: TableIntent): Promise<CommandResult>;
  undo(): Promise<CommandResult>;
  redo(): Promise<CommandResult>;
  scrollToRow(rowId: string): void;
  getSelection(): TableSelection;
  export(options: ExportOptions): Promise<Blob>;
}
```

`Spreadsheet` supports controlled and uncontrolled use:

- `session` for shared advanced use.
- `workbook` and `onWorkbookChange` for controlled use.
- `defaultWorkbook` for uncontrolled use.
- An injectable storage adapter, including an explicit disabled value.
- Feature configuration, service injection, host events, and scoped theme
  overrides.

The default standalone behavior continues to load and save browser state.

## State and Feature Configuration

Record sources keep application values authoritative, while a separate metadata
document stores spreadsheet-like cell concerns without modifying host row
objects:

```ts
type TableMetadataDocument = {
  version: 1;
  cells: Readonly<Record<string, TableCellMetadata>>;
  calculatedColumns: readonly CalculatedColumnDefinition[];
  namedStyles: readonly NamedTableStyle[];
};

type TableCellMetadata = {
  format?: TableCellFormat;
  validation?: TableValidation;
  comment?: string;
  formula?: string;
  readOnly?: boolean;
};
```

Cell metadata keys combine stable row and column IDs, never visible indexes.
Simple uncontrolled tables retain this document for the session lifetime.
Controlled local tables use `document/onDocumentChange`. Remote sources either
return and mutate metadata through declared capabilities or leave the related
feature disabled. Workbook tables continue to store equivalent metadata in the
workbook cells and table definition rather than creating this record document.

Durable table state and ephemeral view state are separate:

- Durable: row values, structured-table metadata, formats, validations,
  comments, calculated columns, and source revisions.
- View: sorting projection, filter projection, grouping, pagination, selection,
  column order, pinning, visibility, widths, open panels, and scroll position.

Both `state/defaultState` and `onStateChange` are supported. Supplying a
controlled state slice makes the host responsible only for that slice.

All source-supported features are enabled by default. The effective feature set
is the intersection of source capabilities, component feature configuration,
permissions, and current selection context. An unsupported operation is hidden
or disabled with a reason; it never falls back to page-local behavior silently.

## Data Flow

### Local records

1. `RecordTableSession` normalizes rows by host ID and derives the visible row
   model from controlled view state.
2. User input becomes a typed command intent.
3. The input pipeline retains raw text, parses a typed value or formula, and
   evaluates client validation.
4. The local source applies one atomic row and metadata transaction.
5. A bounded inverse operation is added to local history.
6. Controlled hosts receive a row updater and change context; uncontrolled
   sessions publish their new snapshot.
7. React reads the session through `useSyncExternalStore`.

### Remote records

1. View-state changes produce a serializable query AST and increment the query
   generation.
2. Any obsolete request is aborted.
3. A response is accepted only when its generation and source revision are
   current.
4. Edits pass local parsing and pre-validation, then create an optimistic
   overlay identified by `clientMutationId`.
5. The source receives the mutation, base revision, row versions, and abort
   signal.
6. A committed response removes the overlay. A corrected response replaces it
   with canonical server values. A rejection rolls it back. A conflict keeps
   the attempted value visible as conflicted and supplies the authoritative
   row for resolution.
7. Subscription events are revision ordered. Events touching optimistic rows
   trigger a rebase or explicit conflict, never last-arrival-wins overwrite.

### Workbook tables

1. The Table tab or table view resolves stable table IDs and column IDs to the
   current live worksheet range.
2. A UI action becomes a workbook command.
3. `WorkbookSession` reduces the command against current state in one serialized
   transaction.
4. Structural formulas, table metadata, row IDs, formatting, validation, and
   history update atomically.
5. HyperFormula receives the resulting workbook change and remains a derived
   calculation projection.
6. Both `Spreadsheet` and `DataTable` subscribers render the same committed
   workbook state.

## Remote Execution Semantics

Capabilities describe execution and scope rather than booleans:

```ts
type OperationCapability = {
  executor: "client" | "server";
  scope: "loadedRows" | "completeDataset";
};

type FormulaCapability =
  | "none"
  | "loadedRows"
  | "fullLocalDataset"
  | "server";
```

`RemoteCapabilities` independently defines sort, filter, group, aggregate,
pagination, edit, bulk edit, validation, formula, subscription, and undo support.

Rules:

- Dataset-wide commands require `completeDataset` scope.
- Loaded-row operations are visibly labeled and never presented as global.
- Arbitrary remote formulas require server support. Loaded-row formulas cannot
  treat missing records as blank.
- Client validation is a precheck; server validation is authoritative.
- Undoing a pending mutation cancels it when possible and removes its overlay.
- Undoing a committed remote mutation sends a version-checked compensating
  mutation only when the source advertises that capability.
- Remote grouping and aggregates come from the server unless the full dataset
  is local.
- Offset, cursor, infinite, unknown-total, and known-total pagination are
  explicit result variants.

## Formula and Value Semantics

The input pipeline distinguishes four values:

- Raw edit text.
- Parsed stored value.
- Formula expression, when present.
- Evaluated display value.

Validation receives raw, parsed, and evaluated candidates. Number, boolean,
date, and date-time input is stored with typed semantics and exported with
native XLSX types when supported.

Workbook formulas continue through HyperFormula. Structural edits use a
token-aware or engine-owned reference service shared by insert, delete, fill,
copy, move, and sort. Regex rewriting of A1-looking substrings is prohibited.

Record-table formulas use stable column IDs and source-declared formula scope.
The formula service is a capability module, not a mandatory stage in every
remote record mutation.

## Error Handling and Recovery

Errors are classified and handled at their source:

- Validation: keep focus in the editor, show cell and summary issues, and do not
  mutate data.
- Permission or unsupported operation: disable the command when known; reject
  with an explanatory result if authority changes between render and commit.
- Query failure: retain the last good rows, show retry state, and ignore stale
  responses.
- Mutation rejection: roll back only the matching optimistic overlay and show
  server issues at the affected cells or row.
- Conflict: retain both attempted and authoritative values, mark the row, and
  offer reload or version-checked retry. The library never silently overwrites.
- Formula error: render spreadsheet error values without throwing from React.
- Import or workbook command failure: abort the whole transaction and preserve
  the previous workbook snapshot.
- Renderer, editor, or feature extension failure: isolate the extension behind
  a boundary and keep the rest of the table usable.
- Persistence quota or connector failure: report through session state and host
  callbacks without destroying in-memory work.

Every async operation carries an ID and abort signal. Session destruction aborts
work and prevents post-unmount state publication. Multiple instances have no
module-level state, storage, authentication, or event-bus singleton.

Diagnostic callbacks contain command IDs, timing, category, and safe metadata;
row values are excluded by default.

## Styling and Accessibility

The approved visual direction is adaptive and neutral:

- White work surfaces.
- Slate chrome and hairline borders.
- Dark neutral text.
- Muted green only for selection, active state, and primary action emphasis.
- A compact default toolbar, contextual edit/formula bar, optional quick-tools
  drawer, and an `Open in Spreadsheet` action when a workbook is available.

Styles are scoped beneath component roots, use exported CSS custom properties,
and size against their container. The library does not style global `body`,
`button`, `input`, or `select` elements.

Accessibility requirements include:

- Correct grid, row, column-header, row-header, and gridcell semantics.
- Roving focus and complete keyboard operation.
- Stable focus through virtualization and async refresh.
- Screen-reader announcements for selection, edits, validation, sorting,
  filtering, loading, conflict, and mutation status.
- Visible focus indicators and WCAG AA contrast.
- Reduced-motion support.

## Packaging

The final package exposes:

- `js-spreadsheet/core`: DOM-free types, session interfaces, query AST,
  capabilities, commands, and source helpers.
- `js-spreadsheet/react`: `Spreadsheet`, `DataTable`, providers, hooks, handles,
  render/editor types, and React-specific helpers.
- `js-spreadsheet/connectors/google`: optional Google Sheets integration.
- `js-spreadsheet/styles.css`: scoped default theme and tokens.

The package build produces ESM and declarations, uses an explicit `exports` map,
and treats React and ReactDOM as peer dependencies. Importing the headless entry
does not touch `window`, `document`, storage, or CSS. Package export shape is
finalized after both React consumers exercise the internal APIs, then frozen by
consumer-build and type tests.

The project and HyperFormula licensing requirements remain documented next to
installation instructions. Hosts that cannot comply with GPL requirements need
appropriate licensing before embedding formula-enabled builds.

## Correctness Prerequisites

The following confirmed current behaviors are fixed and regression-tested before
the Table tab exposes structural operations:

1. Workbook dispatch becomes atomic and does not reduce multiple commands
   against stale render-closed history.
2. Structural formula updates become token-aware and cover quoted text,
   digit-bearing function names, scientific notation, sheet-qualified
   references, and cross-sheet dependents.
3. Workbook sorting uses evaluated typed values and preserves formula semantics
   when rows move.
4. Validation evaluates formulas through computed candidates rather than raw
   formula strings.
5. Grid input parses and preserves number, boolean, date, and date-time types.
6. Blank, empty-string, null, zero, and formula-empty filter behavior is
   explicit and typed.
7. History memory remains bounded under large-sheet edit workloads.
8. The shared viewport supports column as well as row virtualization.

Existing shared-formula import and direct native number and boolean XLSX
round-tripping remain protected by their current regression tests. Date and UI
entry fidelity receive additional coverage because typed values can currently
become strings.

## Incremental Implementation Sequence

The sequence controls risk; it does not create partial public releases.

1. Characterize existing spreadsheet behavior with regression and property
   tests.
2. Fix the confirmed correctness prerequisites.
3. Introduce atomic `WorkbookSession` and controlled `Spreadsheet` without
   changing visual behavior.
4. Move App handlers into typed commands feature by feature.
5. Extract the two-axis viewport and shared interaction hooks while retaining
   the existing spreadsheet adapter.
6. Build the local-record `DataTable` and exercise the public API internally.
7. In parallel, build structured workbook tables and native XLSX support, plus
   remote API capabilities and race-safe mutation semantics.
8. Integrate the contextual Table tab and adaptive workbook table view.
9. Finalize package exports, declarations, scoped CSS, and consumer examples.
10. Run the full pressure gate and ship only after all acceptance criteria pass.

Commits remain small and behavior-focused. Each extraction step retains or adds
tests before the next boundary moves.

## Testing and Pressure Gate

The release gate contains all of the following.

### Unit and type tests

- Typed input parsing and display formatting.
- Command reducers and atomic transactions.
- Query and filter AST serialization.
- Capability gating and unsupported-command behavior.
- Validation using raw, parsed, and evaluated candidates.
- Local inverse-operation history.
- Remote mutation and conflict state machines.
- Public TypeScript inference for rows, columns, sessions, and renderers.

### Adapter contracts

- Workbook, local-record, and remote-mock contract suites.
- Identical observable behavior only when both adapters advertise support.
- Explicit disabled and rejected behavior for unsupported capabilities.
- Stable identity through sort, filter, insert, delete, refresh, and undo.

### Property and fuzz tests

- Random command sequences whose inverse restores original state.
- Random structured-table range changes and row-ID ordering.
- Formula reference transformations containing quoted strings, `LOG10`,
  scientific notation, absolute references, ranges, and cross-sheet references.
- Random query generations, page eviction, subscription events, and mutations.

### Race and failure tests

- Out-of-order remote responses.
- Abort during query and mutation.
- Query changes during optimistic edits.
- Subscription updates touching pending rows.
- Undo during pending and committed remote mutations.
- Conflicts, retries, server corrections, offline recovery, and unmount cleanup.

### XLSX fixtures

- Native Excel tables with headers, totals, styles, filters, calculated columns,
  typed values, dates, and shared formulas.
- Import-export-import behavior using both real Excel fixtures and generated
  workbooks.
- Table rename, resize, conversion to range, sort, and row insertion.

### React integration and browser tests

- Controlled and uncontrolled `Spreadsheet`.
- Simple local `DataTable` and session-backed `DataTable`.
- Multiple simultaneous instances.
- React StrictMode lifecycle.
- Package import without browser globals.
- Chromium, Firefox, and WebKit E2E flows.
- Spreadsheet Table-tab and remote conflict workflows.

### Accessibility tests

- Automated axe checks.
- Keyboard-only feature flows.
- Focus retention through editing, virtualization, sorting, filtering, and page
  refresh.
- Screen-reader role, name, state, and announcement assertions.

### Performance and memory tests

- Local tables with at least 100,000 rows and wide-column scenarios.
- Remote simulations with million-row totals and bounded loaded pages.
- Visible-only row and column DOM counts.
- Production-build scroll, edit-to-paint, query swap, and subscription-burst
  benchmarks.
- Bounded history, cache, optimistic overlay, and formula-engine memory.
- Failure on material regression from an accepted pinned-environment baseline.

### Consumer-build tests

- Vite and webpack React consumers.
- Root, core, React, connector, and CSS subpath imports.
- Declaration and generic-inference tests.
- React peer-dependency and duplicate-React checks.

No public release occurs with a failing gate, unresolved P0/P1 defect, silent
page-local fallback, or unverified structured-table XLSX behavior.

## Acceptance Criteria

The feature is ready when:

- A React app can embed a local table using `rows`, `columns`, `getRowId`, and an
  update callback.
- A React app can connect a remote source with explicit query, mutation,
  subscription, conflict, validation, and capability semantics.
- Dataset-wide operations never run against only a loaded page unless explicitly
  requested and visibly labeled.
- A workbook range can become a structured table and remain live in both the
  spreadsheet and adaptive table view.
- Structured tables import and export as real Excel tables.
- Workbook table sort, insert, delete, formulas, validation, formatting, and
  undo behave atomically.
- Existing spreadsheet features and tests retain parity.
- The approved adaptive neutral design is applied consistently.
- Styles do not leak into host applications.
- Public package subpaths build and type-check in consumer applications.
- Accessibility, browser, race, fuzz, XLSX, performance, and memory gates pass.
- Licensing requirements are clear before installation.

## Independent Review Record

Claude and Auggie independently returned `ACCEPT WITH REQUIRED REVISIONS` for
the initial shared-runtime proposal. Their common findings were incorporated:

- The shared runtime coordinates intent but does not own remote truth.
- Read and command-lifecycle contracts are shared; write, history, formula, and
  conflict semantics remain source-specific.
- Remote execution scope and mutation lifecycle are explicit.
- Structured-table metadata and native XLSX round-tripping are first-class.
- Migration is incremental and pressure testing is executable.

A current-code verification pass removed stale audit claims about shared-formula
import and native model number/boolean export. Eighty-three targeted workbook,
XLSX, validation, and formula-engine tests passed during that verification.
