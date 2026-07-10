# Spreadsheet Setup and Structure Editing Design

Date: 2026-07-10

Status: Interaction design approved; pending written-spec review

## Decision Summary

Improve the existing spreadsheet rather than adding a second workbook surface.
The follow-up release will make three user journeys complete and discoverable:

1. **Import Google Sheet** opens a real import and setup dialog. A fresh clone
   explains what must be configured, validates the current browser origin, and
   lets the standalone app retain a public OAuth client ID. Embedded apps can
   inject either a client ID or their own token provider.
2. **Add sheet** remains an immediate Excel-style action. It creates and
   activates the next sheet, keeps the new tab visible, and focuses `A1`.
3. **Insert columns** offers both left and right actions from the ribbon, cell
   context menu, and column-header context menu. The command is atomic and
   structured-table aware.

The same pass fixes the two confirmed DataTable interaction regressions that
were found while reproducing these workflows: clipped column menus and group
expand/collapse controls that require a second click at narrow widths.

This work keeps the existing neutral slate, white, and muted-green visual
language. It does not introduce a separate design system or a second dev
server.

## Goals

1. Give a user who cloned the repository an actionable Google Sheets setup
   experience instead of a status-bar message that appears to do nothing.
2. Keep the standalone workbook within the visible browser viewport without
   imposing viewport sizing on embedded consumers.
3. Make left and right column insertion equally discoverable and keyboard
   accessible.
4. Preserve workbook and structured-table invariants through structural edits,
   including formula, identity, selection, protection, and undo/redo behavior.
5. Preserve existing public APIs where practical and provide an explicit
   compatibility path for the current Google token-provider factory.
6. Verify the changes against desktop, narrow, standalone, and embedded
   layouts using only the existing server on port `4173` during local review.

## Non-Goals

- Google Sheets write-back, refresh, synchronization, or conflict resolution.
- A backend OAuth service or storage of Google client secrets.
- Laravel integration.
- A general ribbon redesign.
- A second spreadsheet model or a copied DataTable representation of workbook
  data.
- Perfect visual parity with every Excel menu.
- Automatically inferring calculated formulas or totals functions for a newly
  inserted structured-table column.

## Confirmed Current Behavior

### Add sheet is hidden by standalone sizing

Both the Home ribbon action and the bottom `+` tab dispatch `sheet.add`. The
command creates and activates `Sheet2`, and the status becomes `Added sheet`.
On a desktop viewport, however, the standalone workbook's percentage heights
have no definite parent height. The grid expands to thousands of pixels and
pushes the sheet tabs below the visible page. Because the new sheet is blank,
the visible grid looks unchanged.

The public `Spreadsheet` component must continue to use host-controlled height.
The sizing correction therefore belongs to the standalone `App` boundary, not
to every embedded spreadsheet instance.

### Column insertion is asymmetric and table unsafe

The current UI exposes **Insert column left** from the Home ribbon and cell
context menu. There is no right-side action or column-header context menu. The
core command is already direction-neutral:

```ts
{ type: "columns.insert", sheetId, index, count }
```

Left and right are UI intentions that resolve to different absolute indexes.
They do not require separate core commands.

The current structural transformer shifts cells, formats, widths, hidden
state, validations, comments, links, protection maps, conditional formats,
filters, charts, merges, ordinary formulas, and named ranges. It does not shift
`WorkbookModel.tables`. Inserting or deleting a sheet column can therefore
leave table ranges and table-column metadata pointing at the wrong cells.

### Google import has configuration and naming gaps

The current action is shown under File but is named **Link Google Sheet**. It
performs a one-time, read-only workbook replacement and does not retain a link.
When `VITE_GOOGLE_CLIENT_ID` was absent at build time, clicking the action only
changed the status bar. A user could not configure it from the application.

Google Identity Services browser OAuth also cannot use the current
`http://192.168.6.232:4173` raw-IP origin. The app must detect that condition
and explain it. It must not send this environment to a localhost URL. A host
that supplies its own `TokenProvider` is not subject to the built-in browser
OAuth origin check.

## User Experience

### Import Google Sheet

The File ribbon uses the label **Import Google Sheet**. Clicking it always opens
`GoogleSheetsImportDialog`; it never relies on `window.prompt` or status-only
feedback.

The dialog has one continuous workflow with state-dependent sections:

1. **Setup status** shows the exact current origin and one of three states:
   blocked by local rules, eligible but registration unverified, or rejected at
   runtime with `origin_mismatch`.
2. **Configuration** appears when no client ID or host token provider is
   available. It explains how to create a Google Cloud project, enable the
   Sheets API, configure consent/test users, create a Web OAuth client, and add
   the displayed origin to Authorized JavaScript origins.
3. **Client ID entry** accepts only a trimmed public client ID with the expected
   `.apps.googleusercontent.com` form. It never asks for a client secret. A
   standalone user completes a separate **Save and continue** action before the
   import form appears.
4. **Sheet input** accepts a Google Sheets URL or spreadsheet ID and validates
   it before starting OAuth.
5. **Replacement warning** makes the operation explicit: **Import and replace
   workbook**. The import remains undoable through the existing preserved
   workbook history.
6. **Progress and errors** stay inside the dialog. A failed attempt leaves the
   URL in place and offers retry. A successful import closes the dialog and
   reports `Imported <spreadsheet title>` rather than `Linked`.

For the current HTTP LAN-IP origin, the dialog disables built-in sign-in and
states that Google browser OAuth needs an authorized HTTPS DNS origin. The
guidance preserves the current IP-hosted application and does not suggest a
localhost-only workflow. If an embedded host supplies a `TokenProvider`, the
dialog proceeds because authentication is then owned by the host.

Origin assessment uses the origin of the document running `Spreadsheet`, not a
parent frame that may be cross-origin. Opaque origins, `file:` documents,
non-loopback HTTP, and raw non-loopback IPv4 or IPv6 literals are blocked.
HTTPS DNS origins are eligible but cannot be confirmed as registered until
Google answers. Google-supported HTTP loopback origins remain technically
eligible for library correctness, but the current LAN-IP workflow never
recommends them. Scheme, hostname, and port must match the configured origin
exactly. Embedded-host documentation also covers iframe popup permissions and
CSP requirements.

The dialog uses a native modal `<dialog>` with `aria-labelledby` and
`aria-describedby`, making the background inert while open. Initial focus goes
to the client-ID field when setup is required, the sheet field when configured,
or Close when the origin is blocked. Invalid submission focuses the first
invalid field. Pending work sets `aria-busy`, disables form inputs and duplicate
submission, but leaves Cancel available. The dialog body scrolls internally at
narrow widths and 200% zoom. Escape/cancel invalidates pending UI results and
restores focus to the ribbon button.

### Add sheet

Both existing controls retain a single shared handler:

- A click immediately creates the next uniquely named sheet.
- The new sheet becomes active in the same workbook transaction.
- Selection resets to `A1`.
- The grid scrolls to and focuses `A1` after the committed render.
- The active sheet tab scrolls into view if the tab strip overflows.
- Status announces `Added <sheet name>` so the result is specific.

No naming dialog is introduced. Rename remains a separate action.

The standalone document sets `html`, `body`, and `#root` to full height and
removes the browser's default body margin. `App` gives its `Spreadsheet` an
explicit `100dvh` height and a standalone class that allows `min-height: 0`.
The existing narrow-screen `100vh` rule becomes `height: 100%` so it cannot
create a second viewport-height constraint.

The public component keeps `height: 100%` and its documented `420px` minimum,
so a custom app continues to determine its size through its containing element
or the existing `style` prop. Embedded hosts must provide at least `420px` of
block size; a host below that supported minimum may overflow. A fixed-height
embedding fixture at or above the minimum must not grow to browser height.

### Insert columns left and right

The Home ribbon Cells group replaces the single left-insert button with an
**Insert columns** split button. Its primary action inserts left, preserving the
current shortcut path, and its menu names both actions explicitly:

- Insert column left
- Insert column right

The cell context menu adds the right-side action beside the existing left-side
action. Right-clicking a column header first selects that full column (or keeps
the existing multi-column selection when the clicked column is already in it),
then opens a column menu containing insert left, insert right, and delete.

For a normalized selection spanning `N` columns:

```ts
count = selection.end.column - selection.start.column + 1
leftIndex = selection.start.column
rightIndex = selection.end.column + 1
```

Both directions insert `N` columns. The committed selection covers the newly
inserted columns and retains the selection's row span. A whole-column selection
therefore remains a whole-column selection.

## Architecture

### Standalone and embedded sizing boundary

`App` is the standalone composition root. It supplies viewport height and the
standalone Google client-ID store. `Spreadsheet` remains the reusable,
host-sized component.

`SheetTabs` owns active-tab visibility because it renders and identifies the
tab buttons. When `activeSheetId` changes, it scrolls only the active tab using
`block: "nearest"` and `inline: "nearest"`; it does not scroll the document.

`GridScrollApi` gains a focused-cell operation in addition to
`ensureCellVisible`. The add-sheet handler calls it only after `sheet.add`
commits and the active sheet has rendered. Failed or rejected commands do not
move focus.

### Direction-neutral insertion command

The UI centralizes structural insertion in one handler:

```ts
insertSelectedColumns(direction: "left" | "right"): WorkbookCommandResult
```

The handler performs the existing sheet-protection check, derives index and
count, identifies any selected structured-table context, and dispatches one
transaction containing column insertion and selection update.

The contextual candidates are tables whose physical ranges intersect the
selection's row span and contain at least one selected column. A cell selection
normally identifies one table; a whole-column selection can identify vertically
separate tables on the same columns. `expandTableIds` contains only candidates
whose left or right boundary exactly equals the requested insertion index.
Tables with a strictly internal insertion expand deterministically and do not
need an ID in this ambiguity list.

The serializable command stays absolute and gains only the context required to
resolve structured-table boundary ambiguity:

```ts
type InsertColumnsCommand = {
  type: "columns.insert";
  sheetId: string;
  index: number;
  count: number;
  expandTableIds?: readonly string[];
};
```

`expandTableIds` is not a direction flag. It records which selected tables are
intended to consume a new column when insertion occurs exactly at their left or
right boundary. This matters when two tables are adjacent. Invalid or stale
table IDs reject the command rather than silently applying different
semantics. Before generating IDs or mutating state, the reducer verifies that
the list is duplicate-free, every table exists on `sheetId`, and every listed
table has `index` equal to `range.start.column` or `range.end.column + 1`.
Direct API calls receive the same overlap preflight as UI-generated commands.

Core command validation also enforces structural sheet protection so custom
apps cannot bypass it by dispatching directly.

### Structured-table column transformation

Column insertion updates the sheet and all tables on that sheet atomically.
For each table:

- An insertion strictly before the table shifts its range and every
  `sheetColumn` right.
- An insertion strictly after the table leaves it unchanged.
- An insertion where `table.start.column < index <= table.end.column` is inside
  the table. It widens the range, shifts columns at or after the insertion
  point, and creates table-column metadata for every inserted physical column.
- An insertion at `table.start.column` expands the table only when it is in
  `expandTableIds`; otherwise the whole table shifts right.
- An insertion at `table.end.column + 1` expands the table only when it is in
  `expandTableIds`; otherwise the table is unchanged.

New structured-table columns:

- Receive non-reused IDs from the workbook session's `createId` service.
- Receive unique `ColumnN` names using the table's existing normalized-header
  rules.
- Write those names into the header row when headers are enabled.
- Inherit table visual styling through the table's range and style metadata.
- Start with no calculated formula, totals function, data type, filter, sort,
  or key role.

Existing column IDs, names, formulas, totals metadata, filters, sorts, and key
identity remain attached to their logical columns. Ordinary A1 formulas are
rewritten by the existing structure rewriter. Structured references continue
to resolve by table and column identity; the inserted column participates in a
multi-column structured range according to its new order without renaming
existing references.

`StructuredTableColumn.calculatedFormula` is the canonical first-body-row
formula used when table rows are regenerated. Each surviving canonical formula
is rewritten with the same structural operation and its pre-edit sheet context
as the corresponding worksheet formula. A later table row insertion, sort, or
calculated-column regeneration must therefore reproduce the already shifted
formula rather than restoring stale A1 references.

Structural command bounds are strict. `index` and `count` must be integers,
`count` must be positive, insertion permits `0 <= index <= dimension`, and
deletion requires `0 <= index < dimension` plus `index + count <= dimension`.
Out-of-range custom-app commands reject before mutation rather than relying on
the model helper's current clamping behavior.

Deletion is updated in the same structural layer because an insert-only repair
would leave the existing Delete column action corrupting tables. The deletion
interval is half-open, `[index, index + count)`, as is a table's physical column
interval `[range.start.column, range.end.column + 1)`:

- If deletion ends at or before a table's start, its range and all column
  coordinates shift left by `count`.
- If deletion starts at or after a table's end, the table is unchanged.
- On overlap, logical columns inside the deletion interval are removed.
  Surviving columns before the interval keep their coordinates, while surviving
  columns after it shift left by `count`; the table range becomes the minimum
  and maximum surviving physical columns.
- The reducer preflights every affected table before changing any sheet data or
  generating IDs. If any table would lose every column or overlap another
  resulting table, the entire command rejects.

- Deleted table columns are removed from metadata.
- Sort entries referencing removed IDs are removed independently. Because a
  Boolean filter tree cannot safely drop one leaf without changing AND/OR
  meaning, the entire table filter is cleared if any leaf references a removed
  column.
- A removed key column clears `keyColumnId`.
- Surviving logical columns retain their IDs and metadata.
- Deleting every physical column of a table is rejected; users must explicitly
  convert or delete the table.

The transformer must preserve these invariants after every committed command:

```text
table.columns.length == table range width
table.columns[i].sheetColumn == table.range.start.column + i
table header names are nonblank and unique
table IDs are unique across the workbook
column IDs are unique within their owning table
table rowIds length == table body height
table sort/filter/key references point to surviving column IDs
tables remain within their owning sheet and never overlap
```

The existing generic row insert/delete actions are outside the requested UI
addition, but they share the same current corruption risk. They use the same
strict command bounds and preflight every table on the edited sheet:

- Row insertion at or before `table.range.start.row` shifts the whole table
  range down by `count` and preserves row IDs.
- Row insertion at or after `table.range.end.row + 1` leaves the table
  unchanged.
- Row insertion where `table.range.start.row < index <=
  table.range.end.row` is inside the table and rejects with an actionable issue.
- Row deletion uses `[index, index + count)`. A deletion ending at or before the
  table start shifts the range up by `count`; a deletion starting after the
  table end leaves it unchanged; any overlap with header, body, or totals rows
  rejects the whole command.

Existing table-row commands remain the supported route for structural edits
inside table bodies. Thus even an unrelated generic row edit before a table
cannot leave table coordinates stale.

### Google connector boundary

Browser OAuth configuration becomes an optional nested service instead of a
compile-time variable hidden inside `SpreadsheetWorkbook`:

```ts
type GoogleSheetsServiceConfiguration = Readonly<{
  clientId?: string;
  tokenProvider?: TokenProvider;
  tokenProviderFactory?: (clientId: string) => TokenProvider;
  clientIdStorage?: GoogleClientIdStorage | false;
}>;

type GoogleClientIdStorage = Readonly<{
  load(): string | null | Promise<string | null>;
  save(clientId: string): void | Promise<void>;
  clear(): void | Promise<void>;
}>;
```

`SpreadsheetServices` gains `googleSheets?:
GoogleSheetsServiceConfiguration`. The current flat
`googleTokenProviderFactory` remains as a deprecated fallback for one release.

Configuration resolution has two separate stages so every compatibility path
is reachable:

1. A host-supplied `tokenProvider` wins and bypasses client-ID and origin setup.
2. Otherwise resolve the client ID from host configuration, then standalone
   `VITE_GOOGLE_CLIENT_ID`, then standalone saved storage.
3. With that client ID, select the nested `tokenProviderFactory`, then the
   deprecated flat `googleTokenProviderFactory`, then the built-in browser
   provider.

The standalone `App` opts into a local-storage adapter under a namespaced key.
The reusable `Spreadsheet` does not write Google configuration to storage
unless the host explicitly supplies `clientIdStorage`. OAuth access tokens stay
inside the in-memory token provider and are never written by this feature.

Host- and environment-supplied client IDs are managed, read-only values in the
dialog. A storage-supplied ID has **Change client ID** and **Forget client ID**
actions. A manually entered ID is available for the current component session
immediately after validation and is persisted only by the separate **Save and
continue** action. Load, save, or clear failures are recoverable: the dialog
shows an inline warning and continues with session-only state where possible.
Forgetting a saved ID clears the configured provider cache and returns to setup;
it never clears workbook storage or an OAuth token owned by an injected host
provider.

The built-in Google Identity Services script is preloaded after a client ID and
eligible origin are available. The import button remains disabled until the
token client is ready. Its click handler validates the already-entered sheet ID
and invokes token acquisition synchronously in that user gesture before
awaiting fetch or persistence work, preserving popup user activation.

Pure helpers validate client IDs, origins, and sheet URLs independently of
React. `googleAuth.ts` and `googleSheets.ts` share typed connector failures,
exported from the optional Google connector entrypoint, instead of forcing the
UI to inspect message strings:

```ts
type GoogleSheetsErrorCode =
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
```

The component maps each code to a concise recovery action and continues to
emit `onError` diagnostics for host applications. Google API error bodies are
parsed only far enough to distinguish the documented cases; unrecognized
responses use `unknown` without leaking tokens or response bodies.

### DataTable overlay and narrow-click corrections

`DataTableColumnMenu` uses the browser top layer through the Popover API while
remaining a DOM descendant of the DataTable root. This escapes the root's
rounded `overflow: hidden` boundary without breaking scoped CSS variables. Its
position is derived from the trigger rectangle and clamped to the current
viewport. Native outside-click and Escape behavior close it and focus returns
to the trigger. Viewport scrolling, column hiding, or trigger unmounting also
closes the menu so a virtualized header cannot leave an orphaned popover.

The group expand/collapse button handles pointer-down and click directly and
stops propagation into cell selection. One click always toggles the row; grid
selection and horizontal scrolling do not consume the first activation. The
control retains a minimum accessible hit target at narrow widths.

## State and Error Handling

- All workbook structural changes remain session commands and participate in
  the existing bounded history. UI state never mutates `WorkbookModel`
  directly.
- Column insertion plus selection update is one transaction and one undo step.
- New structured-table column IDs are generated through the injected session
  service so tests and controlled hosts stay deterministic.
- A rejected structure command leaves workbook, selection, focus, and history
  unchanged and reports the first actionable `TableIssue`.
- Closing the Google dialog cancels only dialog UI. It does not erase an
  explicitly saved standalone client ID. It invalidates any pending import so
  a late response cannot replace the workbook after the dialog closes.
- Google import is single-flight. Inputs and submit are disabled while one
  attempt is pending; Retry becomes available only after failure. Cancel stays
  available and invalidates the attempt, so a late completion cannot replace
  the workbook.
- Import success uses the current `session.replaceWorkbook(...,
  { history: "preserve", origin: "import" })` contract.
- Host callback exceptions continue to be isolated through the existing
  diagnostic boundary.

## Accessibility

- Dialog title, description, setup status, errors, and progress are connected
  with `aria-labelledby`, `aria-describedby`, `role="alert"`, and live status
  semantics as appropriate.
- The insert split button exposes `aria-haspopup`, `aria-expanded`, arrow-key
  menu navigation, Escape, and focus restoration through the existing
  `SplitButton` behavior.
- Column headers remain keyboard focusable. `Shift+F10` or the context-menu key
  opens the same column menu as right-click.
- Context menus use labelled menu items and do not encode direction using icons
  alone.
- Sheet creation announces the new sheet name and moves focus to the active
  cell only after the command succeeds.
- Popover menus and modal content remain usable at 200% zoom and narrow
  viewport widths.

## Testing and Pressure Gate

### Pure model and command tests

- Insert one and multiple columns at sheet start, middle, and end.
- Insert left and right for cell, range, and whole-column selections.
- Shift a table that is after the insertion point.
- Expand at every internal table position and both explicit boundaries.
- Handle adjacent tables without expanding the wrong table.
- Preserve existing table column IDs and generate deterministic unique IDs for
  new columns.
- Preserve header, body, totals, filters, sorts, key column, calculated
  columns, and structured-reference behavior.
- Delete columns before, inside, across, and after a table; reject deletion of
  the final table column.
- Reject out-of-bounds insert/delete commands and preflight multi-table
  deletion atomically.
- Shift table ranges for safe generic row edits before a table, leave later
  tables unchanged, and reject every row interval that enters a table.
- Reject stale `expandTableIds`, protected sheets, and unsafe generic row
  operations without history changes.
- Verify A1 formulas on the edited sheet and other sheets, named ranges,
  merges, validation, formats, comments, hyperlinks, charts, filters, hidden
  state, and protection maps.
- Rewrite canonical calculated-column formulas and verify that a subsequent
  table row insert and sort do not reintroduce pre-edit references.
- Verify insert and undo restore byte-equivalent workbook state and redo returns
  the exact inserted state.

Property tests generate valid sheets with multiple non-overlapping tables,
random insertion indexes, and counts. Every committed result is checked against
the structured-table invariants above and the workbook migration validator.
The test retains a bounded case count suitable for normal CI rather than adding
an unbounded soak job.

### Component tests

- Fresh-clone Google action opens setup instead of changing only the status
  bar.
- Standalone Google configuration saves only a valid public client ID;
  embedded instances remain Google-client-ID-storage-neutral by default. This
  does not change the Spreadsheet's separate workbook-storage defaults.
- Host token provider bypasses browser-origin validation and is cached per
  component instance.
- Managed, saved, session-only, changed, forgotten, and unavailable-storage
  client-ID states follow the documented precedence and lifecycle.
- Origin tests distinguish blocked, eligible-but-unverified, and runtime
  mismatch states across DNS, loopback, raw IPv4/IPv6, HTTP, HTTPS, file, and
  opaque document origins.
- The preloaded GIS provider begins token acquisition in the import button's
  user gesture; duplicate submissions remain disabled while pending.
- Missing config, incompatible origin, invalid URL, popup failures, API
  disabled, access denied, not found, rate limit, network failure, retry, stale
  completion, success, and cancel states are covered.
- Add sheet activates the new sheet, selects and focuses `A1`, announces its
  name, and scrolls the active tab into view.
- Ribbon, cell menu, mouse column menu, and keyboard column menu dispatch the
  same absolute command semantics.
- DataTable column menus remain visible near every viewport edge.
- A single narrow-screen group-toggle click collapses and expands without
  unintended scrolling.

### Browser and packaging checks

Required release checks are:

```bash
corepack pnpm run typecheck
corepack pnpm test
corepack pnpm run build
corepack pnpm run build:lib
```

Playwright configuration accepts an external `E2E_BASE_URL`. When it is set,
Playwright does not start its configured web server. Local review uses the one
already-running preview:

```bash
E2E_BASE_URL=http://192.168.6.232:4173 corepack pnpm run test:e2e
```

Browser coverage includes desktop and narrow workbook layouts, a fixed-height
embedding fixture, DataTable overlay edges, and both insert directions. A real
Google login is not automated because credentials and consent are external;
the token provider and Google HTTP responses are deterministically mocked. The
live LAN-IP review verifies that the setup dialog reports the origin limitation
accurately.

## Documentation Changes

- `docs/google-sheets-connector.md` describes Import rather than Link, the
  setup dialog, service injection, storage rules, and HTTPS-origin requirement.
- `docs/getting-started.md` points fresh clones to the in-app setup workflow.
- `docs/embedding.md` documents `services.googleSheets`, host token providers,
  the required Google Identity Services and Sheets API CSP origins, and the
  fact that embedded instances do not persist client IDs by default.
- `docs/features.md` continues to describe Google Sheets as read-only import.
- Existing examples avoid localhost-only instructions for the current review
  environment.

## Acceptance Criteria

1. At `192.168.6.232:4173`, clicking **Import Google Sheet** opens the setup
   dialog, shows the exact origin, and explains why built-in OAuth cannot run on
   that origin. The button never appears inert.
2. With a compatible configured origin or injected token provider, a valid URL
   imports once, replaces the workbook through session history, and reports
   **Imported**, not **Linked**.
3. On desktop and narrow layouts, Add sheet immediately creates and activates
   the next sheet, keeps its tab visible, and focuses `A1`.
4. The standalone workbook fits the viewport. An embedded spreadsheet obeys
   its host container and is not forced to viewport height.
5. Users can insert left or right from the ribbon, a cell context menu, and a
   column-header context menu. Multi-column selections insert the same count.
6. Inserting beside or inside a structured table expands that intended table,
   creates valid logical table columns, and preserves all existing column
   identities and references.
7. Column deletion and guarded row operations cannot leave an invalid table.
8. DataTable column menus are not clipped and group rows respond on the first
   click at narrow widths.
9. The required unit, property, browser, build, and library-package checks pass
   without launching an additional local dev server during live review.

## Delivery

Implementation is developed on `feat/spreadsheet-ux-followup`, based directly
on the squash-merged `origin/main`. After verification, it is published as a
new pull request rather than appended to the already merged DataTable branch.
The pull request includes the test evidence above and is not squash-merged
until required checks and actionable review comments are resolved.
