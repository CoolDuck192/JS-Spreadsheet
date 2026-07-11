# Embedding Spreadsheet and DataTable

Use the surface that matches the host application:

- Use `DataTable` with local rows for an in-app table whose data is already in
  the browser.
- Use a remote table session when an API owns querying, editing, conflicts, and
  version history.
- Use `Spreadsheet` and workbook structured tables when the data belongs in an
  Excel-compatible workbook. The same table can open in the focused DataTable
  view without copying it out of the workbook.

The UI package is React-only. The core sessions and source contracts are
framework-independent and do not access the DOM.

## Browser support

The React UI supports current evergreen Chrome/Edge, Firefox, and Safari
releases. Host browsers must provide native CSS `@scope`, container queries,
`:has()`, and `color-mix()` support; legacy browsers that discard scoped rules
are not supported. This contract keeps workbook styles isolated instead of
shipping a second, unscoped stylesheet that could leak into the host app. The
DOM-free `js-spreadsheet/core` entrypoint is not subject to these UI/CSS
requirements.

## Build and install

The package is currently private and is not published to a registry. Build it
from a checkout before linking or packing it:

```bash
# In JS-Spreadsheet
corepack pnpm install --frozen-lockfile
corepack pnpm run build:lib

# In the host application
pnpm add file:/absolute/path/to/JS-Spreadsheet
pnpm add react react-dom
```

For a reproducible artifact, run `corepack pnpm pack` after `build:lib` and add
the resulting `.tgz` file in the host application. Node 22.13 or newer is
required to build this repository. React and ReactDOM are peer dependencies;
supported React versions are 18.2 through 19.

Import the default styles once in the host application's browser entry:

```ts
import "js-spreadsheet/styles.css";
```

The package entrypoints are deliberately split:

| Import | Contents |
| --- | --- |
| `js-spreadsheet` | Alias of the React entry; defaults to `Spreadsheet` |
| `js-spreadsheet/react` | `Spreadsheet`, `DataTable`, workbook table view, hooks, providers, handles, and React column types |
| `js-spreadsheet/core` | DOM-free workbook/table sessions, local and remote sources, query types, commands, and XLSX functions |
| `js-spreadsheet/connectors/google` | Optional Google Sheets authentication and import helpers |
| `js-spreadsheet/styles.css` | Scoped workbook and DataTable theme |

Do not deep-import files under `src/`; those paths are implementation details.

## Local rows

The shortest integration gives `DataTable` rows, typed columns, and a stable
row ID. Add an `update` function to a column to make it editable.

```tsx
import { useState } from "react";
import { DataTable, type ColumnDef } from "js-spreadsheet/react";
import "js-spreadsheet/styles.css";

type Employee = {
  id: string;
  name: string;
  department: string;
  salary: number;
};

const initialEmployees: readonly Employee[] = [
  { id: "e-1", name: "Ada", department: "Engineering", salary: 120_000 },
  { id: "e-2", name: "Grace", department: "Finance", salary: 110_000 }
];

const employeeColumns: readonly ColumnDef<Employee>[] = [
  {
    id: "name",
    header: "Name",
    dataType: "text",
    accessor: (row) => row.name,
    update: (row, value) => ({ ...row, name: String(value) }),
    sortable: true,
    filterable: true,
    width: 180
  },
  {
    id: "department",
    header: "Department",
    dataType: "text",
    accessor: (row) => row.department,
    update: (row, value) => ({ ...row, department: String(value) }),
    sortable: true,
    filterable: true,
    groupable: true
  },
  {
    id: "salary",
    header: "Salary",
    dataType: "number",
    accessor: (row) => row.salary,
    update: (row, value) => ({ ...row, salary: Number(value) }),
    validate: ({ parsed }) => Number(parsed) < 0
      ? [{ code: "salary-negative", message: "Salary cannot be negative" }]
      : [],
    sortable: true,
    filterable: true,
    aggregatable: ["sum", "average", "count"]
  }
];

export function EmployeesTable() {
  const [rows, setRows] = useState<readonly Employee[]>(initialEmployees);

  return (
    <DataTable
      aria-label="Employees"
      rows={rows}
      columns={employeeColumns}
      getRowId={(row) => row.id}
      onRowsChange={(update) => setRows((previous) => update(previous))}
      rowSelection="multiple"
      inlineFilters
      layout="fluid"
    />
  );
}
```

Local sorting, filtering, grouping, aggregation, offset pagination, editing,
metadata, validation, export, and undo/redo operate on the complete supplied
dataset, not only the visible viewport. Column capabilities and the optional
`features` prop determine which controls are enabled. Formulas require a
`formulaService`; they are not guessed from arbitrary row objects.

`rows`, view `state`, and metadata `document` each support controlled and
uncontrolled ownership. When a slice is controlled, provide its matching change
callback. Change callbacks receive an updater and a context containing the
command ID, reason, and revision.

### Supply a prebuilt local session

A session is useful when multiple components or application services need to
dispatch table commands or observe the same snapshot.

```tsx
import { useEffect, useMemo } from "react";
import { createLocalRecordTableSession } from "js-spreadsheet/core";
import { DataTable } from "js-spreadsheet/react";

export function SessionEmployeesTable() {
  const session = useMemo(() => createLocalRecordTableSession({
    source: {
      kind: "local",
      rows: initialEmployees,
      getRowId: (row: Employee) => row.id
    },
    columns: employeeColumns,
    historyLimit: 100
  }), []);

  useEffect(() => () => session.destroy(), [session]);

  return <DataTable aria-label="Session employees" session={session} />;
}
```

`DataTable` does not start, stop, or destroy a supplied session. The code that
creates it owns its lifetime. The React `useTableSession` hook is the alternative
when the component should own that lifecycle.

## Remote API source

The remote source contract keeps the host API authoritative. The table sends a
typed query AST and receives normalized rows plus an opaque revision. It never
silently turns a server operation into a visible-page-only client operation.

This example uses offset pagination, optimistic versioned edits, and
compensating undo:

```tsx
import {
  createRemoteTableSource,
  type QueryResult,
  type RemoteMutationResult,
  type TableAbortSignal,
  type TableCapabilities
} from "js-spreadsheet/core";
import {
  DataTable,
  useTableSession
} from "js-spreadsheet/react";

const serverOperation = {
  executor: "server",
  scope: "completeDataset"
} as const;

const employeeCapabilities = {
  sort: serverOperation,
  filter: serverOperation,
  group: false,
  aggregate: false,
  pagination: { ...serverOperation, modes: ["offset"] },
  edit: serverOperation,
  bulkEdit: serverOperation,
  metadata: false,
  validation: false,
  formula: "none",
  subscription: false,
  undo: serverOperation,
  export: false
} satisfies TableCapabilities;

async function postJson<T>(
  url: string,
  body: unknown,
  signal: TableAbortSignal
): Promise<T> {
  const response = await fetch(url, {
    method: "POST",
    headers: { "content-type": "application/json" },
    body: JSON.stringify(body),
    signal: signal as AbortSignal
  });
  if (!response.ok) throw new Error(`Table API failed: ${response.status}`);
  return response.json() as Promise<T>;
}

const employeeSource = createRemoteTableSource<Employee>({
  getRowId: (row) => row.id,
  capabilities: employeeCapabilities,
  paginationMode: "offset",
  mutationMode: "versioned",
  undoMode: "compensating",
  compareRevisions(candidate, current) {
    const left = Number(candidate);
    const right = Number(current);
    if (!Number.isFinite(left) || !Number.isFinite(right)) return "unknown";
    return left < right ? "older" : left > right ? "newer" : "equal";
  },
  query: (request, { signal, generation, operationId }) =>
    postJson<QueryResult<Employee>>(
      "/api/employees/table/query",
      { request, generation, operationId },
      signal
    ),
  mutate: (mutations, { signal, operationId }) =>
    postJson<readonly RemoteMutationResult<Employee>[]>(
      "/api/employees/table/mutate",
      { mutations, operationId },
      signal
    )
});

export function RemoteEmployeesTable() {
  const session = useTableSession({
    source: employeeSource,
    columns: employeeColumns,
    defaultState: {
      pagination: { kind: "offset", offset: 0, limit: 50 }
    },
    cache: { maxPages: 5 }
  });

  return (
    <DataTable
      aria-label="Remote employees"
      session={session}
      inlineFilters
      layout="fluid"
    />
  );
}
```

The endpoint paths and JSON envelopes above are examples; the callbacks may
adapt any API. Their returned values must satisfy `QueryResult` and
`RemoteMutationResult`. Every normalized data item needs a nonblank stable ID
equal to `getRowId(item.original)`. The source declaration is validated, so an
edit capability without `mutate`, an export capability without `export`, or a
pagination mismatch fails early.

Capability scope matters:

- `completeDataset` means the declared executor applies the operation to the
  whole authoritative result set.
- `loadedRows` means it is safe only for rows currently loaded by the session.
- A host can require complete-dataset behavior through feature configuration;
  the UI then disables a loaded-row-only operation with an explanation.

Queries are abortable and generation-checked. Edits appear as optimistic
overlays while the server response is pending; the page cache remains an
authoritative projection. A mutation response can commit, correct, reject, or
conflict. Conflicts remain visible with explicit **Reload server value** and
**Retry my change** actions. If the API pushes changes, also declare
`subscription: true` and provide `subscribe(listener)`; implement real revision
ordering in `compareRevisions` so stale events can be ignored.

Remote undo is intentionally narrower than local undo:

- Advertise it only with `undoMode: "compensating"`, versioned mutations, a
  `mutate` callback, and an `undo` capability.
- Undoing a pending batch cancels the whole batch, removes its overlays, and
  refreshes authoritative data. A late response cannot overwrite that refresh.
- Undoing an acknowledged batch sends one inverse batch using the original
  values or full metadata, the acknowledged base revision, and the latest
  acknowledged row version.
- A rejected, uncertain, or conflicting compensation remains retryable. Remote
  redo is not supported. The session retains at most 100 acknowledged batches.

`useTableSession` starts the remote session after mount, stops it during cleanup,
and destroys the owned session. If using `createRemoteTableSession` directly,
call `start()` before rendering, and own `stop()`/`destroy()` yourself; a
session-backed `DataTable` does not own those calls.

## Spreadsheet and workbook structured tables

Embed the full spreadsheet as controlled application state, or pass
`defaultWorkbook` for internally owned state. Setting `storage={false}` prevents
the default browser-local persistence when the host owns saving.

```tsx
import { useState } from "react";
import { createBlankWorkbook } from "js-spreadsheet/core";
import { Spreadsheet } from "js-spreadsheet/react";
import "js-spreadsheet/styles.css";

export function WorkbookEditor() {
  const [workbook, setWorkbook] = useState(() => createBlankWorkbook());

  return (
    <Spreadsheet
      workbook={workbook}
      onWorkbookChange={setWorkbook}
      storage={false}
      features={{ structuredTables: true }}
    />
  );
}
```

The built-in workflow is:

1. Select a worksheet range, then choose **Insert → Table**.
2. Select a cell inside that table to reveal the contextual **Table** tab.
3. Use that tab to edit the name/range, header and totals rows, totals function,
   style, key column, calculated column, and filter.
4. Choose **Open table view** for the focused DataTable surface. Edits, sorting,
   filtering, metadata, formulas, and undo still dispatch through the same
   `WorkbookSession`; this is not a copied row array.
5. Use **Open in Spreadsheet** to return to the selected table cell.

To place only a workbook table view in a custom layout, project it from a
workbook session:

```tsx
import { useState } from "react";
import {
  WorkbookTableView,
  useWorkbookSession
} from "js-spreadsheet/react";
import type { WorkbookModel } from "js-spreadsheet/core";

export function WorkbookTablePanel({
  initialWorkbook,
  tableId,
  onClose
}: {
  initialWorkbook: WorkbookModel;
  tableId: string;
  onClose(): void;
}) {
  const [workbook, setWorkbook] = useState(initialWorkbook);
  const workbookSession = useWorkbookSession({
    workbook,
    onWorkbookChange: setWorkbook,
    storage: false
  });

  return (
    <WorkbookTableView
      session={workbookSession.table(tableId)}
      onClose={onClose}
    />
  );
}
```

### XLSX tables and identity

Workbook import/export uses native Excel tables rather than flattening them to
styled ranges. Names, ranges, headers, styles, calculated columns, supported
value/custom filters, and standard or custom totals metadata round-trip.

Internal identity has an explicit boundary. Table IDs and column IDs are fresh
on every import and are never written into OOXML. Row IDs are also regenerated
unless the importer is given a key column:

```ts
import {
  exportWorkbookToXlsx,
  importWorkbookFromXlsx
} from "js-spreadsheet/core";

const workbook = await importWorkbookFromXlsx(fileBytes, {
  tableKeys: {
    SalesTable: { columnName: "Order ID" }
  }
});

const xlsxBytes = await exportWorkbookToXlsx(workbook);
```

Table and column names in `tableKeys` match case-insensitively. Configured key
values produce stable typed SHA-256 row IDs across reorder and re-import. Blank
or duplicate keys reject the complete import. The imported `keyColumnId` points
to the newly generated matching column ID. Without `tableKeys`, do not persist
an imported row ID as an external business identifier.

`WorkbookTableSession.export()` returns a DOM-free artifact with `bytes`,
`mediaType`, and `fileName`:

```ts
const artifact = await workbookSession.table(tableId).export({
  format: "xlsx",
  scope: "completeDataset",
  fileName: "sales.xlsx"
});
```

Table-only XLSX export supports `completeDataset` only and retains the active
table filter as native metadata while including every body row. Formulas must be
self-contained in the selected table/sheet projection; references to another
sheet, named range, structured table, or outside cell are rejected. Use full
workbook export for those cases. CSV table export supports both `currentView`
and `completeDataset`, emits UTF-8 CRLF/RFC 4180 data, exports evaluated formula
results, and always escapes spreadsheet-injection prefixes.

XLSX input is preflighted before ExcelJS reads it. Invalid native table names,
case-normalized duplicates, unsafe XML, encrypted/ZIP64 archives whose bounds
cannot be proven, archive bombs, and external or escaping table relationships
are rejected rather than partially imported.

## Google Sheets import

**File → Import Google Sheet** is a one-time, read-only workbook replacement;
it does not create a live link, refresh job, write-back path, or synchronization
state. Hide the command when it does not belong in the host workflow:

```tsx
<Spreadsheet features={{ googleSheets: false }} />
```

Embedded hosts configure authentication through `services.googleSheets`:

```tsx
<Spreadsheet
  services={{
    googleSheets: {
      tokenProvider,
      // Or use a managed client ID and a host factory:
      // clientId,
      // tokenProviderFactory: (configuredClientId) => createProvider(configuredClientId),
      // clientIdStorage
    }
  }}
/>
```

Use `tokenProvider` to integrate an existing host Google session. Otherwise,
pass `clientId` and optionally `tokenProviderFactory`; without a custom factory,
the component uses the built-in Google Identity Services provider. Embedded
client-ID persistence is opt-in: the reusable `Spreadsheet` writes nothing
unless the host supplies a `GoogleClientIdStorage` as `clientIdStorage`. The
standalone app is the only surface that automatically stores the public client
ID, under `javascript-spreadsheet.google-client-id.v1`. Access tokens stay in
memory.

Built-in browser OAuth requires an eligible HTTPS DNS origin registered exactly
in the Google OAuth client's Authorized JavaScript origins. A raw LAN IP is
diagnosed as incompatible; serve the host from an HTTPS DNS name or inject a
host `tokenProvider`. The host Content Security Policy must allow:

- `script-src https://accounts.google.com`
- `connect-src https://accounts.google.com https://sheets.googleapis.com`
- `frame-src https://accounts.google.com`

A sandboxed iframe must also permit popup sign-in, for example with
`sandbox="allow-scripts allow-same-origin allow-popups"`, and the top-level host
must not block the Google sign-in popup. See
[google-sheets-connector.md](google-sheets-connector.md) for standalone setup,
failure behavior, and connector exports.

## Styling, sizing, and accessibility

The default stylesheet is scoped below
`.js-spreadsheet-root.js-spreadsheet-workbook` and
`.js-spreadsheet-root.js-spreadsheet-data-table`; it does not intentionally
restyle host `body`, `button`, `input`, or `select` elements. Both surfaces size
against their container. A spreadsheet host must provide at least `420px` of
block size and remains responsible for the component's height; set an explicit
host height and render the spreadsheet at `height: 100%` when a full-height grid
is desired. The component's `420px` minimum remains in effect, and an embedded
spreadsheet does not take ownership of the browser viewport.

Theme variables use the `--js-spreadsheet-*` namespace. Override them on one
component root or a wrapping host class to keep multiple embedded instances
independent. Keep visible focus styles and sufficient contrast when overriding
the defaults.

The grid includes keyboard navigation, selection, editing, copy/paste,
virtualized rows and columns, grid semantics, live status announcements, and
reduced-motion handling. Use a meaningful `aria-label` on every standalone
`DataTable`, especially when a page contains more than one.

## Development URLs

Bind the development server to all interfaces and open it with the machine IP:

```bash
corepack pnpm run dev -- --host 0.0.0.0 --port 5173
hostname -I
```

```text
http://<machine-ip>:5173/             # spreadsheet
http://<machine-ip>:5173/datatable    # DataTable demo
```

Production preview also binds all interfaces:

```bash
corepack pnpm run build
corepack pnpm exec vite preview --host 0.0.0.0 --port 4173 --strictPort
```

Keep that single preview running while Playwright reuses it:

```bash
E2E_BASE_URL=http://192.168.6.232:4173 corepack pnpm run test:e2e
```

With `E2E_BASE_URL` set, Playwright omits its configured web server and does not
start a second development listener.

## Licensing

This repository is GPL-3.0-or-later. It also initializes HyperFormula with its
GPLv3 license key. Confirm that the host application's distribution model is
GPL-compatible before embedding this package. A commercial HyperFormula license
does not by itself relicense the surrounding JS-Spreadsheet GPL code; obtain
appropriate licensing advice or separate permission if the host cannot comply.

There is currently no vanilla JavaScript, Web Component, Vue, Svelte, or Angular
UI adapter. Those applications can use the DOM-free core directly, but the
shipped `Spreadsheet`, `DataTable`, and workbook table view require React.
