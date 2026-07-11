# Getting Started

## Prerequisites

| Requirement | Why |
| --- | --- |
| Git | clone the repo |
| **Node.js ≥ 22.13** | the pinned pnpm 11 uses `node:sqlite`, which does not exist on Node 20 — you'll see `ERR_UNKNOWN_BUILTIN_MODULE: node:sqlite` |
| corepack (ships with Node) | provisions the exact pinned pnpm automatically |

No global pnpm install is needed. If your system Node is older than 22.13, install a user-local Node 22+ (nvm, fnm, or a tarball extracted to e.g. `~/.local/node22` with its `bin` prepended to `PATH`).

## Setup

```bash
git clone https://github.com/CoolDuck192/JS-Spreadsheet.git
cd JS-Spreadsheet
corepack enable
corepack pnpm install --frozen-lockfile
```

## Run

```bash
corepack pnpm run dev
```

Vite binds all interfaces. Open the `Network` URL it reports, using the
machine's LAN address.

## Production build

```bash
corepack pnpm run build
corepack pnpm exec vite preview --host 0.0.0.0 --port 4173 --strictPort
```

**Use the production preview when working with large datasets.** The dev server runs React's development build, whose instrumentation makes 100k-row sheets feel dramatically slower than production (see [performance.md](performance.md)).

## Tests

```bash
corepack pnpm test                     # unit tests (Vitest, jsdom)
corepack pnpm exec playwright install  # once, downloads browsers
```

To run browser tests against the one production preview above, leave that
preview running and use:

```bash
E2E_BASE_URL=http://192.168.6.232:4173 corepack pnpm run test:e2e
```

When `E2E_BASE_URL` is set, Playwright does not start its configured development
server.

The 100k-row engine benchmark lives at `src/lib/formulaEngine.perf.test.ts`. On slow or contended CI runners, set `SKIP_PERF_ASSERT=1` to keep its correctness checks but skip the wall-clock thresholds.

## First five minutes in the app

- Type values or formulas (`=SUM(A1:A5)`) into cells; Enter commits, Tab accepts an autocomplete suggestion, ↑/↓ navigate suggestions.
- **Drag an `.xlsx` or `.csv` file anywhere onto the window** to import it.
- Select a data range → Insert → Pivot Table; after creating it, **double-click any pivot value to drill into its source rows**.
- File → Import Google Sheet performs a one-time, read-only workbook replacement. A fresh clone opens the setup dialog when configuration is absent; see [google-sheets-connector.md](google-sheets-connector.md).
- Work autosaves to browser localStorage (up to ~100k cells; beyond that, export to XLSX).
