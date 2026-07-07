---
name: running-the-app
description: Launch and drive the JS-Spreadsheet dev app (Vite + React + HyperFormula). Use when running, starting, screenshotting, or end-to-end verifying this repo's spreadsheet UI, or when pnpm/node commands fail with version errors.
---

# Running JS-Spreadsheet

## Environment (required — system Node 20 is too old)

The repo pins `pnpm@11.7.0`, which needs Node ≥22.13. A user-local Node 22
lives at `~/.local/node22`. Prepend it to PATH for **every** command
(shell state does not persist between Bash calls):

```bash
export PATH="$HOME/.local/node22/bin:$PATH"
```

Using system Node fails with `ERR_UNKNOWN_BUILTIN_MODULE: node:sqlite`.
If `~/.local/node22` is missing, download the latest v22 linux-x64 tarball
from nodejs.org, verify its SHASUMS256 entry, and extract with
`--strip-components=1` into `~/.local/node22`.

## Launch

```bash
export PATH="$HOME/.local/node22/bin:$PATH"
cd ~/JS-Spreadsheet
COREPACK_ENABLE_DOWNLOAD_PROMPT=0 corepack pnpm install   # first time only
corepack pnpm dev        # → http://127.0.0.1:5173/  (run in background)
```

Wait for `Local:` in the output, then smoke-test: `curl -fsS http://127.0.0.1:5173/`.

Tests: `corepack pnpm test` (vitest), `corepack pnpm test:e2e` (Playwright).

## Driving the UI with Playwright

Run driver scripts **from inside the repo** so `@playwright/test` resolves.
The pinned Playwright build may lack its downloaded browser — fall back to
system Chrome:

```js
import { chromium } from '@playwright/test';
let browser;
try { browser = await chromium.launch(); }
catch { browser = await chromium.launch({ channel: 'chrome' }); }
```

### Cell selectors (Grid.tsx)

- Empty cell: `[aria-label="B2"]`
- Filled cell: aria-label becomes `"B2 <value>"` — match both:
  `page.locator('[aria-label="B2"], [aria-label^="B2 "]')`

### Entering data — autocomplete caveat

The formula autocomplete popup **swallows Enter** and replaces input like
`=1+2` with `=SUM(`. To commit:

- Plain numbers/text: type, then `Enter` (safe — no popup on digits).
- Formulas: type, then **click another cell** to commit (or `Escape` may
  dismiss the popup first, but click-away is the verified-reliable path).

Verified working example: type `2`→A1, `3`→B1, `=A1+B1`→C1 (click-away),
then read back C1 → displays `5`; status bar shows "Saved".
