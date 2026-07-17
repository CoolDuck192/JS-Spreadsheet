#!/usr/bin/env node
/**
 * Codemod: namespace every generic workbook CSS class as `js-spreadsheet-<name>`.
 *
 * Why: the workbook surface used ~200 generic class names (`app-shell`, `toolbar`,
 * `cell`, ...). Host applications embedding the package can (and did) ship CSS rules
 * for the same generic names, crushing the component. `@scope` protects the package's
 * own rules from leaking out, but cannot stop HOST css from matching generic names on
 * the package's DOM — so the names themselves must be namespaced.
 *
 * What it rewrites (and nothing else):
 *   - class selectors in src/App.css + src/standalone.css (comments left untouched);
 *   - className-bearing string literals / template chunks in src/App.tsx and
 *     src/components/*.tsx (JSX `className=`/`*ClassName=` attributes, `const
 *     <x>ClassName =` initializers, `className:` object properties) — every
 *     whitespace-separated token, word-boundary exact;
 *   - test selectors in unit tests and e2e specs: `.name` inside string literals,
 *     `\.name` inside regex literals, string/regex arguments to toHaveClass /
 *     classList.contains|add|remove|toggle, and className props (exact rename-map
 *     matches only, so host-supplied classNames in tests stay untouched).
 *
 * It does NOT rewrite comments, aria-labels, titles, roles, test names, or any other
 * prose — a word like "toolbar" outside a class context must not change.
 *
 * Idempotent: a second run performs zero replacements.
 */

import { readFileSync, writeFileSync, readdirSync } from "node:fs";
import { join } from "node:path";
import { fileURLToPath } from "node:url";

const ROOT = fileURLToPath(new URL("..", import.meta.url));
const PREFIX = "js-spreadsheet-";

function walk(dir) {
  const out = [];
  for (const entry of readdirSync(join(ROOT, dir), { withFileTypes: true })) {
    if (entry.name === "node_modules" || entry.name.startsWith(".")) continue;
    const rel = `${dir}/${entry.name}`;
    if (entry.isDirectory()) out.push(...walk(rel));
    else out.push(rel);
  }
  return out;
}

const CSS_FILES = ["src/App.css", "src/standalone.css"];
const SOURCE_FILES = [
  "src/App.tsx",
  ...walk("src/components").filter((f) => f.endsWith(".tsx") && !f.endsWith(".test.tsx"))
];
const TEST_FILES = [...walk("src"), ...walk("tests")].filter(
  (f) => /\.test\.tsx?$/.test(f) || /\.spec\.tsx?$/.test(f)
);

const read = (file) => readFileSync(join(ROOT, file), "utf8");

// ---------------------------------------------------------------------------
// JS/TSX scanner: segments the file so rewrites only ever touch real string
// literals / template chunks, never comments, regex bodies, or code.
// ---------------------------------------------------------------------------

/**
 * @returns {{ segments: Array<{type: "string"|"template-chunk"|"regex"|"line-comment"|"block-comment",
 *   start: number, end: number, contentStart: number, contentEnd: number,
 *   prevInterp?: boolean, nextInterp?: boolean}>, masked: string }}
 * `masked` mirrors the input with every non-code character blanked to a space,
 * so structural regexes (attribute names, braces) can run without being fooled
 * by string/comment/regex contents. Newlines are preserved.
 */
function scanJs(text) {
  const segments = [];
  const masked = text.split("");
  const blank = (from, to) => {
    for (let k = from; k < to; k++) if (masked[k] !== "\n") masked[k] = " ";
  };
  let prevSignificant = "";

  const isRegexPosition = () =>
    prevSignificant === "" ||
    "(,=:[!&|?{};+-*%~^<>".includes(prevSignificant.slice(-1)) ||
    /^(?:return|typeof|case|in|of|new|delete|void|do|else|yield|await)$/.test(prevSignificant);

  /** Scans code from `i`; when stopAtBrace, returns at the unmatched `}`. */
  function scanCode(i, stopAtBrace) {
    let depth = 0;
    while (i < text.length) {
      const ch = text[i];
      const next = text[i + 1];
      if (ch === "/" && next === "/") {
        const nl = text.indexOf("\n", i);
        const stop = nl === -1 ? text.length : nl;
        segments.push({ type: "line-comment", start: i, end: stop, contentStart: i + 2, contentEnd: stop });
        blank(i, stop);
        i = stop;
        continue;
      }
      if (ch === "/" && next === "*") {
        const close = text.indexOf("*/", i + 2);
        const stop = close === -1 ? text.length : close + 2;
        segments.push({ type: "block-comment", start: i, end: stop, contentStart: i + 2, contentEnd: stop - 2 });
        blank(i, stop);
        i = stop;
        continue;
      }
      if (ch === '"' || ch === "'") {
        const start = i;
        i++;
        while (i < text.length && text[i] !== ch && text[i] !== "\n") {
          if (text[i] === "\\") i++;
          i++;
        }
        i++;
        segments.push({ type: "string", start, end: i, contentStart: start + 1, contentEnd: i - 1 });
        blank(start + 1, i - 1);
        prevSignificant = ch;
        continue;
      }
      if (ch === "`") {
        i = scanTemplate(i);
        prevSignificant = "`";
        continue;
      }
      if (ch === "/" && isRegexPosition()) {
        const scanned = scanRegex(i);
        if (scanned !== null) {
          i = scanned;
          prevSignificant = "/";
          continue;
        }
      }
      if (stopAtBrace) {
        if (ch === "{") depth++;
        else if (ch === "}") {
          if (depth === 0) return i;
          depth--;
        }
      }
      if (!/\s/.test(ch)) {
        prevSignificant = /[A-Za-z0-9_$]/.test(ch)
          ? (/[A-Za-z0-9_$]$/.test(prevSignificant) ? prevSignificant + ch : ch)
          : ch;
      }
      i++;
    }
    return i;
  }

  /** text[i] === "`"; returns index just past the closing backtick. */
  function scanTemplate(i) {
    i++;
    let chunkStart = i;
    let prevInterp = false;
    while (i < text.length) {
      if (text[i] === "\\") {
        i += 2;
        continue;
      }
      if (text[i] === "`") {
        segments.push({ type: "template-chunk", start: chunkStart, end: i, contentStart: chunkStart, contentEnd: i, prevInterp, nextInterp: false });
        blank(chunkStart, i);
        return i + 1;
      }
      if (text[i] === "$" && text[i + 1] === "{") {
        segments.push({ type: "template-chunk", start: chunkStart, end: i, contentStart: chunkStart, contentEnd: i, prevInterp, nextInterp: true });
        blank(chunkStart, i);
        i = scanCode(i + 2, true); // returns at the matching "}"
        i++;
        chunkStart = i;
        prevInterp = true;
        continue;
      }
      i++;
    }
    segments.push({ type: "template-chunk", start: chunkStart, end: text.length, contentStart: chunkStart, contentEnd: text.length, prevInterp, nextInterp: false });
    blank(chunkStart, text.length);
    return text.length;
  }

  /** text[i] === "/"; returns index past the regex (with flags) or null if not one. */
  function scanRegex(i) {
    const start = i;
    i++;
    let inClass = false;
    while (i < text.length) {
      const c = text[i];
      if (c === "\\") {
        i += 2;
        continue;
      }
      if (c === "\n") return null;
      if (c === "[") inClass = true;
      else if (c === "]") inClass = false;
      else if (c === "/" && !inClass) break;
      i++;
    }
    if (text[i] !== "/") return null;
    const contentStart = start + 1;
    const contentEnd = i;
    i++;
    while (i < text.length && /[a-z]/i.test(text[i])) i++;
    segments.push({ type: "regex", start, end: i, contentStart, contentEnd });
    blank(contentStart, contentEnd);
    return i;
  }

  scanCode(0, false);
  return { segments, masked: masked.join("") };
}

// ---------------------------------------------------------------------------
// CSS name extraction + selector rewriting
// ---------------------------------------------------------------------------

function splitCssComments(css) {
  const parts = [];
  let i = 0;
  while (i < css.length) {
    const open = css.indexOf("/*", i);
    if (open === -1) {
      parts.push({ text: css.slice(i), isComment: false, start: i });
      break;
    }
    if (open > i) parts.push({ text: css.slice(i, open), isComment: false, start: i });
    let close = css.indexOf("*/", open + 2);
    close = close === -1 ? css.length : close + 2;
    parts.push({ text: css.slice(open, close), isComment: true, start: open });
    i = close;
  }
  return parts;
}

function stripCssStrings(text) {
  return text.replace(/"(?:[^"\\]|\\.)*"|'(?:[^'\\]|\\.)*'/g, (m) => '"'.padEnd(m.length - 1, " ") + '"');
}

function extractCssClassNames(css) {
  const names = new Set();
  for (const part of splitCssComments(css)) {
    if (part.isComment) continue;
    for (const match of stripCssStrings(part.text).matchAll(/\.([A-Za-z_][A-Za-z0-9_-]*)/g)) {
      names.add(match[1]);
    }
  }
  return names;
}

// ---------------------------------------------------------------------------
// className-context detection (shared by source + test passes)
// ---------------------------------------------------------------------------

const CLASS_TOKEN = /^[A-Za-z][A-Za-z0-9_-]*$/;
const HEAD_PARTIAL = /^[A-Za-z][A-Za-z0-9_-]*-$/; // e.g. `status-persistence-` right before ${...}

/** Find [start, end) expression regions attached to className-ish attributes/vars. */
function findClassNameRegions(text, masked) {
  const regions = [];
  const attrRe = /\b([A-Za-z0-9_$]*[cC]lassName)\s*(=|:)\s*/g;
  let m;
  while ((m = attrRe.exec(masked)) !== null) {
    const at = m.index + m[0].length;
    const ch = text[at];
    if (m[2] === ":") {
      // object property: only process an immediate string/template literal or an
      // array/parenthesized expression, so TS type annotations (`className: string`)
      // are never touched (identifiers inside a type never look like string literals).
      if (ch === '"' || ch === "'" || ch === "`") regions.push([at, findLiteralEnd(text, at)]);
      else if (ch === "[" || ch === "(") regions.push([at, findBalanced(masked, at, ch)]);
      continue;
    }
    if (masked[at] === "=") continue; // `className === "x"` comparison, not an assignment
    if (ch === '"' || ch === "'" || ch === "`") {
      regions.push([at, findLiteralEnd(text, at)]);
    } else if (ch === "{" || ch === "[" || ch === "(") {
      regions.push([at, findBalanced(masked, at, ch)]);
    } else {
      const before = masked.slice(0, m.index);
      if (/\b(?:const|let|var)\s+$/.test(before)) regions.push([at, findStatementEnd(masked, at)]);
    }
  }
  return regions;
}

function findLiteralEnd(text, start) {
  const quote = text[start];
  let i = start + 1;
  while (i < text.length) {
    if (text[i] === "\\") i += 2;
    else if (text[i] === quote) return i + 1;
    else if (quote === "`" && text[i] === "$" && text[i + 1] === "{") {
      let depth = 1;
      i += 2;
      while (i < text.length && depth > 0) {
        if (text[i] === "{") depth++;
        else if (text[i] === "}") depth--;
        i++;
      }
    } else i++;
  }
  return i;
}

function findBalanced(masked, start, open) {
  const close = { "{": "}", "[": "]", "(": ")" }[open];
  let depth = 0;
  for (let i = start; i < masked.length; i++) {
    if (masked[i] === open) depth++;
    else if (masked[i] === close) {
      depth--;
      if (depth === 0) return i + 1;
    }
  }
  return masked.length;
}

function findStatementEnd(masked, start) {
  let depth = 0;
  for (let i = start; i < masked.length; i++) {
    const ch = masked[i];
    if ("([{".includes(ch)) depth++;
    else if (")]}".includes(ch)) depth--;
    else if (ch === ";" && depth <= 0) return i;
  }
  return masked.length;
}

/** True when the literal is a comparison operand (`=== "table"`), not a class token. */
function isComparisonOperand(masked, seg) {
  const before = masked.slice(0, seg.start);
  const after = masked.slice(seg.end);
  return /(?:[=!]==?|\bcase)\s*$/.test(before) || /^\s*[=!]==?/.test(after);
}

// ---------------------------------------------------------------------------
// Rewrite passes
// ---------------------------------------------------------------------------

const report = {
  renamed: new Set(),
  filesTouched: new Set(),
  replacements: 0,
  manual: []
};

function applyEdits(text, edits) {
  edits.sort((a, b) => b.start - a.start);
  for (const e of edits) text = text.slice(0, e.start) + e.replacement + text.slice(e.end);
  return text;
}

function rewriteTokens(content, seg, file, line, renameToken) {
  const parts = content.split(/(\s+)/);
  const tokenIndices = parts.map((p, idx) => (p !== "" && !/^\s/.test(p) ? idx : -1)).filter((idx) => idx !== -1);
  if (tokenIndices.length === 0) return content;
  const firstIdx = tokenIndices[0];
  const lastIdx = tokenIndices[tokenIndices.length - 1];
  let changed = false;
  const rewritten = parts.map((part, idx) => {
    if (part === "" || /^\s/.test(part)) return part;
    const touchesPrevInterp = Boolean(seg.prevInterp) && idx === firstIdx && !/^\s/.test(content);
    const touchesNextInterp = Boolean(seg.nextInterp) && idx === lastIdx && !/\s$/.test(content);
    const result = renameToken(part, { touchesPrevInterp, touchesNextInterp, file, line });
    if (result !== part) changed = true;
    return result;
  });
  return changed ? rewritten.join("") : content;
}

/** Source components: every token is package-owned — prefix all unprefixed tokens. */
function sourceRenameToken(token, { touchesPrevInterp, touchesNextInterp, file, line }) {
  if (token.startsWith("js-spreadsheet")) return token;
  if (touchesPrevInterp) {
    report.manual.push(`${file}:${line}: token "${token}" continues an interpolation — fix by hand`);
    return token;
  }
  if (touchesNextInterp && HEAD_PARTIAL.test(token)) {
    report.renamed.add(token.replace(/-$/, ""));
    report.replacements++;
    return PREFIX + token;
  }
  if (CLASS_TOKEN.test(token)) {
    report.renamed.add(token);
    report.replacements++;
    return PREFIX + token;
  }
  report.manual.push(`${file}:${line}: unrecognized className token "${token}" — fix by hand`);
  return token;
}

function processClassNameContexts(file, text, renameToken) {
  const { segments, masked } = scanJs(text);
  const regions = findClassNameRegions(text, masked);
  const edits = [];
  for (const seg of segments) {
    if (seg.type !== "string" && seg.type !== "template-chunk") continue;
    if (!regions.some(([s, e]) => seg.contentStart >= s && seg.contentEnd <= e)) continue;
    if (isComparisonOperand(masked, seg)) continue;
    const content = text.slice(seg.contentStart, seg.contentEnd);
    if (content === "") continue;
    const line = text.slice(0, seg.contentStart).split("\n").length;
    const next = rewriteTokens(content, seg, file, line, renameToken);
    if (next !== content) edits.push({ start: seg.contentStart, end: seg.contentEnd, replacement: next });
  }
  return applyEdits(text, edits);
}

/** Rewrite `.name` occurrences inside string/template literals (selector strings). */
function processDotSelectorsInStrings(file, text, alternation) {
  const { segments } = scanJs(text);
  const edits = [];
  for (const seg of segments) {
    if (seg.type !== "string" && seg.type !== "template-chunk") continue;
    const content = text.slice(seg.contentStart, seg.contentEnd);
    let changed = false;
    const next = content.replace(alternation, (full, name) => {
      changed = true;
      report.replacements++;
      report.renamed.add(name);
      return `.${PREFIX}${name}`;
    });
    if (changed) edits.push({ start: seg.contentStart, end: seg.contentEnd, replacement: next });
  }
  return applyEdits(text, edits);
}

/** Rewrite `\.name` occurrences inside regex literals (test regexes over CSS/DOM). */
function processEscapedDotsInRegexes(file, text, names) {
  const alternation = makeAlternation(names, String.raw`\\\.`);
  if (!alternation) return text;
  const { segments } = scanJs(text);
  const edits = [];
  for (const seg of segments) {
    if (seg.type !== "regex") continue;
    const content = text.slice(seg.contentStart, seg.contentEnd);
    const next = content.replace(alternation, (full, name) => {
      report.replacements++;
      return `\\.${PREFIX}${name}`;
    });
    if (next !== content) edits.push({ start: seg.contentStart, end: seg.contentEnd, replacement: next });
  }
  return applyEdits(text, edits);
}

/** Rewrite string + regex args of class assertions: toHaveClass / classList.*(...). */
function processClassAssertionCalls(file, text, names) {
  const { segments, masked } = scanJs(text);
  const callRe = /(?:\btoHaveClass|\bclassList\s*\.\s*(?:contains|add|remove|toggle))\s*\(/g;
  const regions = [];
  let m;
  while ((m = callRe.exec(masked)) !== null) {
    const open = m.index + m[0].length - 1;
    regions.push([open, findBalanced(masked, open, "(")]);
  }
  if (regions.length === 0) return text;
  const bareRe = makeAlternation(names, "(?<![A-Za-z0-9_.$-])");
  const edits = [];
  for (const seg of segments) {
    if (!regions.some(([s, e]) => seg.start >= s && seg.end <= e)) continue;
    const content = text.slice(seg.contentStart, seg.contentEnd);
    if (seg.type === "string" || seg.type === "template-chunk") {
      const next = rewriteTokens(content, seg, file, 0, (token) =>
        names.has(token) ? (report.replacements++, report.renamed.add(token), PREFIX + token) : token
      );
      if (next !== content) edits.push({ start: seg.contentStart, end: seg.contentEnd, replacement: next });
    } else if (seg.type === "regex" && bareRe) {
      const next = content.replace(bareRe, (full, name) => {
        report.replacements++;
        report.renamed.add(name);
        return PREFIX + name;
      });
      if (next !== content) edits.push({ start: seg.contentStart, end: seg.contentEnd, replacement: next });
    }
  }
  return applyEdits(text, edits);
}

function makeAlternation(names, prefixPattern) {
  const sorted = [...names].sort((a, b) => b.length - a.length);
  if (sorted.length === 0) return null;
  return new RegExp(`${prefixPattern}(${sorted.join("|")})(?![A-Za-z0-9_-])`, "g");
}

// ---------------------------------------------------------------------------
// Run
// ---------------------------------------------------------------------------

// PASS A: collect renameable class names from the CSS sources.
const cssNames = new Set();
for (const file of CSS_FILES) {
  for (const name of extractCssClassNames(read(file))) {
    if (!name.startsWith("js-spreadsheet")) cssNames.add(name);
  }
}

// PASS B: source components — prefix every token in className-bearing contexts.
for (const file of SOURCE_FILES) {
  const original = read(file);
  const next = processClassNameContexts(file, original, sourceRenameToken);
  if (next !== original) {
    writeFileSync(join(ROOT, file), next);
    report.filesTouched.add(file);
  }
}
const tsxNames = new Set(report.renamed);
const fullNames = new Set([...cssNames, ...tsxNames]);

// PASS C: CSS selector rewriting (comments untouched).
{
  const alternation = makeAlternation(cssNames, String.raw`\.`);
  if (alternation) {
    for (const file of CSS_FILES) {
      const original = read(file);
      let count = 0;
      const next = splitCssComments(original)
        .map((part) =>
          part.isComment
            ? part.text
            : part.text.replace(alternation, (full, name) => {
                count++;
                report.renamed.add(name);
                return `.${PREFIX}${name}`;
              })
        )
        .join("");
      if (next !== original) {
        writeFileSync(join(ROOT, file), next);
        report.filesTouched.add(file);
        report.replacements += count;
      }
    }
  }
}

// PASS D: test files — className props (map-exact), class assertions, selector
// strings, selector regexes.
{
  const dotAlternation = makeAlternation(fullNames, String.raw`\.`);
  for (const file of TEST_FILES) {
    const original = read(file);
    let next = original;
    next = processClassNameContexts(file, next, (token) =>
      fullNames.has(token) ? (report.replacements++, PREFIX + token) : token
    );
    next = processClassAssertionCalls(file, next, fullNames);
    if (dotAlternation) next = processDotSelectorsInStrings(file, next, dotAlternation);
    next = processEscapedDotsInRegexes(file, next, fullNames);
    if (next !== original) {
      writeFileSync(join(ROOT, file), next);
      report.filesTouched.add(file);
    }
  }
}

// PASS E: post-scan — report anything class-shaped the codemod did not rewrite.
{
  const allFiles = [...walk("src"), ...walk("tests")].filter((f) => /\.(tsx?|css)$/.test(f));
  const dotRe = makeAlternation(fullNames, String.raw`\.`);
  for (const file of allFiles) {
    const text = read(file);
    if (file.endsWith(".css")) {
      if (!dotRe) continue;
      for (const part of splitCssComments(text)) {
        if (part.isComment) continue;
        for (const m of stripCssStrings(part.text).matchAll(dotRe)) {
          const line = text.slice(0, part.start + m.index).split("\n").length;
          report.manual.push(`${file}:${line}: leftover selector ".${m[1]}"`);
        }
      }
      continue;
    }
    const { segments } = scanJs(text);
    for (const seg of segments) {
      if (seg.type !== "string" && seg.type !== "template-chunk") continue;
      const content = text.slice(seg.contentStart, seg.contentEnd);
      const line = text.slice(0, seg.contentStart).split("\n").length;
      if (dotRe) {
        for (const m of content.matchAll(dotRe)) {
          report.manual.push(`${file}:${line}: leftover selector reference ".${m[1]}"`);
        }
      }
      const tokens = content.split(/\s+/).filter(Boolean);
      if (
        tokens.length > 0 &&
        tokens.every((t) => /^[a-z][a-z0-9]*(?:--?[a-z0-9]+)*$/.test(t)) &&
        tokens.some((t) => fullNames.has(t))
      ) {
        report.manual.push(`${file}:${line}: class-list-looking literal "${content}" contains un-renamed map name(s)`);
      }
    }
  }
}

// Summary
console.log(`namespace-css: ${report.renamed.size} class names renamed with prefix "${PREFIX}"`);
console.log(`namespace-css: ${report.replacements} total replacements`);
console.log(`namespace-css: ${report.filesTouched.size} files touched`);
for (const file of [...report.filesTouched].sort()) console.log(`  - ${file}`);
if (report.manual.length > 0) {
  console.log(`namespace-css: ${report.manual.length} spot(s) need manual review:`);
  for (const item of report.manual) console.log(`  ! ${item}`);
} else {
  console.log("namespace-css: nothing left for manual review");
}
