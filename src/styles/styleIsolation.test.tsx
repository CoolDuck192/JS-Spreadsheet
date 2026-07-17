import { existsSync, readFileSync } from "node:fs";
import { render, screen } from "@testing-library/react";
import { describe, expect, it } from "vitest";
import App from "../App";
import { DataTable } from "../react/DataTable";
import type { ColumnDef } from "../react/tableTypes";
import "../entry/styles.css";

/**
 * Every class selector in a CSS source, namespaced or not.
 *
 * Comments and quoted strings (font names, `content:` values, attribute-selector
 * values) are stripped first, so the only remaining `.foo` tokens are real class
 * selectors. Nothing else needs exempting: pseudo-classes/elements (`:hover`,
 * `::before`) and keyframe names never start with a dot, and numeric values like
 * `.5rem` fail the leading `[A-Za-z_]` requirement.
 */
function extractClassSelectors(css: string): string[] {
  const withoutComments = css.replace(/\/\*[\s\S]*?\*\//g, " ");
  const withoutStrings = withoutComments.replace(/"(?:[^"\\]|\\.)*"|'(?:[^'\\]|\\.)*'/g, '""');
  const names = [...withoutStrings.matchAll(/\.([A-Za-z_][A-Za-z0-9_-]*)/g)].map((match) => match[1]);
  return [...new Set(names)].sort();
}

type Person = { id: string; name: string };
const rows: readonly Person[] = [{ id: "1", name: "Ada" }];
const columns: readonly ColumnDef<Person>[] = [{
  id: "name",
  header: "Name",
  accessor: (row) => row.name,
  update: (row, value) => ({ ...row, name: String(value) })
}];

describe("embedded style isolation", () => {
  it("keeps standalone document sizing out of embedded and published styles", () => {
    const standalonePath = "src/standalone.css";
    expect(existsSync(standalonePath)).toBe(true);
    if (!existsSync(standalonePath)) {
      return;
    }

    expect(readFileSync(standalonePath, "utf8").trim()).toBe(`html,
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

.js-spreadsheet-standalone .js-spreadsheet-app-shell {
  min-height: 0;
}`);

    const publicCss = readFileSync("src/entry/styles.css", "utf8");
    expect(publicCss).not.toContain("standalone.css");
    expect(publicCss).not.toMatch(/^(?:\s*)(?:#root|body|html)\s*[{,]/m);
  });

  it("lets only the standalone app shell shrink below the embedded minimum", () => {
    const standaloneCss = readFileSync("src/standalone.css", "utf8");
    const workbookCss = readFileSync("src/App.css", "utf8");

    expect(standaloneCss).toMatch(
      /\.js-spreadsheet-standalone\s+\.js-spreadsheet-app-shell\s*\{[^}]*min-height:\s*0;[^}]*\}/
    );
    expect(workbookCss).toMatch(
      /\.js-spreadsheet-root\.js-spreadsheet-workbook\s*\{[^}]*min-height:\s*420px;[^}]*height:\s*100%;[^}]*\}/
    );
    expect(workbookCss).toMatch(
      /@scope \(\.js-spreadsheet-root\.js-spreadsheet-workbook\)[\s\S]*?\.js-spreadsheet-app-shell\s*\{[^}]*min-height:\s*420px;[^}]*height:\s*100%;[^}]*\}/
    );
  });

  it("does not decorate host controls and keeps table instances separate", () => {
    render(
      <>
        <button data-testid="host">Host</button>
        <DataTable className="theme-a" rows={rows} columns={columns} getRowId={(row) => row.id} />
        <DataTable className="theme-b" rows={rows} columns={columns} getRowId={(row) => row.id} />
      </>
    );

    expect(screen.getByTestId("host").className).toBe("");
    expect(document.querySelector(".theme-a")).not.toBe(document.querySelector(".theme-b"));
    expect(document.querySelectorAll(".js-spreadsheet-root.js-spreadsheet-data-table")).toHaveLength(2);
  });

  it("keeps public variables namespaced and legacy workbook rules inside a component scope", () => {
    const files = ["src/App.css", "src/styles/data-table.css", "src/styles/tokens.css"];
    const css = files.map((file) => readFileSync(file, "utf8")).join("\n");
    const variables = [
      ...css.matchAll(/var\((--[a-zA-Z0-9-]+)/g),
      ...css.matchAll(/^\s*(--[a-zA-Z0-9-]+)\s*:/gm)
    ].map((match) => match[1]);
    expect(variables.every((name) => name.startsWith("--js-spreadsheet-"))).toBe(true);

    const workbookCss = readFileSync("src/App.css", "utf8");
    expect(workbookCss).toContain("@scope (.js-spreadsheet-root.js-spreadsheet-workbook)");
    expect(workbookCss).not.toMatch(/^(?:\s*)(?::root|body|html|button|input|select|\*)\s*[{,]/m);
    expect(workbookCss).toMatch(
      /@media \(max-width: 720px\)[\s\S]*?\.js-spreadsheet-spreadsheet-surface\s*\{[^}]*height:\s*100%;[^}]*\}/
    );
  });

  it("namespaces every class selector in the workbook CSS sources as js-spreadsheet-*", () => {
    // Host applications ship CSS for generic names like `.app-shell` or `.toolbar`;
    // any unprefixed class on this package's DOM is a real production collision
    // (a host `.app-shell` rule crushed the embedded grid to 0px). `@scope` cannot
    // prevent host rules from matching generic names, so the names themselves must
    // be namespaced.
    const sources = ["src/App.css", "src/standalone.css", "src/styles/data-table.css", "src/styles/tokens.css"];
    for (const file of sources) {
      const offenders = extractClassSelectors(readFileSync(file, "utf8")).filter(
        (name) => !name.startsWith("js-spreadsheet-")
      );
      expect(offenders, `${file} must only use js-spreadsheet-* class selectors`).toEqual([]);
    }
  });

  it("ships only js-spreadsheet-* class selectors in the published stylesheet", () => {
    // The published file is compiled from the sources asserted above, but this
    // guards the bundling path too (App.css reaches dist/lib/styles.css through
    // src/styles/spreadsheet.css and the vite lib build). The file only exists
    // after `pnpm build:lib`; when absent, the source-level gate still applies.
    const published = "dist/lib/styles.css";
    if (!existsSync(published)) {
      return;
    }
    const offenders = extractClassSelectors(readFileSync(published, "utf8")).filter(
      (name) => !name.startsWith("js-spreadsheet-")
    );
    expect(offenders, `${published} must only use js-spreadsheet-* class selectors`).toEqual([]);
  });

  it("renders the workbook with only js-spreadsheet-* class tokens", () => {
    localStorage.clear();
    try {
      const { container } = render(<App />);
      const elements = [container, ...Array.from(container.querySelectorAll("*"))];
      expect(elements.length).toBeGreaterThan(100);

      const offenders = new Set<string>();
      for (const element of elements) {
        for (const token of Array.from(element.classList)) {
          // lucide-react owns the `lucide` / `lucide-*` classes on its icon SVGs;
          // they are third-party, unstyled by this package, and not ours to rename.
          if (token === "lucide" || token.startsWith("lucide-")) {
            continue;
          }
          if (!token.startsWith("js-spreadsheet-")) {
            offenders.add(token);
          }
        }
      }
      expect([...offenders].sort(), "workbook DOM must only carry js-spreadsheet-* class tokens").toEqual([]);
    } finally {
      localStorage.clear();
    }
  });

  it("keeps visually hidden descriptions compatible with modern and fallback clipping", () => {
    const workbookCss = readFileSync("src/App.css", "utf8");
    const tableCss = readFileSync("src/styles/data-table.css", "utf8");

    expect(workbookCss).toMatch(/\.js-spreadsheet-visually-hidden\s*\{(?=[^}]*clip-path:\s*inset\(50%\);)(?=[^}]*clip:\s*rect\()[^}]*\}/s);
    expect(tableCss).toMatch(/__visually-hidden\s*\{(?=[^}]*clip-path:\s*inset\(50%\);)(?=[^}]*clip:\s*rect\()[^}]*\}/s);
    expect(workbookCss).not.toMatch(/\b(?:currentColor|optimizeLegibility)\b/);
  });

  it("gives the Quick Tools toolbar and drawer full-size neutral scrollbars", () => {
    const css = readFileSync("src/styles/data-table.css", "utf8");
    const baseCss = css.slice(0, css.indexOf("@container"));

    expect(css).not.toMatch(/__toolbar\s*\{[^}]*scrollbar-width:\s*thin;/s);
    expect(baseCss).toMatch(/__toolbar\s*\{(?=[^}]*overflow-x:\s*auto;)(?=[^}]*overflow-y:\s*hidden;)[^}]*\}/s);
    expect(css).toMatch(/__toolbar::\-webkit-scrollbar\s*\{[^}]*height:\s*12px;/s);
    expect(css).toMatch(/__toolbar::\-webkit-scrollbar-thumb\s*\{[^}]*min-width:\s*40px;/s);
    expect(css).toMatch(/__quick-tools::\-webkit-scrollbar\s*\{[^}]*width:\s*12px;/s);
    expect(css).toMatch(/__quick-tools::\-webkit-scrollbar-thumb\s*\{[^}]*min-height:\s*40px;/s);
    expect(baseCss).not.toMatch(/__quick-tools\s*\{[^}]*position:\s*fixed;/s);
    expect(baseCss).toMatch(
      /__quick-tools\s*\{(?=[^}]*position:\s*absolute;)(?=[^}]*max-height:\s*min\(420px, calc\(100dvh - 16px\), calc\(100% - 16px\)\);)(?=[^}]*overflow-x:\s*hidden;)(?=[^}]*overflow-y:\s*auto;)(?=[^}]*pointer-events:\s*auto;)[^}]*\}/s
    );
  });

  it("keeps the clipped table root while sizing top-layer column menus to the viewport", () => {
    const css = readFileSync("src/styles/data-table.css", "utf8");
    const rootRule = css.match(/\.js-spreadsheet-root\.js-spreadsheet-data-table\s*\{[^}]*\}/s)?.[0] ?? "";
    const menuRule = css.match(/\.js-spreadsheet-data-table__column-menu\s*\{[^}]*\}/s)?.[0] ?? "";
    const openMenuRule = css.match(/\.js-spreadsheet-data-table__column-menu:popover-open\s*\{[^}]*\}/s)?.[0] ?? "";
    const treeContentRule = css.match(/\.js-spreadsheet-data-table__cell-content--tree\s*\{[^}]*\}/s)?.[0] ?? "";
    const treeToggleRule = css.match(/\.js-spreadsheet-data-table__tree-toggle\s*\{[^}]*\}/s)?.[0] ?? "";

    expect(rootRule).toMatch(/overflow:\s*hidden;/);
    expect(menuRule).toMatch(/position:\s*fixed;/);
    expect(menuRule).toMatch(/margin:\s*0;/);
    expect(menuRule).toMatch(/width:\s*min\(248px, calc\(100vw - 16px\)\);/);
    expect(menuRule).toMatch(/max-height:\s*min\(520px, calc\(100dvh - 16px\)\);/);
    expect(menuRule).not.toMatch(/\binset(?:-|:)/);
    expect(menuRule).not.toMatch(/\b(?:cqi|cqb)\b/);
    expect(menuRule).not.toMatch(/\bdisplay\s*:/);
    expect(openMenuRule).toMatch(/display:\s*grid;/);
    expect(css).not.toMatch(/__cell-content > button\s*\{/);
    expect(treeContentRule).toMatch(/padding-block:\s*1px;/);
    expect(treeToggleRule).toMatch(/(?:width|min-width):\s*30px;/);
    expect(treeToggleRule).toMatch(/min-height:\s*30px;/);
  });
});
