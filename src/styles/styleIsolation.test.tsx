import { readFileSync } from "node:fs";
import { render, screen } from "@testing-library/react";
import { describe, expect, it } from "vitest";
import { DataTable } from "../react/DataTable";
import type { ColumnDef } from "../react/tableTypes";
import "../entry/styles.css";

type Person = { id: string; name: string };
const rows: readonly Person[] = [{ id: "1", name: "Ada" }];
const columns: readonly ColumnDef<Person>[] = [{
  id: "name",
  header: "Name",
  accessor: (row) => row.name,
  update: (row, value) => ({ ...row, name: String(value) })
}];

describe("embedded style isolation", () => {
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
  });
});
