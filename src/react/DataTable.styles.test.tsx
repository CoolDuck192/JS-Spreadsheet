import { readFileSync } from "node:fs";
import { resolve } from "node:path";
import { render, screen } from "@testing-library/react";
import { describe, expect, it } from "vitest";
import { DataTable } from "./DataTable";

type Row = { id: string; name: string };

describe("DataTable style boundary", () => {
  it("renders the exact public style root", () => {
    render(<DataTable
      aria-label="Employees"
      rows={[{ id: "1", name: "Ada" }]}
      columns={[{
        id: "name",
        header: "Name",
        accessor: (row: Row) => row.name,
        update: (row, value) => ({ ...row, name: String(value) })
      }]}
      getRowId={(row) => row.id}
    />);
    const root = screen.getByRole("grid", { name: "Employees" }).closest(".js-spreadsheet-root");
    expect(root).toHaveClass("js-spreadsheet-data-table");
    expect(root).toHaveAttribute("data-js-spreadsheet-root", "data-table");
  });

  it("contains no unscoped element or universal selectors", () => {
    const css = readFileSync(resolve(process.cwd(), "src/styles/data-table.css"), "utf8");
    const withoutComments = css.replace(/\/\*[\s\S]*?\*\//g, "");
    expect(withoutComments).not.toMatch(
      /(?:^|[{}])\s*(?:body|html|:root|button|input|select|textarea|\*)(?=\s|,|\{|\.|:|#|\[)/m
    );
  });
});
