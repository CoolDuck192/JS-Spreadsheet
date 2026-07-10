import { render, screen } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import { describe, expect, it, vi } from "vitest";
import type { StructuredTable } from "../types";
import { SpreadsheetTableTab, type SpreadsheetTableTabProps } from "./SpreadsheetTableTab";

function createTable(overrides: Partial<StructuredTable> = {}): StructuredTable {
  return {
    id: "table-sales",
    name: "SalesTable",
    sheetId: "sheet-1",
    range: { start: { row: 0, column: 0 }, end: { row: 3, column: 1 } },
    headerRow: true,
    totalsRow: false,
    columns: [
      { id: "column-region", name: "Region", sheetColumn: 0 },
      { id: "column-sales", name: "Sales", sheetColumn: 1, totalsFunction: "sum" }
    ],
    rowIds: ["row-west", "row-east", "row-north"],
    style: { theme: "TableStyleLight1", showRowStripes: true },
    ...overrides
  };
}

function createProps(overrides: Partial<SpreadsheetTableTabProps> = {}): SpreadsheetTableTabProps {
  return {
    table: createTable(),
    activeColumnId: "column-sales",
    issues: [],
    onRename: vi.fn(),
    onResize: vi.fn(),
    onHeaderRow: vi.fn(),
    onTotalsRow: vi.fn(),
    onTotalsFunction: vi.fn(),
    onStyle: vi.fn(),
    onKeyColumn: vi.fn(),
    onCalculatedColumn: vi.fn(),
    onFilter: vi.fn(),
    onExport: vi.fn(),
    onConvertToRange: vi.fn(),
    onOpenTableView: vi.fn(),
    ...overrides
  };
}

describe("SpreadsheetTableTab", () => {
  it("commits name and range drafts once each on Enter", async () => {
    const user = userEvent.setup();
    const props = createProps();
    render(<SpreadsheetTableTab {...props} />);

    const name = screen.getByLabelText("Table name");
    await user.clear(name);
    await user.type(name, "RevenueTable{Enter}");

    const range = screen.getByLabelText("Table range");
    await user.clear(range);
    await user.type(range, "B2:D8{Enter}");

    expect(props.onRename).toHaveBeenCalledOnce();
    expect(props.onRename).toHaveBeenCalledWith("RevenueTable");
    expect(props.onResize).toHaveBeenCalledOnce();
    expect(props.onResize).toHaveBeenCalledWith({
      start: { row: 1, column: 1 },
      end: { row: 7, column: 3 }
    });
  });

  it("retains an invalid range draft, focus, and an assertive field issue", async () => {
    const user = userEvent.setup();
    const props = createProps();
    render(<SpreadsheetTableTab {...props} />);

    const range = screen.getByLabelText("Table range");
    await user.clear(range);
    await user.type(range, "A1:B{Enter}");

    expect(props.onResize).not.toHaveBeenCalled();
    expect(range).toHaveValue("A1:B");
    expect(range).toHaveFocus();
    expect(range).toHaveAttribute("aria-describedby", expect.stringContaining("table-range-error"));
    expect(screen.getByRole("alert")).toHaveTextContent("Enter a valid cell range");
  });

  it("associates rejected command issues with the affected field", () => {
    const props = createProps({
      issues: [{ code: "TABLE_NAME_INVALID", message: "Table names cannot contain spaces", sheetId: "sheet-1" }]
    });
    render(<SpreadsheetTableTab {...props} />);

    expect(screen.getByLabelText("Table name")).toHaveAttribute(
      "aria-describedby",
      expect.stringContaining("table-name-error")
    );
    expect(screen.getByRole("alert")).toHaveTextContent("Table names cannot contain spaces");
  });

  it("exposes all table metadata and table actions through accessible controls", async () => {
    const user = userEvent.setup();
    const props = createProps();
    render(<SpreadsheetTableTab {...props} />);

    await user.click(screen.getByRole("checkbox", { name: "Header row" }));
    await user.click(screen.getByRole("checkbox", { name: "Totals row" }));
    await user.selectOptions(screen.getByLabelText("Totals column"), "column-region");
    await user.selectOptions(screen.getByLabelText("Totals function"), "count");
    await user.selectOptions(screen.getByLabelText("Table style"), "TableStyleMedium2");
    await user.selectOptions(screen.getByLabelText("Key column"), "column-region");

    expect(props.onHeaderRow).toHaveBeenCalledWith(false);
    expect(props.onTotalsRow).toHaveBeenCalledWith(true);
    expect(props.onTotalsFunction).toHaveBeenCalledWith("column-region", "count");
    expect(props.onStyle).toHaveBeenCalledWith(expect.objectContaining({ theme: "TableStyleMedium2" }));
    expect(props.onKeyColumn).toHaveBeenCalledWith("column-region");

    await user.click(screen.getByRole("button", { name: "Calculated column" }));
    expect(screen.getByRole("button", { name: "Apply calculated column" })).toBeInTheDocument();
    await user.click(screen.getByRole("button", { name: "Filter table" }));
    expect(screen.getByRole("button", { name: "Apply table filter" })).toBeInTheDocument();

    const exportButton = screen.getByRole("button", { name: "Export table" });
    expect(exportButton).toBeDisabled();
    expect(exportButton).toHaveAccessibleDescription("Export is not supported by this table");

    await user.click(screen.getByRole("button", { name: "Convert to range" }));
    await user.click(screen.getByRole("button", { name: "Open table view" }));
    expect(props.onConvertToRange).toHaveBeenCalledOnce();
    expect(props.onOpenTableView).toHaveBeenCalledOnce();
    expect(props.onExport).not.toHaveBeenCalled();
  });

  it("resets local drafts when committed metadata or the active table changes", async () => {
    const user = userEvent.setup();
    const props = createProps();
    const { rerender } = render(<SpreadsheetTableTab {...props} />);

    const name = screen.getByLabelText("Table name");
    await user.clear(name);
    await user.type(name, "Uncommitted");

    rerender(<SpreadsheetTableTab {...props} table={createTable({ name: "CommittedName" })} />);
    expect(screen.getByLabelText("Table name")).toHaveValue("CommittedName");

    rerender(
      <SpreadsheetTableTab
        {...props}
        table={createTable({ id: "table-costs", name: "CostsTable", range: { start: { row: 2, column: 2 }, end: { row: 4, column: 3 } } })}
      />
    );
    expect(screen.getByLabelText("Table name")).toHaveValue("CostsTable");
    expect(screen.getByLabelText("Table range")).toHaveValue("C3:D5");
  });
});
