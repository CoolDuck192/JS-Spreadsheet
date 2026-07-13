import {
  createRef,
  StrictMode,
  type KeyboardEvent,
  type ReactNode
} from "react";
import { act, fireEvent, render, screen, waitFor, within } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import { afterEach, describe, expect, it, vi } from "vitest";
import { createLocalRecordTableSession } from "../table/local/RecordTableSession";
import { createTableMetadataKey } from "../table/local/tableMetadata";
import type { ChangeContext, RowUpdater, TableDiagnosticEvent, TableViewState } from "../table/core/types";
import { DataTable } from "./DataTable";
import type {
  CellEditorProps,
  ColumnDef,
  DataTableHandle,
  DataTablePresentationProps
} from "./tableTypes";

type Employee = {
  id: string;
  name: string;
  department: string;
  salary: number;
  active: boolean;
  children?: readonly Employee[];
};

const employees: readonly Employee[] = [
  {
    id: "e1",
    name: "Ada",
    department: "Engineering",
    salary: 120,
    active: true,
    children: [{ id: "e1-child", name: "Byron", department: "Engineering", salary: 80, active: true }]
  },
  { id: "e2", name: "Grace", department: "Engineering", salary: 200, active: false },
  { id: "e3", name: "Lin", department: "Finance", salary: 150, active: true }
];

const columns: readonly ColumnDef<Employee>[] = [
  {
    id: "name",
    header: "Name",
    dataType: "text",
    accessor: (row) => row.name,
    update: (row, value) => ({ ...row, name: String(value) }),
    sortable: true,
    filterable: true,
    groupable: true,
    width: 160
  },
  {
    id: "department",
    header: "Department",
    dataType: "text",
    accessor: (row) => row.department,
    update: (row, value) => ({ ...row, department: String(value) }),
    sortable: true,
    filterable: true,
    groupable: true,
    width: 150
  },
  {
    id: "salary",
    header: "Salary",
    dataType: "number",
    accessor: (row) => row.salary,
    update: (row, value) => ({ ...row, salary: Number(value) }),
    validate: ({ parsed }) => Number(parsed) < 0
      ? [{ code: "salary-negative", message: "Salary must be non-negative" }]
      : [],
    sortable: true,
    filterable: true,
    aggregatable: ["sum", "average", "count"],
    width: 110
  },
  {
    id: "active",
    header: "Active",
    dataType: "boolean",
    accessor: (row) => row.active,
    update: (row, value) => ({ ...row, active: Boolean(value) }),
    filterable: true,
    width: 90
  }
];

afterEach(() => {
  vi.restoreAllMocks();
  vi.unstubAllGlobals();
});

describe("DataTable", () => {
  it("renders simple local rows and typed headers", () => {
    renderTable();

    const grid = screen.getByRole("grid", { name: "Employees" });
    expect(grid).toHaveAttribute("data-viewport-kernel", "shared");
    expect(screen.getByRole("columnheader", { name: "Name" })).toBeInTheDocument();
    expect(screen.getByRole("gridcell", { name: "e1 Name" })).toHaveTextContent("Ada");
    expect(screen.getByRole("gridcell", { name: "e1 Name" })).toHaveAttribute("tabindex", "0");
    expect(screen.getByRole("gridcell", { name: "e2 Salary" })).toHaveTextContent("200");
  });

  it("edits a simple controlled row through one onRowsChange updater", async () => {
    const user = userEvent.setup();
    const onRowsChange = vi.fn();
    renderTable({ onRowsChange });

    await editDefaultCell(user, "e1 Name", "Katherine");

    expect(onRowsChange).toHaveBeenCalledTimes(1);
    const updater = onRowsChange.mock.calls[0][0] as (rows: readonly Employee[]) => readonly Employee[];
    expect(updater(employees)[0]).toMatchObject({ id: "e1", name: "Katherine", salary: 120 });
  });

  it("retains uncontrolled simple edits across rerenders until resetKey changes", async () => {
    const user = userEvent.setup();
    const { rerender } = renderTable({ resetKey: 1 });
    await editDefaultCell(user, "e1 Name", "Katherine");
    expect(screen.getByRole("gridcell", { name: "e1 Name" })).toHaveTextContent("Katherine");

    rerender(<DataTable aria-label="Employees" rows={[...employees]} columns={columns} getRowId={getRowId} resetKey={1} />);

    expect(screen.getByRole("gridcell", { name: "e1 Name" })).toHaveTextContent("Katherine");
  });

  it("resets uncontrolled simple rows and history when resetKey changes", async () => {
    const user = userEvent.setup();
    const { rerender } = renderTable({ resetKey: 1 });
    await editDefaultCell(user, "e1 Name", "Katherine");
    expect(screen.getByRole("button", { name: "Undo" })).toBeEnabled();

    const resetRows = [{ ...employees[0], name: "Reset" }, ...employees.slice(1)];
    rerender(<DataTable aria-label="Employees" rows={resetRows} columns={columns} getRowId={getRowId} resetKey={2} />);

    expect(screen.getByRole("gridcell", { name: "e1 Name" })).toHaveTextContent("Reset");
    expect(screen.getByRole("button", { name: "Undo" })).toBeDisabled();
  });

  it("renders a prebuilt session without creating or destroying it", async () => {
    const session = createSession();
    const destroy = vi.spyOn(session, "destroy");
    const { unmount } = render(<DataTable aria-label="Session employees" session={session} />);

    expect(screen.getByRole("grid", { name: "Session employees" })).toBeInTheDocument();
    expect(screen.getByRole("gridcell", { name: "e1 Name" })).toHaveTextContent("Ada");
    unmount();
    await act(async () => Promise.resolve());
    expect(destroy).not.toHaveBeenCalled();
    session.destroy();
  });

  it("uses custom cell, header, and editor renderers", async () => {
    const user = userEvent.setup();
    const customColumns: readonly ColumnDef<Employee>[] = [{
      ...columns[0],
      header: () => <span>Employee name</span>,
      cell: ({ row }) => <strong>{row.original.name.toUpperCase()}</strong>,
      editor: CustomNameEditor
    }];
    render(<DataTable aria-label="Custom employees" rows={employees} columns={customColumns} getRowId={getRowId} />);

    expect(screen.getByText("Employee name")).toBeInTheDocument();
    expect(screen.getByRole("gridcell", { name: "e1 Name" })).toHaveTextContent("ADA");
    await user.dblClick(screen.getByRole("gridcell", { name: "e1 Name" }));
    const editor = screen.getByRole("textbox", { name: "Custom name editor" });
    await user.clear(editor);
    await user.type(editor, "Katherine");
    await user.click(screen.getByRole("button", { name: "Commit custom name" }));
    expect(screen.getByRole("gridcell", { name: "e1 Name" })).toHaveTextContent("KATHERINE");
  });

  it("isolates a throwing custom renderer and reports safe diagnostics without row values", async () => {
    vi.spyOn(console, "error").mockImplementation(() => {});
    const onDiagnostic = vi.fn<(event: TableDiagnosticEvent) => void>();
    const throwingColumns: readonly ColumnDef<Employee>[] = [{
      ...columns[0],
      cell: () => {
        throw new Error("Ada secret row value");
      }
    }];
    render(
      <DataTable
        aria-label="Throwing renderer"
        rows={employees}
        columns={throwingColumns}
        getRowId={getRowId}
        onDiagnostic={onDiagnostic}
      />
    );

    expect(await screen.findAllByText("Cell renderer unavailable")).not.toHaveLength(0);
    expect(onDiagnostic).toHaveBeenCalledWith({
      category: "extension",
      metadata: { columnId: "name", slot: "cell" }
    });
    expect(JSON.stringify(onDiagnostic.mock.calls)).not.toContain("Ada secret row value");
  });

  it("keeps the editor focused when validation rejects a commit", async () => {
    const user = userEvent.setup();
    renderTable();
    await user.dblClick(screen.getByRole("gridcell", { name: "e1 Salary" }));
    const editor = screen.getByRole("textbox", { name: "Edit e1 Salary" });
    await user.clear(editor);
    await user.type(editor, "-5{Enter}");

    expect(editor).toBeInTheDocument();
    expect(editor).toHaveFocus();
    expect(screen.getByRole("alert")).toHaveTextContent("Salary must be non-negative");
    expect(screen.getByRole("gridcell", { name: "e1 Salary" })).toHaveTextContent("120");
  });

  it("replaces a cell from the first printable key", async () => {
    const user = userEvent.setup();
    renderTable();
    const salary = screen.getByRole("gridcell", { name: "e1 Salary" });
    await user.click(salary);

    fireEvent.keyDown(salary, { key: "4" });

    const editor = screen.getByRole("textbox", { name: "Edit e1 Salary" });
    expect(editor).toHaveValue("4");
    await user.type(editor, "5{Enter}");
    expect(screen.getByRole("gridcell", { name: "e1 Salary" })).toHaveTextContent("45");
  });

  it("commits the option displayed by a select editor", async () => {
    const user = userEvent.setup();
    const session = createSession({
      defaultDocument: {
        version: 1,
        cells: {
          [createTableMetadataKey("e1", "name")]: {
            validation: { kind: "list", values: ["Grace", "Lin"] }
          }
        },
        calculatedColumns: [],
        namedStyles: []
      }
    });
    render(<DataTable aria-label="List editor employees" session={session} />);

    await user.dblClick(screen.getByRole("gridcell", { name: "e1 Name" }));
    const editor = screen.getByRole("combobox", { name: "Edit e1 Name" });
    expect(editor).toHaveValue("Grace");
    await user.keyboard("{Enter}");

    expect(screen.getByRole("gridcell", { name: "e1 Name" })).toHaveTextContent("Grace");
    session.destroy();
  });

  it("sorts, filters, groups, aggregates, and paginates the complete local dataset", async () => {
    const user = userEvent.setup();
    const ref = createRef<DataTableHandle>();
    renderTable({ ref });

    await dispatch(ref, { type: "set-sorting", sorting: [{ columnId: "salary", direction: "desc" }] });
    expect(dataRowTexts()[0]).toContain("Grace");
    await dispatch(ref, {
      type: "set-filter",
      filter: { kind: "comparison", columnId: "department", operator: "eq", value: { type: "string", value: "Engineering" } }
    });
    expect(dataRowTexts()).toHaveLength(2);
    await dispatch(ref, { type: "set-grouping", grouping: [{ columnId: "department" }] });
    await dispatch(ref, {
      type: "set-aggregates",
      aggregates: [{ id: "salary-sum", columnId: "salary", function: "sum" }]
    });
    expect(screen.getByRole("row", { name: /Engineering/ })).toHaveTextContent("320");

    await openColumnMenu(user, "Department");
    await user.click(screen.getByRole("button", { name: "Ungroup Department" }));
    await dispatch(ref, { type: "set-filter", filter: null });
    await dispatch(ref, { type: "set-pagination", pagination: { kind: "offset", offset: 0, limit: 1 } });
    expect(dataRowTexts()).toHaveLength(1);
    expect(screen.getByRole("row", { name: "Totals" })).toHaveAttribute("aria-rowindex", "5");
    await userClick("Next page");
    expect(dataRowTexts()[0]).toContain("Lin");
  });

  it("keeps one visible roving tab stop when selected rows or columns disappear", async () => {
    const user = userEvent.setup();
    const ref = createRef<DataTableHandle>();
    renderTable({ ref });
    const grid = screen.getByRole("grid", { name: "Employees" });
    await user.click(screen.getByRole("gridcell", { name: "e3 Salary" }));

    await dispatch(ref, {
      type: "set-filter",
      filter: {
        kind: "comparison",
        columnId: "department",
        operator: "eq",
        value: { type: "string", value: "Engineering" }
      }
    });

    expect(grid.querySelectorAll('[role="gridcell"][tabindex="0"]')).toHaveLength(1);
    expect(screen.getByRole("gridcell", { name: "e1 Name" })).toHaveAttribute("tabindex", "0");

    await dispatch(ref, { type: "set-filter", filter: null });
    await dispatch(ref, { type: "set-column-visibility", columnId: "salary", visible: false });

    expect(grid.querySelectorAll('[role="gridcell"][tabindex="0"]')).toHaveLength(1);
    expect(screen.getByRole("gridcell", { name: "e1 Name" })).toHaveAttribute("tabindex", "0");
  });

  it("numbers grouped data rows sequentially across group boundaries", async () => {
    const ref = createRef<DataTableHandle>();
    renderTable({ ref });

    await dispatch(ref, { type: "set-grouping", grouping: [{ columnId: "department" }] });

    const dataRows = screen.getAllByRole("row")
      .filter((row) => row.getAttribute("data-row-kind") === "data");
    expect(dataRows.map((row) => row.getAttribute("aria-label"))).toEqual([
      "Row 1",
      "Row 2",
      "Row 3"
    ]);
    expect(dataRows.map((row) => within(row).getByRole("rowheader").getAttribute("aria-label"))).toEqual([
      "Row 1",
      "Row 2",
      "Row 3"
    ]);
  });

  it("copies and atomically pastes a tab/newline matrix", async () => {
    const user = userEvent.setup();
    renderTable();
    await user.click(screen.getByRole("gridcell", { name: "e1 Name" }));
    const grid = screen.getByRole("grid", { name: "Employees" });
    const writeText = vi.spyOn(navigator.clipboard, "writeText").mockResolvedValue();

    fireEvent.copy(grid);
    await waitFor(() => expect(writeText).toHaveBeenCalledWith("Ada"));

    fireEvent.paste(grid, {
      clipboardData: { getData: () => "Changed\tEngineering\t130\nInvalid\tFinance\t-1" }
    });

    await waitFor(() => expect(screen.getByRole("alert")).toHaveTextContent("Salary must be non-negative"));
    expect(screen.getByRole("gridcell", { name: "e1 Name" })).toHaveTextContent("Ada");
    expect(screen.getByRole("gridcell", { name: "e2 Name" })).toHaveTextContent("Grace");
  });

  it("reports copy failure when the Clipboard API is unavailable", async () => {
    const user = userEvent.setup();
    renderTable();
    await user.click(screen.getByRole("gridcell", { name: "e1 Name" }));
    vi.stubGlobal("navigator", { ...navigator, clipboard: undefined });

    fireEvent.copy(screen.getByRole("grid", { name: "Employees" }));

    expect(await screen.findByRole("alert")).toHaveTextContent("Could not copy selection");
  });

  it("undoes and redoes edits from toolbar and handle", async () => {
    const user = userEvent.setup();
    const ref = createRef<DataTableHandle>();
    renderTable({ ref });
    await editDefaultCell(user, "e1 Name", "Katherine");

    await user.click(screen.getByRole("button", { name: "Undo" }));
    expect(screen.getByRole("gridcell", { name: "e1 Name" })).toHaveTextContent("Ada");
    await act(async () => {
      await ref.current!.redo();
    });
    expect(screen.getByRole("gridcell", { name: "e1 Name" })).toHaveTextContent("Katherine");
  });

  it("edits format, validation, comment, read-only, and formula metadata from quick tools", async () => {
    const user = userEvent.setup();
    const session = createSession({
      formulaService: { evaluate: () => ({ value: 42, displayValue: "42" }) }
    });
    render(<DataTable aria-label="Metadata employees" session={session} />);
    await user.click(screen.getByRole("gridcell", { name: "e1 Salary" }));
    await user.click(screen.getByRole("button", { name: "Quick tools" }));
    await user.selectOptions(screen.getByLabelText("Number format"), "currency");
    await user.click(screen.getByLabelText("Bold"));
    await user.type(screen.getByLabelText("Fill color"), "#eaf7f2");
    await user.type(screen.getByLabelText("Validation list"), "10,42,100");
    await user.type(screen.getByLabelText("Comment"), "Reviewed");
    await user.type(screen.getByLabelText("Formula"), "=salary");
    await user.click(screen.getByLabelText("Read only"));
    await user.click(screen.getByRole("button", { name: "Apply quick tools" }));

    const metadata = session.getSnapshot().getCell("e1", "salary").metadata;
    expect(metadata).toMatchObject({
      format: { numberFormat: "currency", bold: true, backgroundColor: "#eaf7f2" },
      validation: { kind: "list", values: ["10", "42", "100"] },
      comment: "Reviewed",
      formula: "=salary",
      readOnly: true
    });
    session.destroy();
  });

  it("clears quick-tool fill color, validation, comment, and formula values", async () => {
    const user = userEvent.setup();
    const session = createSession({
      formulaService: { evaluate: () => ({ value: 42, displayValue: "42" }) }
    });
    render(<DataTable aria-label="Metadata employees" session={session} />);
    await user.click(screen.getByRole("gridcell", { name: "e1 Salary" }));
    await user.click(screen.getByRole("button", { name: "Quick tools" }));
    await user.type(screen.getByLabelText("Fill color"), "#eaf7f2");
    await user.type(screen.getByLabelText("Validation list"), "10,42,100");
    await user.type(screen.getByLabelText("Comment"), "Reviewed");
    await user.type(screen.getByLabelText("Formula"), "=salary");
    await user.click(screen.getByRole("button", { name: "Apply quick tools" }));

    await user.clear(screen.getByLabelText("Fill color"));
    await user.clear(screen.getByLabelText("Validation list"));
    await user.clear(screen.getByLabelText("Comment"));
    await user.clear(screen.getByLabelText("Formula"));
    await user.click(screen.getByRole("button", { name: "Apply quick tools" }));

    expect(session.getSnapshot().getCell("e1", "salary").metadata).toEqual({ format: {} });
    session.destroy();
  });

  it("does not overwrite untouched metadata when applying one quick tool", async () => {
    const user = userEvent.setup();
    const session = createSession({
      defaultDocument: {
        version: 1,
        cells: {
          [createTableMetadataKey("e1", "salary")]: {
            format: { numberFormat: "currency", bold: true, textColor: "red" },
            readOnly: true
          }
        },
        calculatedColumns: [],
        namedStyles: []
      }
    });
    render(<DataTable aria-label="Protected metadata employees" session={session} />);
    await user.click(screen.getByRole("gridcell", { name: "e1 Salary" }));
    await user.click(screen.getByRole("button", { name: "Quick tools" }));
    await user.type(screen.getByLabelText("Comment"), "Reviewed");

    await user.click(screen.getByRole("button", { name: "Apply quick tools" }));

    expect(session.getSnapshot().getCell("e1", "salary").metadata).toEqual({
      format: { numberFormat: "currency", bold: true, textColor: "red" },
      comment: "Reviewed",
      readOnly: true
    });
    session.destroy();
  });

  it("does not reuse applied quick-tool state for another selection", async () => {
    const user = userEvent.setup();
    const session = createSession();
    const dispatch = vi.spyOn(session, "dispatch");
    render(<DataTable aria-label="Metadata employees" session={session} />);

    await user.click(screen.getByRole("gridcell", { name: "e1 Name" }));
    await user.click(screen.getByRole("button", { name: "Quick tools" }));
    await user.click(screen.getByLabelText("Read only"));
    await user.click(screen.getByRole("button", { name: "Apply quick tools" }));
    expect(session.getSnapshot().getCell("e1", "name").metadata.readOnly).toBe(true);

    await user.click(screen.getByRole("button", { name: "Quick tools" }));
    await user.click(screen.getByRole("gridcell", { name: "e2 Name" }));
    await user.click(screen.getByRole("button", { name: "Quick tools" }));

    expect(screen.getByLabelText("Read only")).not.toBeChecked();
    await user.type(screen.getByLabelText("Comment"), "B only");
    await user.click(screen.getByRole("button", { name: "Apply quick tools" }));

    expect(dispatch).toHaveBeenLastCalledWith({
      type: "update-cell-metadata",
      updates: [{ rowId: "e2", columnId: "name", patch: { comment: "B only" } }]
    });
    expect(session.getSnapshot().getCell("e2", "name").metadata).toEqual({ comment: "B only" });
    session.destroy();
  });

  it("resizes, hides, and pins columns by stable id", async () => {
    const user = userEvent.setup();
    const session = createSession();
    render(<DataTable aria-label="Column employees" session={session} />);
    await openColumnMenu(user, "Salary");
    const width = screen.getByLabelText("Salary column width");
    await user.clear(width);
    await user.type(width, "180");
    expect(width).toHaveValue(180);
    await user.click(screen.getByRole("button", { name: "Resize Salary" }));
    await user.click(screen.getByRole("button", { name: "Pin Salary left" }));
    await user.click(screen.getByRole("button", { name: "Hide Salary" }));

    expect(session.getSnapshot().state.columnWidths.salary).toBe(180);
    expect(session.getSnapshot().state.columnPinning.left).toContain("salary");
    expect(screen.queryByRole("columnheader", { name: "Salary" })).not.toBeInTheDocument();
    expect(screen.getByRole("grid", { name: "Column employees" })).toHaveAttribute("aria-colcount", "5");
    await user.click(screen.getByRole("button", { name: "Show column Salary" }));
    expect(screen.getByRole("columnheader", { name: "Salary" })).toBeInTheDocument();
    session.destroy();
  });

  it("opens column options in a scoped native popover and restores focus on Escape", async () => {
    const user = userEvent.setup();
    const showPopover = vi.spyOn(HTMLElement.prototype, "showPopover");
    const hidePopover = vi.spyOn(HTMLElement.prototype, "hidePopover");
    renderTable();

    const trigger = screen.getByRole("button", { name: "Column options for Name" });
    expect(trigger).toHaveAttribute("aria-haspopup", "dialog");
    await user.click(trigger);
    const menu = screen.getByRole("dialog", { name: "Name column menu" });

    expect(menu).toHaveAttribute("popover", "auto");
    expect(menu.closest(".js-spreadsheet-data-table")).not.toBeNull();
    expect(showPopover).toHaveBeenCalledTimes(1);
    expect(screen.getByRole("button", { name: "Sort Name ascending" })).toHaveFocus();

    fireEvent.keyDown(menu, { key: "Escape" });

    expect(hidePopover).toHaveBeenCalledTimes(1);
    expect(screen.queryByRole("dialog", { name: "Name column menu" })).not.toBeInTheDocument();
    expect(trigger).toHaveFocus();
  });

  it("keeps a column menu open through StrictMode effect replay", async () => {
    const user = userEvent.setup();
    render(
      <StrictMode>
        <DataTable
          aria-label="Strict employees"
          rows={employees}
          columns={columns}
          getRowId={getRowId}
        />
      </StrictMode>
    );

    const trigger = screen.getByRole("button", { name: "Column options for Name" });
    await user.click(trigger);

    expect(screen.getByRole("dialog", { name: "Name column menu" })).toBeInTheDocument();
    expect(trigger).toHaveAttribute("aria-expanded", "true");
    expect(screen.getByRole("button", { name: "Sort Name ascending" })).toHaveFocus();
  });

  it("closes and cleanly reopens a column menu from the same trigger", async () => {
    const user = userEvent.setup();
    const showPopover = vi.spyOn(HTMLElement.prototype, "showPopover");
    const hidePopover = vi.spyOn(HTMLElement.prototype, "hidePopover");
    renderTable();
    const trigger = screen.getByRole("button", { name: "Column options for Name" });

    await user.click(trigger);
    const firstMenu = screen.getByRole("dialog", { name: "Name column menu" });
    await user.click(trigger);

    expect(firstMenu).not.toBeInTheDocument();
    expect(firstMenu).not.toHaveAttribute("data-popover-open");
    expect(firstMenu.style.display).toBe("");
    expect(trigger).toHaveAttribute("aria-expanded", "false");
    expect(trigger).toHaveFocus();
    expect(showPopover).toHaveBeenCalledTimes(1);
    expect(hidePopover).toHaveBeenCalledTimes(1);

    await user.click(trigger);

    expect(screen.getByRole("dialog", { name: "Name column menu" })).not.toBe(firstMenu);
    expect(screen.getByRole("button", { name: "Sort Name ascending" })).toHaveFocus();
    expect(showPopover).toHaveBeenCalledTimes(2);
    expect(hidePopover).toHaveBeenCalledTimes(1);
  });

  it("switches directly between column triggers without stale menu ownership", async () => {
    const user = userEvent.setup();
    const showPopover = vi.spyOn(HTMLElement.prototype, "showPopover");
    const hidePopover = vi.spyOn(HTMLElement.prototype, "hidePopover");
    renderTable();
    const nameTrigger = screen.getByRole("button", { name: "Column options for Name" });
    const departmentTrigger = screen.getByRole("button", { name: "Column options for Department" });

    await user.click(nameTrigger);
    const staleMenu = screen.getByRole("dialog", { name: "Name column menu" });
    await user.click(departmentTrigger);

    expect(staleMenu).not.toBeInTheDocument();
    expect(screen.getByRole("dialog", { name: "Department column menu" })).toBeInTheDocument();
    expect(nameTrigger).toHaveAttribute("aria-expanded", "false");
    expect(departmentTrigger).toHaveAttribute("aria-expanded", "true");
    expect(screen.getByRole("button", { name: "Sort Department ascending" })).toHaveFocus();
    expect(showPopover).toHaveBeenCalledTimes(2);
    expect(hidePopover).toHaveBeenCalledTimes(1);

    const staleToggle = new Event("toggle");
    Object.defineProperty(staleToggle, "newState", { value: "closed" });
    fireEvent(staleMenu, staleToggle);

    expect(screen.getByRole("dialog", { name: "Department column menu" })).toBeInTheDocument();
    expect(screen.getByRole("button", { name: "Sort Department ascending" })).toHaveFocus();
  });

  it.each([
    ["top", { left: 120, right: 142, top: 0, bottom: 22 }],
    ["right", { left: 298, right: 320, top: 90, bottom: 112 }],
    ["bottom", { left: 120, right: 142, top: 218, bottom: 240 }],
    ["left", { left: 0, right: 22, top: 90, bottom: 112 }]
  ])("keeps a fixed column menu inside the viewport at the %s edge", async (_edge, anchorBounds) => {
    vi.stubGlobal("innerWidth", 320);
    vi.stubGlobal("innerHeight", 240);
    const user = userEvent.setup();
    renderTable();
    const trigger = screen.getByRole("button", { name: "Column options for Name" });
    const nativeRect = HTMLElement.prototype.getBoundingClientRect;
    vi.spyOn(HTMLElement.prototype, "getBoundingClientRect").mockImplementation(function (this: HTMLElement) {
      if (this === trigger) return rect(anchorBounds);
      if (this.classList.contains("js-spreadsheet-data-table__column-menu")) {
        return rect({ left: 0, right: 248, top: 0, bottom: 180 });
      }
      return nativeRect.call(this);
    });

    await user.click(trigger);
    const menu = screen.getByRole("dialog", { name: "Name column menu" });
    const left = Number.parseFloat(menu.style.left);
    const top = Number.parseFloat(menu.style.top);

    expect(menu).toHaveStyle({ position: "fixed" });
    expect(left).toBeGreaterThanOrEqual(8);
    expect(top).toBeGreaterThanOrEqual(8);
    expect(left + 248).toBeLessThanOrEqual(312);
    expect(top + 180).toBeLessThanOrEqual(232);
  });

  it("reclamps an open column menu when the viewport resizes", async () => {
    vi.stubGlobal("innerWidth", 640);
    vi.stubGlobal("innerHeight", 480);
    const user = userEvent.setup();
    renderTable();
    const trigger = screen.getByRole("button", { name: "Column options for Name" });
    const nativeRect = HTMLElement.prototype.getBoundingClientRect;
    vi.spyOn(HTMLElement.prototype, "getBoundingClientRect").mockImplementation(function (this: HTMLElement) {
      if (this === trigger) return rect({ left: 590, right: 612, top: 430, bottom: 452 });
      if (this.classList.contains("js-spreadsheet-data-table__column-menu")) {
        return rect({ left: 0, right: 248, top: 0, bottom: 180 });
      }
      return nativeRect.call(this);
    });
    await user.click(trigger);
    const menu = screen.getByRole("dialog", { name: "Name column menu" });

    vi.stubGlobal("innerWidth", 320);
    vi.stubGlobal("innerHeight", 240);
    fireEvent.resize(window);

    expect(Number.parseFloat(menu.style.left) + 248).toBeLessThanOrEqual(312);
    expect(Number.parseFloat(menu.style.top) + 180).toBeLessThanOrEqual(232);
  });

  it("clears menu state and restores trigger focus after native light dismiss", async () => {
    const user = userEvent.setup();
    renderTable();
    const trigger = screen.getByRole("button", { name: "Column options for Name" });
    await user.click(trigger);
    const menu = screen.getByRole("dialog", { name: "Name column menu" });
    const toggle = new Event("toggle");
    Object.defineProperty(toggle, "newState", { value: "closed" });

    fireEvent(menu, toggle);

    expect(screen.queryByRole("dialog", { name: "Name column menu" })).not.toBeInTheDocument();
    expect(trigger).toHaveFocus();
  });

  it("preserves externally acquired cell focus after native light dismiss", async () => {
    const user = userEvent.setup();
    renderTable();
    const trigger = screen.getByRole("button", { name: "Column options for Name" });
    await user.click(trigger);
    const menu = screen.getByRole("dialog", { name: "Name column menu" });
    const cell = screen.getByRole("gridcell", { name: "e1 Name" });
    cell.focus();
    const toggle = new Event("toggle");
    Object.defineProperty(toggle, "newState", { value: "closed" });

    fireEvent(menu, toggle);

    expect(screen.queryByRole("dialog", { name: "Name column menu" })).not.toBeInTheDocument();
    expect(cell).toHaveFocus();
    expect(trigger).not.toHaveFocus();
  });

  it("closes an open column menu on viewport scroll", async () => {
    const user = userEvent.setup();
    renderTable();
    const trigger = screen.getByRole("button", { name: "Column options for Name" });
    await user.click(trigger);

    fireEvent.scroll(screen.getByRole("grid", { name: "Employees" }));

    expect(screen.queryByRole("dialog", { name: "Name column menu" })).not.toBeInTheDocument();
    expect(trigger).toHaveFocus();
  });

  it("preserves externally acquired cell focus when incidental scrolling closes the menu", async () => {
    const user = userEvent.setup();
    renderTable();
    const trigger = screen.getByRole("button", { name: "Column options for Name" });
    await user.click(trigger);
    const cell = screen.getByRole("gridcell", { name: "e1 Name" });
    cell.focus();

    fireEvent.scroll(screen.getByRole("grid", { name: "Employees" }));

    expect(screen.queryByRole("dialog", { name: "Name column menu" })).not.toBeInTheDocument();
    expect(cell).toHaveFocus();
    expect(trigger).not.toHaveFocus();
  });

  it("keeps a column menu open when focus scrolls inside its narrow header", async () => {
    const user = userEvent.setup();
    renderTable();
    const trigger = screen.getByRole("button", { name: "Column options for Name" });
    await user.click(trigger);
    const header = trigger.closest('[role="columnheader"]');

    expect(header).not.toBeNull();
    fireEvent.scroll(header!);

    expect(screen.getByRole("dialog", { name: "Name column menu" })).toBeInTheDocument();
    expect(trigger).toHaveAttribute("aria-expanded", "true");
  });

  it("observes only the nearest DataTable root while a column menu is open", async () => {
    const observe = vi.spyOn(MutationObserver.prototype, "observe");
    const user = userEvent.setup();
    const { container } = renderTable();
    const tableRoot = container.querySelector('[data-js-spreadsheet-root="data-table"]');

    await user.click(screen.getByRole("button", { name: "Column options for Name" }));

    expect(tableRoot).not.toBeNull();
    expect(observe).toHaveBeenCalledWith(tableRoot, { childList: true, subtree: true });
    expect(observe).not.toHaveBeenCalledWith(document, { childList: true, subtree: true });
  });

  it("hides the native popover when its column trigger unmounts", async () => {
    const user = userEvent.setup();
    const hidePopover = vi.spyOn(HTMLElement.prototype, "hidePopover");
    const { rerender } = renderTable();
    await user.click(screen.getByRole("button", { name: "Column options for Name" }));

    rerender(
      <DataTable
        aria-label="Employees"
        rows={employees}
        columns={columns.filter((column) => column.id !== "name")}
        getRowId={getRowId}
      />
    );

    expect(screen.queryByRole("dialog", { name: "Name column menu" })).not.toBeInTheDocument();
    expect(screen.queryByRole("button", { name: "Column options for Name" })).not.toBeInTheDocument();
    expect(hidePopover).toHaveBeenCalledTimes(1);
  });

  it("removes popover listeners and hides the menu when the table unmounts", async () => {
    const user = userEvent.setup();
    const hidePopover = vi.spyOn(HTMLElement.prototype, "hidePopover");
    const getBounds = vi.spyOn(HTMLElement.prototype, "getBoundingClientRect");
    const { unmount } = renderTable();
    await user.click(screen.getByRole("button", { name: "Column options for Name" }));
    const callsBeforeUnmount = getBounds.mock.calls.length;

    unmount();
    fireEvent.resize(window);
    fireEvent.scroll(window);

    expect(hidePopover).toHaveBeenCalledTimes(1);
    expect(getBounds).toHaveBeenCalledTimes(callsBeforeUnmount);
  });

  it("allows explicit visibility and pinning state to override column defaults", async () => {
    const user = userEvent.setup();
    const defaultedColumns: readonly ColumnDef<Employee>[] = [
      { ...columns[0], pin: "left" },
      { ...columns[2], visible: false }
    ];
    const session = createLocalRecordTableSession<Employee, ColumnDef<Employee>>({
      source: { kind: "local", rows: employees, getRowId },
      columns: defaultedColumns
    });
    render(<DataTable aria-label="Defaulted columns" session={session} />);

    expect(screen.queryByRole("columnheader", { name: "Salary" })).not.toBeInTheDocument();
    await user.click(screen.getByRole("button", { name: "Show column Salary" }));
    expect(screen.getByRole("columnheader", { name: "Salary" })).toBeInTheDocument();
    await openColumnMenu(user, "Name");
    await user.click(screen.getByRole("button", { name: "Unpin Name" }));

    expect(session.getSnapshot().state.columnPinning.left).not.toContain("name");
    expect(screen.getAllByRole("columnheader", { name: "Name" })).toHaveLength(1);
    session.destroy();
  });

  it("reorders columns visibly by pointer drag and keyboard controls", async () => {
    const user = userEvent.setup();
    const session = createSession();
    render(<DataTable aria-label="Order employees" session={session} />);
    const nameHandle = screen.getByRole("button", { name: "Reorder Name" });
    const departmentHeader = screen.getByRole("columnheader", { name: "Department" });
    fireEvent.dragStart(nameHandle);
    fireEvent.dragOver(departmentHeader);
    fireEvent.drop(departmentHeader);
    await waitFor(() => expect(session.getSnapshot().state.columnOrder.indexOf("name")).toBeGreaterThan(0));

    await openColumnMenu(user, "Name");
    await user.click(screen.getByRole("button", { name: "Move Name left" }));
    expect(visibleHeaderNames()[0]).toBe("Name");
    expect(screen.getByRole("status")).toHaveTextContent("Name moved to position 1");
    session.destroy();
  });

  it("does not announce a column move when the session reports no change", async () => {
    const user = userEvent.setup();
    const session = createSession();
    vi.spyOn(session, "dispatch").mockResolvedValueOnce({
      status: "committed",
      revision: session.getSnapshot().revision,
      changed: false
    });
    render(<DataTable aria-label="No-op order employees" session={session} />);

    await openColumnMenu(user, "Name");
    await user.click(screen.getByRole("button", { name: "Move Name right" }));

    expect(screen.getByRole("status")).not.toHaveTextContent("Name moved");
    session.destroy();
  });

  it("drops a right-moving column after its target and reaches the final position", async () => {
    const session = createSession();
    render(<DataTable aria-label="Right drop employees" session={session} />);

    fireEvent.dragStart(screen.getByRole("button", { name: "Reorder Name" }));
    fireEvent.drop(screen.getByRole("columnheader", { name: "Department" }));
    await waitFor(() => expect(visibleHeaderNames()).toEqual([
      "Department",
      "Name",
      "Salary",
      "Active"
    ]));

    fireEvent.dragStart(screen.getByRole("button", { name: "Reorder Name" }));
    fireEvent.drop(screen.getByRole("columnheader", { name: "Active" }));
    await waitFor(() => expect(visibleHeaderNames()).toEqual([
      "Department",
      "Salary",
      "Active",
      "Name"
    ]));
    expect(screen.getByRole("status")).toHaveTextContent("Name moved to position 4");
    session.destroy();
  });

  it("renders checkbox row selection with single and multiple modes", async () => {
    const user = userEvent.setup();
    const onRowSelectionChange = vi.fn();
    const first = renderTable({ rowSelection: "multiple", onRowSelectionChange });
    await user.click(screen.getByRole("checkbox", { name: "Select row e1" }));
    await user.click(screen.getByRole("checkbox", { name: "Select row e2" }));
    expect(onRowSelectionChange).toHaveBeenLastCalledWith(["e1", "e2"]);
    await user.click(screen.getByRole("checkbox", { name: "Select all rows" }));
    expect(onRowSelectionChange).toHaveBeenLastCalledWith(["e1", "e2", "e3"]);
    first.unmount();

    renderTable({ rowSelection: "single", onRowSelectionChange });
    await user.click(screen.getByRole("checkbox", { name: "Select row e1" }));
    await user.click(screen.getByRole("checkbox", { name: "Select row e2" }));
    expect(onRowSelectionChange).toHaveBeenLastCalledWith(["e2"]);
  });

  it("expands and collapses tree rows by stable id", async () => {
    const user = userEvent.setup();
    renderTable({ getSubRows: (row) => row.children });

    expect(screen.queryByRole("gridcell", { name: "e1-child Name" })).not.toBeInTheDocument();
    await user.click(screen.getByRole("button", { name: "Expand Ada" }));
    expect(screen.getByRole("gridcell", { name: "e1-child Name" })).toHaveTextContent("Byron");
    await user.click(screen.getByRole("button", { name: "Collapse Ada" }));
    expect(screen.queryByRole("gridcell", { name: "e1-child Name" })).not.toBeInTheDocument();
  });

  it("toggles a tree row on the first click without selecting or scrolling the grid", () => {
    const ref = createRef<DataTableHandle>();
    renderTable({
      ref,
      getSubRows: (row) => row.children,
      defaultState: { expandedRowIds: ["e1"] }
    });
    const grid = screen.getByRole("grid", { name: "Employees" });
    grid.scrollLeft = 37;
    const selectionBefore = ref.current?.getSelection();
    const toggle = screen.getByRole("button", { name: "Collapse Ada" });
    const treeContent = toggle.closest(".js-spreadsheet-data-table__cell-content--tree");
    const pointerDown = new Event("pointerdown", { bubbles: true, cancelable: true });
    const mouseDown = new MouseEvent("mousedown", { bubbles: true, cancelable: true });

    fireEvent(toggle, pointerDown);
    fireEvent(toggle, mouseDown);
    fireEvent.click(toggle);

    expect(pointerDown.defaultPrevented).toBe(true);
    expect(mouseDown.defaultPrevented).toBe(true);
    expect(toggle).toHaveClass("js-spreadsheet-data-table__tree-toggle");
    expect(treeContent).not.toBeNull();
    expect(treeContent?.closest("[role=gridcell]")).toHaveStyle({ height: "32px" });
    expect(screen.queryByRole("gridcell", { name: "e1-child Name" })).not.toBeInTheDocument();
    expect(ref.current?.getSelection()).toEqual(selectionBefore);
    expect(grid.scrollLeft).toBe(37);
  });

  it("does not apply tree-toggle sizing to host-rendered cell buttons", () => {
    const actionColumns: readonly ColumnDef<Employee>[] = [{
      ...columns[0],
      cell: ({ row }) => <button type="button" aria-label={`Host action ${row.id}`}>Action</button>
    }];
    render(
      <DataTable
        aria-label="Tree actions"
        rows={employees}
        columns={actionColumns}
        getRowId={getRowId}
        getSubRows={(row) => row.children}
      />
    );

    const toggle = screen.getByRole("button", { name: "Expand Ada" });
    const hostAction = screen.getByRole("button", { name: "Host action e1" });
    expect(toggle).toHaveClass("js-spreadsheet-data-table__tree-toggle");
    expect(hostAction).not.toHaveClass("js-spreadsheet-data-table__tree-toggle");
    expect(hostAction.closest(".js-spreadsheet-data-table__cell-content--tree")).not.toBeNull();
  });

  it("renders an inline typed filter row without replacing column menus", async () => {
    const user = userEvent.setup();
    render(
      <DataTable
        aria-label="Employees"
        rows={employees}
        columns={columns}
        getRowId={getRowId}
        inlineFilters
        defaultState={{
          filter: {
            kind: "comparison",
            columnId: "department",
            operator: "eq",
            value: { type: "string", value: "Finance" }
          }
        }}
      />
    );

    const filter = screen.getByRole("textbox", { name: "Inline filter Name" });
    expect(screen.getByRole("spinbutton", { name: "Inline filter Salary" })).toBeInTheDocument();
    expect(screen.getByRole("combobox", { name: "Inline filter Active" })).toBeInTheDocument();
    await user.type(filter, "Grace");
    await waitFor(() => expect(dataRowTexts()).toHaveLength(0));
    await user.clear(filter);
    await waitFor(() => expect(dataRowTexts()).toHaveLength(1));
    expect(screen.getByRole("gridcell", { name: "e3 Name" })).toHaveTextContent("Lin");
    expect(screen.getByRole("button", { name: "Column options for Name" })).toBeInTheDocument();
  });

  it("disables capability-gated toolbar actions with an accessible reason", async () => {
    const user = userEvent.setup();
    render(
      <DataTable
        aria-label="Restricted employees"
        rows={employees}
        columns={columns}
        getRowId={getRowId}
        features={{ metadata: false, export: false }}
      />
    );
    await user.click(screen.getByRole("gridcell", { name: "e1 Name" }));
    await user.click(screen.getByRole("button", { name: "Quick tools" }));

    expect(screen.getByRole("button", { name: "Apply quick tools" })).toBeDisabled();
    expect(screen.getByRole("button", { name: "Export CSV" })).toBeDisabled();
    expect(screen.getAllByText("Disabled by host configuration").length).toBeGreaterThan(0);
  });

  it("does not expose quick metadata tools for an unpermitted selected cell", async () => {
    const user = userEvent.setup();
    const restrictedColumns: readonly ColumnDef<Employee>[] = [{ ...columns[0], permitted: () => false }];
    render(
      <DataTable
        aria-label="Permission restricted employees"
        rows={employees}
        columns={restrictedColumns}
        getRowId={getRowId}
      />
    );
    await user.click(screen.getByRole("gridcell", { name: "e1 Name" }));
    await user.click(screen.getByRole("button", { name: "Quick tools" }));

    expect(screen.getByRole("button", { name: "Apply quick tools" })).toBeDisabled();
    expect(screen.getByText("Selection contains cells that do not permit metadata changes")).toBeVisible();
  });

  it("runs custom header actions and reports safe host errors", async () => {
    const user = userEvent.setup();
    const onDiagnostic = vi.fn<(event: TableDiagnosticEvent) => void>();
    const run = vi.fn().mockRejectedValue(new Error("secret employee data"));
    const disabledRun = vi.fn();
    const actionColumns: readonly ColumnDef<Employee>[] = [{
      ...columns[0],
      headerActions: [
        { id: "fail", label: "Fail safely", run },
        { id: "disabled", label: "Unavailable action", disabled: () => true, run: disabledRun }
      ]
    }];
    render(
      <DataTable
        aria-label="Header actions"
        rows={employees}
        columns={actionColumns}
        getRowId={getRowId}
        onDiagnostic={onDiagnostic}
      />
    );
    await openColumnMenu(user, "Name");
    expect(screen.getByRole("button", { name: "Unavailable action" })).toBeDisabled();
    await user.click(screen.getByRole("button", { name: "Fail safely" }));

    expect(run).toHaveBeenCalledTimes(1);
    expect(onDiagnostic).toHaveBeenCalledWith({
      category: "extension",
      metadata: { columnId: "name", slot: "header-action" }
    });
    expect(JSON.stringify(onDiagnostic.mock.calls)).not.toContain("secret employee data");
    expect(disabledRun).not.toHaveBeenCalled();
  });

  it("measures wrapped auto-height rows without breaking virtualization", async () => {
    vi.stubGlobal("ResizeObserver", class {
      constructor(private readonly callback: ResizeObserverCallback) {}
      observe(target: Element) {
        if ((target as HTMLElement).dataset.rowMeasure) {
          const gridCell = target.closest<HTMLElement>("[role=gridcell]");
          const height = gridCell?.getAttribute("aria-label") === "e1 Name" ? 72 : 28;
          this.callback([{ target, contentRect: { height } } as ResizeObserverEntry], this as never);
        }
      }
      unobserve() {}
      disconnect() {}
    });
    renderTable({ rowHeight: "auto" });

    await waitFor(() => expect(screen.getByRole("row", { name: "Row 1" })).toHaveStyle({ height: "72px" }));
    expect(screen.getByRole("grid", { name: "Employees" })).toHaveAttribute("data-viewport-kernel", "shared");
  });

  it("preserves the vertical scroll anchor when an earlier auto-height row grows", async () => {
    const observations: Array<{ target: Element; callback: ResizeObserverCallback; observer: ResizeObserver }> = [];
    vi.stubGlobal("ResizeObserver", class {
      constructor(private readonly callback: ResizeObserverCallback) {}
      observe(target: Element) {
        observations.push({ target, callback: this.callback, observer: this as never });
      }
      unobserve() {}
      disconnect() {}
    });
    renderTable({ rowHeight: "auto" });
    const grid = screen.getByRole("grid", { name: "Employees" });
    await waitFor(() => expect(screen.getByRole("row", { name: "Row 1" })).toHaveStyle({ height: "28px" }));
    grid.scrollTop = 80;
    const observed = observations.find(({ target }) =>
      target.closest<HTMLElement>("[role=gridcell]")?.getAttribute("aria-label") === "e1 Name"
    );
    expect(observed).toBeDefined();

    act(() => {
      observed!.callback([{
        target: observed!.target,
        contentRect: { height: 72 }
      } as ResizeObserverEntry], observed!.observer);
    });

    expect(grid.scrollTop).toBe(124);
  });

  it("supports fixed fluid and ratio layouts plus a custom empty message", () => {
    const { rerender } = render(
      <DataTable
        aria-label="Empty employees"
        rows={[] as Employee[]}
        columns={columns}
        getRowId={getRowId}
        layout="ratio"
        noDataMessage={<span>No employees yet</span>}
      />
    );
    expect(screen.getByText("No employees yet").closest("[data-table-layout]")).toHaveAttribute("data-table-layout", "ratio");

    rerender(<DataTable aria-label="Empty employees" rows={[]} columns={columns} getRowId={getRowId} layout="fluid" />);
    expect(screen.getByRole("grid", { name: "Empty employees" }).closest("[data-table-layout]")).toHaveAttribute("data-table-layout", "fluid");
    rerender(<DataTable aria-label="Empty employees" rows={[]} columns={columns} getRowId={getRowId} layout="fixed" />);
    expect(screen.getByRole("grid", { name: "Empty employees" }).closest("[data-table-layout]")).toHaveAttribute("data-table-layout", "fixed");
  });

  it("sizes ratio columns from the table container instead of a viewport constant", async () => {
    vi.stubGlobal("ResizeObserver", class {
      constructor(private readonly callback: ResizeObserverCallback) {}
      observe(target: Element) {
        this.callback([{ target, contentRect: { width: 456 } } as ResizeObserverEntry], this as never);
      }
      unobserve() {}
      disconnect() {}
    });
    renderTable({ layout: "ratio" });

    await waitFor(() => {
      expect(Number.parseFloat(screen.getByRole("columnheader", { name: "Name" }).style.width)).toBeCloseTo(125.49, 1);
    });
  });

  it("visibly disables a formula action when no formula service exists", async () => {
    const user = userEvent.setup();
    renderTable();
    await user.click(screen.getByRole("gridcell", { name: "e1 Name" }));
    await user.click(screen.getByRole("button", { name: "Quick tools" }));

    expect(screen.getByLabelText("Formula")).toBeDisabled();
    expect(screen.getByText("Formula service is not configured")).toBeVisible();
  });

  it("owns simultaneous quick-tools drawers outside each toolbar overflow box", async () => {
    const user = userEvent.setup();
    render(
      <>
        <DataTable aria-label="First tools" rows={employees} columns={columns} getRowId={getRowId} />
        <DataTable aria-label="Second tools" rows={employees} columns={columns} getRowId={getRowId} />
      </>
    );
    const roots = [...document.querySelectorAll<HTMLElement>(".js-spreadsheet-data-table")];
    const [firstRoot, secondRoot] = roots;
    const firstToolbar = within(firstRoot).getByRole("toolbar", { name: "Data table toolbar" });
    const secondToolbar = within(secondRoot).getByRole("toolbar", { name: "Data table toolbar" });
    const firstToggle = within(firstToolbar).getByRole("button", { name: "Quick tools" });
    const secondToggle = within(secondToolbar).getByRole("button", { name: "Quick tools" });
    expect(firstToggle).not.toHaveAttribute("aria-controls");
    expect(secondToggle).not.toHaveAttribute("aria-controls");

    await user.click(firstToggle);
    await user.click(secondToggle);

    const firstDrawer = within(firstRoot).getByLabelText("Quick tools drawer");
    const secondDrawer = within(secondRoot).getByLabelText("Quick tools drawer");
    expect(firstToolbar).not.toContainElement(firstDrawer);
    expect(secondToolbar).not.toContainElement(secondDrawer);
    expect(firstToggle).toHaveAttribute("aria-controls", firstDrawer.id);
    expect(secondToggle).toHaveAttribute("aria-controls", secondDrawer.id);
    expect(firstDrawer.id).not.toBe(secondDrawer.id);

    await user.type(within(firstDrawer).getByLabelText("Fill color"), "#111111");
    await user.type(within(secondDrawer).getByLabelText("Fill color"), "#222222");
    expect(within(firstDrawer).getByLabelText("Fill color")).toHaveValue("#111111");
    expect(within(secondDrawer).getByLabelText("Fill color")).toHaveValue("#222222");
  });

  it("isolates two simultaneous DataTable instances", async () => {
    const user = userEvent.setup();
    render(
      <>
        <DataTable aria-label="First employees" rows={employees} columns={columns} getRowId={getRowId} />
        <DataTable aria-label="Second employees" rows={employees} columns={columns} getRowId={getRowId} />
      </>
    );
    const firstGrid = screen.getByRole("grid", { name: "First employees" });
    const secondGrid = screen.getByRole("grid", { name: "Second employees" });
    await user.dblClick(within(firstGrid).getByRole("gridcell", { name: "e1 Name" }));
    const editor = within(firstGrid).getByRole("textbox", { name: "Edit e1 Name" });
    await user.clear(editor);
    await user.type(editor, "First only{Enter}");

    expect(within(firstGrid).getByRole("gridcell", { name: "e1 Name" })).toHaveTextContent("First only");
    expect(within(secondGrid).getByRole("gridcell", { name: "e1 Name" })).toHaveTextContent("Ada");
    expect(firstGrid.getAttribute("aria-describedby")).not.toBe(secondGrid.getAttribute("aria-describedby"));
  });
});

function renderTable(
  overrides: Partial<DataTablePresentationProps & {
    resetKey: string | number;
    onRowsChange(updater: RowUpdater<Employee>, context: ChangeContext): void;
    getSubRows(row: Employee): readonly Employee[] | undefined;
    defaultState: Partial<TableViewState>;
    ref: React.Ref<DataTableHandle>;
  }> = {}
) {
  return render(
    <DataTable
      aria-label="Employees"
      rows={employees}
      columns={columns}
      getRowId={getRowId}
      {...overrides}
    />
  );
}

function createSession(
  overrides: Partial<Parameters<typeof createLocalRecordTableSession<Employee, ColumnDef<Employee>>>[0]> = {}
) {
  return createLocalRecordTableSession<Employee, ColumnDef<Employee>>({
    source: { kind: "local", rows: employees, getRowId },
    columns,
    ...overrides
  });
}

function getRowId(row: Employee) {
  return row.id;
}

async function editDefaultCell(user: ReturnType<typeof userEvent.setup>, name: string, value: string) {
  await user.dblClick(screen.getByRole("gridcell", { name }));
  const editor = screen.getByRole("textbox", { name: `Edit ${name}` });
  await user.clear(editor);
  await user.type(editor, `${value}{Enter}`);
}

function CustomNameEditor({ rawText, onChange, onCommit, onCancel }: CellEditorProps<Employee, unknown>) {
  function handleKeyDown(event: KeyboardEvent<HTMLInputElement>) {
    if (event.key === "Escape") onCancel();
  }
  return (
    <div>
      <input
        autoFocus
        aria-label="Custom name editor"
        value={rawText}
        onChange={(event) => onChange(event.currentTarget.value)}
        onKeyDown={handleKeyDown}
      />
      <button type="button" onClick={onCommit}>Commit custom name</button>
    </div>
  );
}

async function dispatch(ref: React.RefObject<DataTableHandle | null>, intent: Parameters<DataTableHandle["dispatch"]>[0]) {
  await act(async () => {
    await ref.current!.dispatch(intent);
  });
}

function dataRowTexts(): string[] {
  return screen.getAllByRole("row")
    .filter((row) => row.getAttribute("data-row-kind") === "data")
    .map((row) => row.textContent ?? "");
}

async function userClick(name: string) {
  const user = userEvent.setup();
  await user.click(screen.getByRole("button", { name }));
}

async function openColumnMenu(user: ReturnType<typeof userEvent.setup>, name: string) {
  await user.click(screen.getByRole("button", { name: `Column options for ${name}` }));
}

function visibleHeaderNames(): string[] {
  return screen.getAllByRole("columnheader")
    .filter((header) => header.getAttribute("aria-label") !== "Row headers")
    .map((header) => header.getAttribute("aria-label") ?? "");
}

function rect({
  left,
  right,
  top,
  bottom
}: {
  left: number;
  right: number;
  top: number;
  bottom: number;
}): DOMRect {
  return {
    x: left,
    y: top,
    left,
    right,
    top,
    bottom,
    width: right - left,
    height: bottom - top,
    toJSON: () => ({})
  } as DOMRect;
}
