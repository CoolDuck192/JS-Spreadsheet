import { fireEvent, render, screen, within } from "@testing-library/react";
import { useState } from "react";
import { describe, expect, it } from "vitest";
import type { TableSelection } from "../../table/core/types";
import { GridViewport } from "./GridViewport";
import type {
  GridEditorState,
  GridViewportCell,
  GridViewportColumn,
  GridViewportInteraction,
  GridViewportProps,
  GridViewportRow
} from "./types";

const rows: readonly GridViewportRow[] = [
  { id: "row/ada", label: "Ada", height: 32, kind: "data", ariaRowIndex: 2 },
  { id: "row/grace", label: "Grace", height: 36, kind: "data", ariaRowIndex: 3 },
  { id: "row/lin", label: "Lin", height: 40, kind: "aggregate", ariaRowIndex: 4 }
];
const columns: readonly GridViewportColumn[] = [
  { id: "full name", label: "Name", width: 120, minWidth: 80, maxWidth: 240 },
  { id: "salary", label: "Salary", width: 100, minWidth: 80, maxWidth: 180 },
  { id: "active", label: "Active", width: 90, minWidth: 70, maxWidth: 140 }
];

function StatefulViewport({
  idPrefix = "people",
  ariaLabel = "People grid",
  viewportRows = rows,
  viewportColumns = columns,
  ariaRowCount = viewportRows.length + 1,
  ariaColumnCount = viewportColumns.length + 1,
  withRowHeaders = true,
  onColumnHeaderContextMenu,
  initialSelection = {
    anchor: { rowId: viewportRows[0]?.id ?? "", columnId: viewportColumns[0]?.id ?? "" },
    focus: { rowId: viewportRows[0]?.id ?? "", columnId: viewportColumns[0]?.id ?? "" }
  }
}: {
  idPrefix?: string;
  ariaLabel?: string;
  viewportRows?: readonly GridViewportRow[];
  viewportColumns?: readonly GridViewportColumn[];
  ariaRowCount?: number;
  ariaColumnCount?: number;
  withRowHeaders?: boolean;
  onColumnHeaderContextMenu?: GridViewportProps["onColumnHeaderContextMenu"];
  initialSelection?: TableSelection | null;
}) {
  const [selection, setSelection] = useState<TableSelection | null>(initialSelection);
  const [editing, setEditing] = useState<GridEditorState | null>(null);
  const [lastInteraction, setLastInteraction] = useState<GridViewportInteraction | null>(null);

  function onInteraction(interaction: GridViewportInteraction) {
    setLastInteraction(interaction);
    if (interaction.type === "selection-change") {
      setSelection(interaction.selection);
    } else if (interaction.type === "edit-start") {
      setEditing({ ...interaction.cell, rawText: interaction.initialRawText });
    } else if (interaction.type === "edit-change") {
      setEditing((current) => (current ? { ...current, rawText: interaction.rawText } : current));
    } else if (interaction.type === "edit-commit" || interaction.type === "edit-cancel") {
      setEditing(null);
    }
  }

  return (
    <div data-testid={`${idPrefix}-harness`}>
      <GridViewport
        idPrefix={idPrefix}
        ariaLabel={ariaLabel}
        rows={viewportRows}
        ariaRowCount={ariaRowCount}
        columns={viewportColumns}
        ariaColumnCount={ariaColumnCount}
        getCell={(rowId, columnId) => createCell(rowId, columnId, viewportRows, viewportColumns)}
        selection={selection}
        editing={editing}
        onInteraction={onInteraction}
        renderRowHeader={withRowHeaders ? (row) => row.label : undefined}
        onColumnHeaderContextMenu={onColumnHeaderContextMenu}
        announce={`${viewportRows.length} rows loaded`}
      />
      <output data-testid={`${idPrefix}-last-interaction`}>{JSON.stringify(lastInteraction)}</output>
    </div>
  );
}

describe("GridViewport", () => {
  it("renders correct grid, row, header, and cell roles with one-based indexes", () => {
    render(<StatefulViewport />);
    const grid = screen.getByRole("grid", { name: "People grid" });

    expect(grid).toHaveAttribute("data-viewport-kernel", "shared");
    expect(grid).toHaveAttribute("aria-rowcount", "4");
    expect(grid).toHaveAttribute("aria-colcount", "4");
    expect(screen.getByRole("row", { name: "Column headers" })).toHaveAttribute("aria-rowindex", "1");
    expect(screen.getByRole("columnheader", { name: "Name" })).toHaveAttribute("aria-colindex", "2");
    expect(screen.getByRole("row", { name: "Row Grace" })).toHaveAttribute("aria-rowindex", "3");
    expect(screen.getByRole("rowheader", { name: "Grace" })).toHaveAttribute("aria-colindex", "1");
    expect(screen.getByRole("gridcell", { name: "Grace Salary" })).toHaveAttribute("aria-colindex", "3");
    expect(screen.getByRole("gridcell", { name: "Grace Salary" })).toHaveAttribute("aria-readonly", "false");
    expect(screen.getByRole("gridcell", { name: "Grace Active" })).toHaveAttribute("aria-readonly", "true");
    expect(screen.getByRole("gridcell", { name: "Grace Salary" })).toHaveAttribute("aria-invalid", "true");
  });

  it("forwards native context-menu events from column headers", () => {
    const contexts: Array<{ column: string; x: number; y: number }> = [];
    render(
      <StatefulViewport
        onColumnHeaderContextMenu={(column, event) => {
          contexts.push({ column: column.id, x: event.clientX, y: event.clientY });
        }}
      />
    );

    fireEvent.contextMenu(screen.getByRole("columnheader", { name: "Salary" }), {
      clientX: 240,
      clientY: 60
    });

    expect(contexts).toEqual([{ column: "salary", x: 240, y: 60 }]);
  });

  it("uses explicit offset row indexes and minus one for an unknown total", () => {
    const offsetRows: readonly GridViewportRow[] = [
      { ...rows[0], ariaRowIndex: 102 },
      { ...rows[1], ariaRowIndex: 103 }
    ];
    render(
      <StatefulViewport
        ariaLabel="Remote grid"
        viewportRows={offsetRows}
        ariaRowCount={-1}
      />
    );

    expect(screen.getByRole("grid", { name: "Remote grid" })).toHaveAttribute("aria-rowcount", "-1");
    expect(screen.getByRole("row", { name: "Row Ada" })).toHaveAttribute("aria-rowindex", "102");
    expect(screen.getByRole("row", { name: "Row Grace" })).toHaveAttribute("aria-rowindex", "103");
  });

  it("selects by click and extends by shift-click and pointer drag", () => {
    render(<StatefulViewport />);
    const adaName = screen.getByRole("gridcell", { name: "Ada Name" });
    const graceSalary = screen.getByRole("gridcell", { name: "Grace Salary" });
    const linActive = screen.getByRole("gridcell", { name: "Lin Active" });

    fireEvent.click(graceSalary);
    expect(readLastInteraction("people")).toEqual({
      type: "selection-change",
      selection: {
        anchor: { rowId: "row/grace", columnId: "salary" },
        focus: { rowId: "row/grace", columnId: "salary" }
      }
    });

    fireEvent.click(linActive, { shiftKey: true });
    expect(readLastInteraction("people").selection).toEqual({
      anchor: { rowId: "row/grace", columnId: "salary" },
      focus: { rowId: "row/lin", columnId: "active" }
    });

    fireEvent.pointerDown(adaName);
    fireEvent.pointerEnter(linActive);
    fireEvent.pointerUp(screen.getByRole("grid", { name: "People grid" }));
    expect(readLastInteraction("people").selection).toEqual({
      anchor: { rowId: "row/ada", columnId: "full name" },
      focus: { rowId: "row/lin", columnId: "active" }
    });
  });

  it("navigates by keyboard and keeps exactly one roving tab stop", () => {
    render(<StatefulViewport />);
    const grid = screen.getByRole("grid", { name: "People grid" });
    const adaName = screen.getByRole("gridcell", { name: "Ada Name" });
    adaName.focus();

    fireEvent.keyDown(adaName, { key: "ArrowRight" });
    expect(screen.getByRole("gridcell", { name: "Ada Salary" })).toHaveAttribute("tabindex", "0");
    fireEvent.keyDown(grid, { key: "ArrowDown", shiftKey: true });
    expect(readLastInteraction("people").selection).toEqual({
      anchor: { rowId: "row/ada", columnId: "salary" },
      focus: { rowId: "row/grace", columnId: "salary" }
    });
    fireEvent.keyDown(grid, { key: "End" });
    expect(screen.getByRole("gridcell", { name: "Grace Active" })).toHaveAttribute("tabindex", "0");
    expect(grid.querySelectorAll('[role="gridcell"][tabindex="0"]')).toHaveLength(1);
  });

  it("starts edit with F2 or printable text and commits with Enter or Tab", () => {
    render(<StatefulViewport />);
    const adaName = screen.getByRole("gridcell", { name: "Ada Name" });
    adaName.focus();
    fireEvent.keyDown(adaName, { key: "F2" });

    const editor = screen.getByRole("textbox", { name: "Edit Ada Name" });
    expect(editor).toHaveValue("Ada Name");
    expect(editor.parentElement).toHaveAttribute("data-grid-editor-overlay", "true");
    expect(editor.parentElement).toHaveStyle({ top: "28px", left: "56px", width: "120px", height: "32px" });
    fireEvent.change(editor, { target: { value: "Ada Lovelace" } });
    fireEvent.keyDown(editor, { key: "Enter" });
    expect(readLastInteraction("people")).toEqual({
      type: "edit-commit",
      cell: { rowId: "row/ada", columnId: "full name" },
      rawText: "Ada Lovelace",
      move: "down"
    });

    const active = screen.getByRole("gridcell", { name: "Ada Name" });
    fireEvent.keyDown(active, { key: "Z" });
    const printableEditor = screen.getByRole("textbox", { name: "Edit Ada Name" });
    expect(printableEditor).toHaveValue("Z");
    fireEvent.keyDown(printableEditor, { key: "Tab" });
    expect(readLastInteraction("people")).toMatchObject({ type: "edit-commit", rawText: "Z", move: "right" });
  });

  it("cancels edit with Escape and restores the active cell focus", () => {
    render(<StatefulViewport />);
    const cell = screen.getByRole("gridcell", { name: "Ada Name" });
    cell.focus();
    fireEvent.keyDown(cell, { key: "F2" });
    const editor = screen.getByRole("textbox", { name: "Edit Ada Name" });
    fireEvent.change(editor, { target: { value: "discard" } });
    fireEvent.keyDown(editor, { key: "Escape" });

    expect(screen.queryByRole("textbox", { name: "Edit Ada Name" })).not.toBeInTheDocument();
    expect(readLastInteraction("people")).toEqual({ type: "edit-cancel" });
    expect(screen.getByRole("gridcell", { name: "Ada Name" })).toHaveFocus();
  });

  it("emits copy and plain-text paste without mutating displayed data", () => {
    render(<StatefulViewport />);
    const grid = screen.getByRole("grid", { name: "People grid" });
    fireEvent.keyDown(grid, { key: "c", ctrlKey: true });
    expect(readLastInteraction("people").type).toBe("copy");

    fireEvent.paste(grid, {
      clipboardData: { getData: (type: string) => (type === "text/plain" ? "A\tB\n1\t2" : "") }
    });

    expect(readLastInteraction("people")).toEqual({ type: "paste", text: "A\tB\n1\t2" });
    expect(screen.getByRole("gridcell", { name: "Ada Name" })).toHaveTextContent("Ada Name");
  });

  it("keeps native editor copy and paste isolated from grid clipboard interactions", () => {
    render(<StatefulViewport />);
    const cell = screen.getByRole("gridcell", { name: "Ada Name" });
    fireEvent.keyDown(cell, { key: "F2" });
    const editor = screen.getByRole("textbox", { name: "Edit Ada Name" });
    expect(readLastInteraction("people").type).toBe("edit-start");

    const copyAllowed = fireEvent.copy(editor, {
      clipboardData: { getData: () => "", setData: () => undefined }
    });
    const pasteAllowed = fireEvent.paste(editor, {
      clipboardData: { getData: () => "native editor text", setData: () => undefined }
    });

    expect(copyAllowed).toBe(true);
    expect(pasteAllowed).toBe(true);
    expect(readLastInteraction("people").type).toBe("edit-start");
  });

  it("renders only the two-axis virtual window", () => {
    const manyRows = Array.from({ length: 1_000 }, (_, index): GridViewportRow => ({
      id: `row-${index}`,
      label: `Row ${index}`,
      height: 28,
      kind: "data",
      ariaRowIndex: index + 2
    }));
    const manyColumns = Array.from({ length: 100 }, (_, index): GridViewportColumn => ({
      id: `column-${index}`,
      label: `Column ${index}`,
      width: 80,
      minWidth: 60,
      maxWidth: 160
    }));
    render(
      <StatefulViewport
        ariaLabel="Large shared grid"
        viewportRows={manyRows}
        viewportColumns={manyColumns}
        ariaRowCount={1_001}
        ariaColumnCount={100}
        withRowHeaders={false}
        initialSelection={null}
      />
    );
    const grid = screen.getByRole("grid", { name: "Large shared grid" });
    Object.defineProperties(grid, {
      clientHeight: { configurable: true, value: 280 },
      clientWidth: { configurable: true, value: 400 },
      scrollTop: { configurable: true, writable: true, value: 500 * 28 },
      scrollLeft: { configurable: true, writable: true, value: 50 * 80 }
    });
    fireEvent.scroll(grid);

    expect(screen.getByRole("gridcell", { name: "Row 500 Column 50" })).toBeInTheDocument();
    expect(screen.getAllByRole("gridcell").length).toBeLessThan(200);
    expect(screen.queryByRole("gridcell", { name: "Row 0 Column 0" })).not.toBeInTheDocument();
  });

  it("falls back to grid focus while scrolling and restores the off-screen roving cell", () => {
    const focusRows: readonly GridViewportRow[] = [
      { id: "focus-row", label: "Focus row", height: 28, kind: "data", ariaRowIndex: 2 }
    ];
    const focusColumns = Array.from({ length: 100 }, (_, index): GridViewportColumn => ({
      id: `focus-column-${index}`,
      label: `Focus column ${index}`,
      width: 80,
      minWidth: 60,
      maxWidth: 120
    }));
    render(
      <StatefulViewport
        ariaLabel="Focus grid"
        viewportRows={focusRows}
        viewportColumns={focusColumns}
        ariaRowCount={2}
        ariaColumnCount={100}
        withRowHeaders={false}
      />
    );
    const grid = screen.getByRole("grid", { name: "Focus grid" });
    Object.defineProperties(grid, {
      clientHeight: { configurable: true, value: 100 },
      clientWidth: { configurable: true, value: 240 },
      scrollTop: { configurable: true, writable: true, value: 0 },
      scrollLeft: { configurable: true, writable: true, value: 0 }
    });
    const first = screen.getByRole("gridcell", { name: "Focus row Focus column 0" });
    first.focus();

    fireEvent.keyDown(first, { key: "End" });

    expect(grid.scrollLeft).toBeGreaterThan(7_000);
    expect(grid).toHaveFocus();
    fireEvent.scroll(grid);
    expect(screen.getByRole("gridcell", { name: "Focus row Focus column 99" })).toHaveFocus();
    expect(grid.querySelectorAll('[role="gridcell"][tabindex="0"]')).toHaveLength(1);
  });

  it("uses instance-prefixed ids for cells and live regions", () => {
    render(
      <>
        <StatefulViewport idPrefix="first" ariaLabel="First grid" />
        <StatefulViewport idPrefix="second" ariaLabel="Second grid" />
      </>
    );
    const first = screen.getByTestId("first-harness");
    const second = screen.getByTestId("second-harness");
    const firstCellId = within(first).getByRole("gridcell", { name: "Ada Name" }).id;
    const secondCellId = within(second).getByRole("gridcell", { name: "Ada Name" }).id;
    const firstStatusId = first.querySelector('[aria-live="polite"]')!.id;
    const secondStatusId = second.querySelector('[aria-live="polite"]')!.id;

    expect(firstCellId).toMatch(/^first-/);
    expect(secondCellId).toMatch(/^second-/);
    expect(firstCellId).not.toBe(secondCellId);
    expect(firstStatusId).toMatch(/^first-/);
    expect(secondStatusId).toMatch(/^second-/);
    expect(firstStatusId).not.toBe(secondStatusId);
  });
});

function createCell(
  rowId: string,
  columnId: string,
  viewportRows: readonly GridViewportRow[],
  viewportColumns: readonly GridViewportColumn[]
): GridViewportCell {
  const row = viewportRows.find((candidate) => candidate.id === rowId)!;
  const column = viewportColumns.find((candidate) => candidate.id === columnId)!;
  const label = `${row.label} ${column.label}`;
  return {
    ref: { rowId, columnId },
    ariaLabel: label,
    displayValue: label,
    editable: columnId !== "active",
    invalid: rowId === "row/grace" && columnId === "salary"
  };
}

function readLastInteraction(idPrefix: string): any {
  return JSON.parse(screen.getByTestId(`${idPrefix}-last-interaction`).textContent ?? "null");
}
