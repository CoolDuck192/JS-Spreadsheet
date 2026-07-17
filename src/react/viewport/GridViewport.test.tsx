import { act, fireEvent, render, screen, within } from "@testing-library/react";
import { useState } from "react";
import { afterEach, describe, expect, it, vi } from "vitest";
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
  onRegisterApi,
  interactionEventMode = "pointer",
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
  onRegisterApi?: GridViewportProps["onRegisterApi"];
  interactionEventMode?: GridViewportProps["interactionEventMode"];
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
        onRegisterApi={onRegisterApi}
        interactionEventMode={interactionEventMode}
        announce={`${viewportRows.length} rows loaded`}
      />
      <output data-testid={`${idPrefix}-last-interaction`}>{JSON.stringify(lastInteraction)}</output>
    </div>
  );
}

describe("GridViewport", () => {
  afterEach(() => {
    vi.restoreAllMocks();
    vi.unstubAllGlobals();
  });

  it("focuses a cell through the registered viewport API", () => {
    let api: Parameters<NonNullable<GridViewportProps["onRegisterApi"]>>[0] | null = null;
    render(
      <StatefulViewport
        onRegisterApi={(next) => {
          api = next;
        }}
      />
    );

    act(() => api?.focusCell("row/ada", "salary"));

    expect(screen.getByRole("gridcell", { name: "Ada Salary" })).toHaveFocus();
  });

  it("keeps imperative focus inside the invoking root when cell ids are duplicated", () => {
    let secondApi: Parameters<NonNullable<GridViewportProps["onRegisterApi"]>>[0] | null = null;
    render(
      <>
        <StatefulViewport idPrefix="duplicate" ariaLabel="First duplicate grid" />
        <StatefulViewport
          idPrefix="duplicate"
          ariaLabel="Second duplicate grid"
          onRegisterApi={(next) => {
            secondApi = next;
          }}
        />
      </>
    );
    const firstGrid = screen.getByRole("grid", { name: "First duplicate grid" });
    const secondGrid = screen.getByRole("grid", { name: "Second duplicate grid" });

    act(() => secondApi?.focusCell("row/ada", "salary"));

    expect(within(secondGrid).getByRole("gridcell", { name: "Ada Salary" })).toHaveFocus();
    expect(within(firstGrid).getByRole("gridcell", { name: "Ada Salary" })).not.toHaveFocus();
  });

  it("cancels a queued focus retry when the viewport unmounts", () => {
    const animationFrames = installAnimationFrameQueue();
    const virtualRows = createVirtualRows(1_000);
    let api: Parameters<NonNullable<GridViewportProps["onRegisterApi"]>>[0] | null = null;
    const { unmount } = render(
      <StatefulViewport
        idPrefix="unmount-focus"
        viewportRows={virtualRows}
        withRowHeaders={false}
        initialSelection={null}
        onRegisterApi={(next) => {
          api = next;
        }}
      />
    );

    act(() => api?.focusCell("virtual-row-999", "salary"));
    expect(animationFrames.pendingCount()).toBe(1);

    unmount();

    expect(animationFrames.pendingCount()).toBe(0);
  });

  it("cancels a stale focus retry when a rendered target supersedes it", () => {
    const animationFrames = installAnimationFrameQueue();
    const virtualRows = createVirtualRows(1_000);
    let api: Parameters<NonNullable<GridViewportProps["onRegisterApi"]>>[0] | null = null;
    render(
      <StatefulViewport
        idPrefix="superseded-focus"
        viewportRows={virtualRows}
        withRowHeaders={false}
        initialSelection={null}
        onRegisterApi={(next) => {
          api = next;
        }}
      />
    );

    act(() => api?.focusCell("virtual-row-999", "salary"));
    expect(animationFrames.pendingCount()).toBe(1);

    act(() => api?.focusCell("virtual-row-0", "salary"));

    expect(screen.getByRole("gridcell", { name: "Virtual row 0 Salary" })).toHaveFocus();
    expect(animationFrames.pendingCount()).toBe(0);
  });

  it("focuses a virtualized target on the queued retry after ensure-visible renders it", () => {
    const animationFrames = installAnimationFrameQueue();
    const virtualRows = createVirtualRows(1_000);
    let api: Parameters<NonNullable<GridViewportProps["onRegisterApi"]>>[0] | null = null;
    render(
      <StatefulViewport
        idPrefix="virtual-focus"
        ariaLabel="Virtual focus grid"
        viewportRows={virtualRows}
        withRowHeaders={false}
        initialSelection={null}
        onRegisterApi={(next) => {
          api = next;
        }}
      />
    );
    const grid = screen.getByRole("grid", { name: "Virtual focus grid" });
    Object.defineProperties(grid, {
      clientHeight: { configurable: true, value: 280 },
      clientWidth: { configurable: true, value: 400 },
      scrollTop: { configurable: true, writable: true, value: 0 },
      scrollLeft: { configurable: true, writable: true, value: 0 }
    });

    expect(screen.queryByRole("gridcell", { name: "Virtual row 900 Salary" })).not.toBeInTheDocument();
    act(() => api?.focusCell("virtual-row-900", "salary"));
    expect(grid.scrollTop).toBeGreaterThan(20_000);
    expect(screen.queryByRole("gridcell", { name: "Virtual row 900 Salary" })).not.toBeInTheDocument();

    fireEvent.scroll(grid);
    const target = screen.getByRole("gridcell", { name: "Virtual row 900 Salary" });
    expect(target).not.toHaveFocus();

    act(() => animationFrames.flush());

    expect(target).toHaveFocus();
  });

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

  it("keeps pinned columns in their natural grid slots", () => {
    const pinnedColumns: readonly GridViewportColumn[] = [
      { ...columns[0], pinned: "left" },
      { ...columns[1], pinned: "left" },
      { ...columns[2], pinned: "right" }
    ];
    render(<StatefulViewport viewportColumns={pinnedColumns} />);

    expect(screen.getByRole("row", { name: "Column headers" })).toHaveStyle({
      display: "grid",
      gridTemplateColumns: "56px 120px 100px 90px"
    });
    expect(screen.getByRole("columnheader", { name: "Name" })).toHaveStyle({
      position: "sticky",
      gridColumn: "2",
      left: "56px"
    });
    expect(screen.getByRole("columnheader", { name: "Salary" })).toHaveStyle({
      position: "sticky",
      gridColumn: "3",
      left: "176px"
    });
    expect(screen.getByRole("columnheader", { name: "Active" })).toHaveStyle({
      position: "sticky",
      gridColumn: "4",
      right: "0px"
    });
    expect(screen.getByRole("gridcell", { name: "Ada Active" })).toHaveStyle({
      gridColumn: "4",
      right: "0px"
    });
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

  it("captures pointer drags and ends them when release happens outside the grid", () => {
    render(<StatefulViewport />);
    const adaName = screen.getByRole("gridcell", { name: "Ada Name" });
    const graceSalary = screen.getByRole("gridcell", { name: "Grace Salary" });
    const linActive = screen.getByRole("gridcell", { name: "Lin Active" });
    const setPointerCapture = vi.fn();
    Object.defineProperty(adaName, "setPointerCapture", { configurable: true, value: setPointerCapture });

    fireEvent.pointerDown(adaName, { pointerId: 7 });
    expect(setPointerCapture).toHaveBeenCalledWith(7);
    fireEvent.pointerEnter(linActive);
    const releasedSelection = readLastInteraction("people").selection;
    fireEvent.pointerUp(window, { pointerId: 7 });
    fireEvent.pointerEnter(graceSalary);

    expect(readLastInteraction("people").selection).toEqual(releasedSelection);
  });

  it("ends legacy mouse drags when mouseup happens outside the grid", () => {
    render(<StatefulViewport interactionEventMode="mouse" />);
    const adaName = screen.getByRole("gridcell", { name: "Ada Name" });
    const graceSalary = screen.getByRole("gridcell", { name: "Grace Salary" });
    const linActive = screen.getByRole("gridcell", { name: "Lin Active" });

    fireEvent.mouseDown(adaName);
    fireEvent.mouseEnter(linActive);
    const releasedSelection = readLastInteraction("people").selection;
    fireEvent.mouseUp(window);
    fireEvent.mouseEnter(graceSalary);

    expect(readLastInteraction("people").selection).toEqual(releasedSelection);
  });

  it("keeps scroll position when selecting and editing a pinned cell", () => {
    const pinnedRows: readonly GridViewportRow[] = [
      { ...rows[0], pinned: "top" },
      ...rows.slice(1)
    ];
    const pinnedColumns: readonly GridViewportColumn[] = [
      { ...columns[0], pinned: "left" },
      ...columns.slice(1)
    ];
    render(<StatefulViewport viewportRows={pinnedRows} viewportColumns={pinnedColumns} />);
    const grid = screen.getByRole("grid", { name: "People grid" });
    Object.defineProperties(grid, {
      clientHeight: { configurable: true, value: 280 },
      clientWidth: { configurable: true, value: 400 },
      scrollTop: { configurable: true, writable: true, value: 240 },
      scrollLeft: { configurable: true, writable: true, value: 180 }
    });
    const pinnedCell = screen.getByRole("gridcell", { name: "Ada Name" });

    fireEvent.pointerDown(pinnedCell);
    expect({ top: grid.scrollTop, left: grid.scrollLeft }).toEqual({ top: 240, left: 180 });

    fireEvent.doubleClick(pinnedCell);
    const overlay = screen.getByRole("textbox", { name: "Edit Ada Name" }).parentElement;
    expect(overlay).toHaveStyle({ top: "268px", left: "236px" });
    expect({ top: grid.scrollTop, left: grid.scrollLeft }).toEqual({ top: 240, left: 180 });
  });

  it("scrolls upward targets below the pinned-row band", () => {
    const pinnedRows: readonly GridViewportRow[] = Array.from({ length: 100 }, (_, index) => ({
      id: `pinned-band-row-${index}`,
      label: `Pinned band row ${index}`,
      height: index === 0 ? 32 : 28,
      kind: "data",
      ariaRowIndex: index + 2,
      ...(index === 0 ? { pinned: "top" as const } : {})
    }));
    let api: Parameters<NonNullable<GridViewportProps["onRegisterApi"]>>[0] | null = null;
    render(
      <StatefulViewport
        ariaLabel="Pinned band grid"
        viewportRows={pinnedRows}
        onRegisterApi={(next) => { api = next; }}
        initialSelection={null}
      />
    );
    const grid = screen.getByRole("grid", { name: "Pinned band grid" });
    Object.defineProperties(grid, {
      clientHeight: { configurable: true, value: 280 },
      clientWidth: { configurable: true, value: 400 },
      scrollTop: { configurable: true, writable: true, value: 1_500 },
      scrollLeft: { configurable: true, writable: true, value: 0 }
    });

    act(() => api?.ensureCellVisible("pinned-band-row-50", "salary"));

    const targetStart = 32 + 49 * 28;
    expect(grid.scrollTop).toBe(targetStart - 32);
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
    expect(editor.parentElement).toHaveClass(
      "js-spreadsheet-grid__cell",
      "js-spreadsheet-grid__editing-cell"
    );
    expect(editor.parentElement).not.toHaveClass("js-spreadsheet-cell", "js-spreadsheet-editing-cell");
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

function createVirtualRows(count: number): readonly GridViewportRow[] {
  return Array.from({ length: count }, (_, index) => ({
    id: `virtual-row-${index}`,
    label: `Virtual row ${index}`,
    height: 28,
    kind: "data" as const,
    ariaRowIndex: index + 2
  }));
}

function installAnimationFrameQueue() {
  let nextId = 1;
  const callbacks = new Map<number, FrameRequestCallback>();
  vi.stubGlobal("requestAnimationFrame", (callback: FrameRequestCallback) => {
    const id = nextId;
    nextId += 1;
    callbacks.set(id, callback);
    return id;
  });
  vi.stubGlobal("cancelAnimationFrame", (id: number) => {
    callbacks.delete(id);
  });
  return {
    pendingCount: () => callbacks.size,
    flush() {
      const pending = [...callbacks.values()];
      callbacks.clear();
      for (const callback of pending) {
        callback(performance.now());
      }
    }
  };
}
