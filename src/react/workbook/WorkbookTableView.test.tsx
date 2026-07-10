import { render, screen, waitFor } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import { describe, expect, it, vi } from "vitest";
import { createWorkbookSession } from "../../core/workbook/WorkbookSession";
import { createBlankWorkbook, getCellContent } from "../../lib/workbook";
import type { StructuredTable, WorkbookModel } from "../../types";
import { WorkbookTableView } from "./WorkbookTableView";

describe("WorkbookTableView", () => {
  it("stays live with its workbook session and edits through the shared table adapter", async () => {
    const parent = createWorkbookSession({ workbook: workbookFixture() });
    const table = parent.table("table-people");
    render(<WorkbookTableView session={table} onClose={() => undefined} />);

    expect(screen.getByRole("gridcell", { name: "row-ada Name" })).toHaveTextContent("Ada");
    parent.dispatch({ type: "cell.set", sheetId: "sheet-1", address: "A2", input: "Updated" });
    await waitFor(() => {
      expect(screen.getByRole("gridcell", { name: "row-ada Name" })).toHaveTextContent("Updated");
    });

    await table.dispatch({
      type: "edit-cells",
      edits: [{ rowId: "row-ada", columnId: "column-score", rawText: "42" }]
    });
    expect(getCellContent(parent.getSnapshot().workbook, "sheet-1", "B2")).toBe(42);
    parent.destroy();
  });

  it("opens the current stable table cell and closes independently", async () => {
    const user = userEvent.setup();
    const parent = createWorkbookSession({
      workbook: workbookFixture(),
      selection: { start: { row: 1, column: 1 }, end: { row: 1, column: 1 } }
    });
    const onOpen = vi.fn();
    const onClose = vi.fn();
    render(<WorkbookTableView session={parent.table("table-people")} onClose={onClose} onOpenInSpreadsheet={onOpen} />);

    await user.click(screen.getByRole("button", { name: "Open in Spreadsheet" }));
    expect(onOpen).toHaveBeenCalledWith({ rowId: "row-ada", columnId: "column-score" });
    await user.click(screen.getByRole("button", { name: "Close table view" }));
    expect(onClose).toHaveBeenCalledTimes(1);
    parent.destroy();
  });

  it("keeps simultaneous workbook views isolated", () => {
    const first = createWorkbookSession({ workbook: workbookFixture({ A2: "First" }) });
    const second = createWorkbookSession({ workbook: workbookFixture({ A2: "Second" }) });
    render(
      <>
        <WorkbookTableView session={first.table("table-people")} onClose={() => undefined} />
        <WorkbookTableView session={second.table("table-people")} onClose={() => undefined} />
      </>
    );

    expect(screen.getAllByRole("gridcell", { name: "row-ada Name" }).map((cell) => cell.textContent))
      .toEqual(["First", "Second"]);
    first.destroy();
    second.destroy();
  });

  it("keeps Close available when the table is deleted", async () => {
    const user = userEvent.setup();
    const parent = createWorkbookSession({ workbook: workbookFixture() });
    const onClose = vi.fn();
    render(<WorkbookTableView session={parent.table("table-people")} onClose={onClose} onOpenInSpreadsheet={() => undefined} />);
    parent.dispatch({ type: "table.convertToRange", tableId: "table-people" });

    expect(await screen.findByRole("alert")).toHaveTextContent("Structured table does not exist");
    expect(screen.getByRole("button", { name: "Open in Spreadsheet" })).toBeDisabled();
    await user.click(screen.getByRole("button", { name: "Close table view" }));
    expect(onClose).toHaveBeenCalledTimes(1);
    parent.destroy();
  });
});

function workbookFixture(cellOverrides: Record<string, string | number> = {}): WorkbookModel {
  const workbook = createBlankWorkbook();
  const table: StructuredTable = {
    id: "table-people",
    name: "People",
    sheetId: "sheet-1",
    range: { start: { row: 0, column: 0 }, end: { row: 2, column: 1 } },
    headerRow: true,
    totalsRow: false,
    columns: [
      { id: "column-name", name: "Name", sheetColumn: 0, dataType: "text" },
      { id: "column-score", name: "Score", sheetColumn: 1, dataType: "number" }
    ],
    rowIds: ["row-ada", "row-grace"]
  };
  return {
    ...workbook,
    activeSheetId: "sheet-1",
    tables: [table],
    sheets: [{
      ...workbook.sheets[0],
      id: "sheet-1",
      cells: {
        ...workbook.sheets[0].cells,
        A1: "Name", B1: "Score", A2: "Ada", B2: 10, A3: "Grace", B3: 20,
        ...cellOverrides
      }
    }]
  };
}
