import { fireEvent, render, screen } from "@testing-library/react";
import { describe, expect, it } from "vitest";
import { Grid } from "./Grid";
import { createFormulaEngine } from "../lib/formulaEngine";
import type { CellRange, SheetModel, WorkbookModel } from "../types";

describe("Grid", () => {
  it("selects full columns, rows, and the sheet from headers", () => {
    const sheet: SheetModel = {
      id: "sheet-1",
      name: "Data",
      rowCount: 4,
      columnCount: 3,
      cells: {},
      formats: {},
      columnWidths: {},
      rowHeights: {},
      comments: {},
      hyperlinks: {},
      validations: {},
      conditionalFormats: [],
      filters: [],
      charts: [],
      merges: [],
      protection: { isProtected: false, lockedCells: {}, unlockedCells: {} }
    };
    const workbook: WorkbookModel = {
      version: 1,
      activeSheetId: sheet.id,
      sheets: [sheet],
      namedRanges: []
    };
    const selections: CellRange[] = [];

    render(
      <Grid
        sheet={sheet}
        formulaEngine={createFormulaEngine(workbook)}
        selection={{ start: { row: 0, column: 0 }, end: { row: 0, column: 0 } }}
        editingCell={null}
        getCellFormat={() => undefined}
        onSelectionChange={(range) => selections.push(range)}
        onStartEdit={() => undefined}
        onEditValueChange={() => undefined}
        onCommitEdit={() => undefined}
        onCancelEdit={() => undefined}
        onPasteText={() => undefined}
        onKeyCommand={() => undefined}
      />
    );

    fireEvent.click(screen.getByRole("columnheader", { name: "Column B" }));
    fireEvent.click(screen.getByRole("rowheader", { name: "Row 3" }));
    fireEvent.click(screen.getByRole("button", { name: "Select sheet" }));

    expect(selections).toEqual([
      { start: { row: 0, column: 1 }, end: { row: 3, column: 1 } },
      { start: { row: 2, column: 0 }, end: { row: 2, column: 2 } },
      { start: { row: 0, column: 0 }, end: { row: 3, column: 2 } }
    ]);
  });

  it("does not render hidden rows or columns", () => {
    const sheet: SheetModel = {
      id: "sheet-1",
      name: "Data",
      rowCount: 4,
      columnCount: 3,
      cells: {
        A1: "Visible",
        B1: "Hidden column",
        A2: "Hidden row",
        C3: "Still visible"
      },
      formats: {},
      columnWidths: {},
      rowHeights: {},
      hiddenRows: { "1": true },
      hiddenColumns: { "1": true },
      comments: {},
      hyperlinks: {},
      validations: {},
      conditionalFormats: [],
      filters: [],
      charts: [],
      merges: [],
      protection: { isProtected: false, lockedCells: {}, unlockedCells: {} }
    };
    const workbook: WorkbookModel = {
      version: 1,
      activeSheetId: sheet.id,
      sheets: [sheet],
      namedRanges: []
    };

    render(
      <Grid
        sheet={sheet}
        formulaEngine={createFormulaEngine(workbook)}
        selection={{ start: { row: 0, column: 0 }, end: { row: 0, column: 0 } }}
        editingCell={null}
        getCellFormat={() => undefined}
        onSelectionChange={() => undefined}
        onStartEdit={() => undefined}
        onEditValueChange={() => undefined}
        onCommitEdit={() => undefined}
        onCancelEdit={() => undefined}
        onPasteText={() => undefined}
        onKeyCommand={() => undefined}
      />
    );

    expect(screen.getByRole("columnheader", { name: "Column A" })).toBeInTheDocument();
    expect(screen.queryByRole("columnheader", { name: "Column B" })).not.toBeInTheDocument();
    expect(screen.getByRole("columnheader", { name: "Column C" })).toBeInTheDocument();
    expect(screen.queryByRole("rowheader", { name: "Row 2" })).not.toBeInTheDocument();
    expect(screen.getByRole("gridcell", { name: "A1 Visible" })).toBeInTheDocument();
    expect(screen.queryByRole("gridcell", { name: "B1 Hidden column" })).not.toBeInTheDocument();
    expect(screen.queryByRole("gridcell", { name: "A2 Hidden row" })).not.toBeInTheDocument();
    expect(screen.getByRole("gridcell", { name: "C3 Still visible" })).toBeInTheDocument();
  });

  it("renders only a bounded viewport for large sheets and can scroll to distant rows", () => {
    const sheet: SheetModel = {
      id: "sheet-1",
      name: "Data",
      rowCount: 260,
      columnCount: 26,
      cells: {
        A1: "Top",
        A121: "Deep row"
      },
      formats: {},
      columnWidths: {},
      rowHeights: {},
      comments: {},
      hyperlinks: {},
      validations: {},
      conditionalFormats: [],
      filters: [],
      charts: [],
      merges: [],
      protection: { isProtected: false, lockedCells: {}, unlockedCells: {} }
    };
    const workbook: WorkbookModel = {
      version: 1,
      activeSheetId: sheet.id,
      sheets: [sheet],
      namedRanges: []
    };
    const selection: CellRange = {
      start: { row: 0, column: 0 },
      end: { row: 0, column: 0 }
    };

    render(
      <Grid
        sheet={sheet}
        formulaEngine={createFormulaEngine(workbook)}
        selection={selection}
        editingCell={null}
        getCellFormat={() => undefined}
        onSelectionChange={() => undefined}
        onStartEdit={() => undefined}
        onEditValueChange={() => undefined}
        onCommitEdit={() => undefined}
        onCancelEdit={() => undefined}
        onPasteText={() => undefined}
        onKeyCommand={() => undefined}
      />
    );

    const grid = screen.getByRole("grid", { name: "Spreadsheet grid" });
    Object.defineProperty(grid, "clientHeight", { configurable: true, value: 280 });

    expect(screen.getByRole("gridcell", { name: "A1 Top" })).toBeInTheDocument();
    expect(screen.getAllByRole("gridcell").length).toBeLessThan(1600);

    grid.scrollTop = 120 * 28;
    fireEvent.scroll(grid);

    expect(screen.getByRole("gridcell", { name: "A121 Deep row" })).toBeInTheDocument();
    expect(screen.getAllByRole("gridcell").length).toBeLessThan(1600);
  });

  it("marks frozen top-row and first-column cells for sticky pane styling", () => {
    const sheet: SheetModel = {
      id: "sheet-1",
      name: "Data",
      rowCount: 100,
      columnCount: 26,
      cells: {
        A1: "Frozen corner",
        B1: "Frozen row",
        A2: "Frozen column"
      },
      formats: {},
      columnWidths: {},
      rowHeights: {},
      comments: {},
      hyperlinks: {},
      validations: {},
      conditionalFormats: [],
      filters: [],
      charts: [],
      merges: [],
      protection: { isProtected: false, lockedCells: {}, unlockedCells: {} }
    };
    const workbook: WorkbookModel = {
      version: 1,
      activeSheetId: sheet.id,
      sheets: [sheet],
      namedRanges: []
    };

    render(
      <Grid
        sheet={sheet}
        formulaEngine={createFormulaEngine(workbook)}
        selection={{ start: { row: 0, column: 0 }, end: { row: 0, column: 0 } }}
        editingCell={null}
        freezeTopRow
        freezeFirstColumn
        getCellFormat={() => undefined}
        onSelectionChange={() => undefined}
        onStartEdit={() => undefined}
        onEditValueChange={() => undefined}
        onCommitEdit={() => undefined}
        onCancelEdit={() => undefined}
        onPasteText={() => undefined}
        onKeyCommand={() => undefined}
      />
    );

    expect(screen.getByRole("gridcell", { name: "A1 Frozen corner" })).toHaveClass(
      "frozen-top-row-cell",
      "frozen-first-column-cell"
    );
    expect(screen.getByRole("gridcell", { name: "B1 Frozen row" })).toHaveClass("frozen-top-row-cell");
    expect(screen.getByRole("gridcell", { name: "A2 Frozen column" })).toHaveClass("frozen-first-column-cell");
  });

  it("commits row and column size changes from resize handles", () => {
    const sheet: SheetModel = {
      id: "sheet-1",
      name: "Data",
      rowCount: 100,
      columnCount: 26,
      cells: {},
      formats: {},
      columnWidths: {},
      rowHeights: {},
      comments: {},
      hyperlinks: {},
      validations: {},
      conditionalFormats: [],
      filters: [],
      charts: [],
      merges: [],
      protection: { isProtected: false, lockedCells: {}, unlockedCells: {} }
    };
    const workbook: WorkbookModel = {
      version: 1,
      activeSheetId: sheet.id,
      sheets: [sheet],
      namedRanges: []
    };
    let resizedColumn: { column: number; width: number } | null = null;
    let resizedRow: { row: number; height: number } | null = null;

    render(
      <Grid
        sheet={sheet}
        formulaEngine={createFormulaEngine(workbook)}
        selection={{ start: { row: 0, column: 0 }, end: { row: 0, column: 0 } }}
        editingCell={null}
        getCellFormat={() => undefined}
        onSelectionChange={() => undefined}
        onStartEdit={() => undefined}
        onEditValueChange={() => undefined}
        onCommitEdit={() => undefined}
        onCancelEdit={() => undefined}
        onPasteText={() => undefined}
        onKeyCommand={() => undefined}
        onColumnResize={(column, width) => {
          resizedColumn = { column, width };
        }}
        onRowResize={(row, height) => {
          resizedRow = { row, height };
        }}
      />
    );

    fireEvent.mouseDown(screen.getByLabelText("Resize column A"), { clientX: 100 });
    fireEvent.mouseMove(window, { clientX: 132 });
    fireEvent.mouseUp(window);

    fireEvent.mouseDown(screen.getByLabelText("Resize row 1"), { clientY: 20 });
    fireEvent.mouseMove(window, { clientY: 34 });
    fireEvent.mouseUp(window);

    expect(resizedColumn).toEqual({ column: 0, width: 128 });
    expect(resizedRow).toEqual({ row: 0, height: 42 });
  });
});
