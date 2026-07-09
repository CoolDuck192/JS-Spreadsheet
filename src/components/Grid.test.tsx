import { fireEvent, render, screen, within } from "@testing-library/react";
import { describe, expect, it, vi } from "vitest";
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

  it("selects a span of columns by dragging across column headers", () => {
    const { sheet, workbook } = createFixtureSheet();
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

    fireEvent.mouseDown(screen.getByRole("columnheader", { name: "Column A" }));
    fireEvent.mouseEnter(screen.getByRole("columnheader", { name: "Column C" }));
    fireEvent.mouseUp(screen.getByRole("grid", { name: "Spreadsheet grid" }));

    expect(selections.at(-1)).toEqual({ start: { row: 0, column: 0 }, end: { row: 3, column: 2 } });
  });

  it("extends a cell selection with shift-click", () => {
    const { sheet, workbook } = createFixtureSheet();
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

    fireEvent.click(screen.getByRole("gridcell", { name: "B3" }), { shiftKey: true });

    expect(selections.at(-1)).toEqual({ start: { row: 0, column: 0 }, end: { row: 2, column: 1 } });
  });

  it("reports fill targets from the selection fill handle, including shrink", () => {
    const { sheet, workbook } = createFixtureSheet();
    const fills: Array<{ source: CellRange; target: CellRange }> = [];

    render(
      <Grid
        sheet={sheet}
        formulaEngine={createFormulaEngine(workbook)}
        selection={{ start: { row: 0, column: 0 }, end: { row: 1, column: 0 } }}
        editingCell={null}
        getCellFormat={() => undefined}
        onSelectionChange={() => undefined}
        onStartEdit={() => undefined}
        onEditValueChange={() => undefined}
        onCommitEdit={() => undefined}
        onCancelEdit={() => undefined}
        onPasteText={() => undefined}
        onKeyCommand={() => undefined}
        onAutoFill={(source, target) => fills.push({ source, target })}
      />
    );

    const grid = screen.getByRole("grid", { name: "Spreadsheet grid" });
    const handle = screen.getByRole("button", { name: "AutoFill selection" });

    fireEvent.mouseDown(handle);
    fireEvent.mouseEnter(screen.getByRole("gridcell", { name: "A4" }));
    fireEvent.mouseUp(grid);

    expect(fills).toEqual([
      {
        source: { start: { row: 0, column: 0 }, end: { row: 1, column: 0 } },
        target: { start: { row: 0, column: 0 }, end: { row: 3, column: 0 } }
      }
    ]);
  });

  it("extends fill targets in all four directions from the handle", () => {
    const { sheet, workbook } = createFixtureSheet();
    const fills: Array<{ source: CellRange; target: CellRange }> = [];

    render(
      <Grid
        sheet={sheet}
        formulaEngine={createFormulaEngine(workbook)}
        selection={{ start: { row: 1, column: 1 }, end: { row: 2, column: 1 } }}
        editingCell={null}
        getCellFormat={() => undefined}
        onSelectionChange={() => undefined}
        onStartEdit={() => undefined}
        onEditValueChange={() => undefined}
        onCommitEdit={() => undefined}
        onCancelEdit={() => undefined}
        onPasteText={() => undefined}
        onKeyCommand={() => undefined}
        onAutoFill={(source, target) => fills.push({ source, target })}
      />
    );

    const grid = screen.getByRole("grid", { name: "Spreadsheet grid" });

    // Fill left: source B2:B3 dragged to A3 extends to A2:B3.
    fireEvent.mouseDown(screen.getByRole("button", { name: "AutoFill selection" }));
    fireEvent.mouseEnter(screen.getByRole("gridcell", { name: "A3" }));
    fireEvent.mouseUp(grid);

    // Fill up: source B2:B3 dragged to B1 extends to B1:B3.
    fireEvent.mouseDown(screen.getByRole("button", { name: "AutoFill selection" }));
    fireEvent.mouseEnter(screen.getByRole("gridcell", { name: "B1" }));
    fireEvent.mouseUp(grid);

    expect(fills).toEqual([
      {
        source: { start: { row: 1, column: 1 }, end: { row: 2, column: 1 } },
        target: { start: { row: 1, column: 0 }, end: { row: 2, column: 1 } }
      },
      {
        source: { start: { row: 1, column: 1 }, end: { row: 2, column: 1 } },
        target: { start: { row: 0, column: 1 }, end: { row: 2, column: 1 } }
      }
    ]);
  });

  it("auto-fits columns and rows when their resize handles are double-clicked", () => {
    const { sheet, workbook } = createFixtureSheet();
    const columnAutoFits: number[] = [];
    const rowAutoFits: number[] = [];

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
        onColumnAutoFit={(column) => columnAutoFits.push(column)}
        onRowAutoFit={(row) => rowAutoFits.push(row)}
      />
    );

    fireEvent.doubleClick(screen.getByLabelText("Resize column A"));
    fireEvent.doubleClick(screen.getByLabelText("Resize row 1"));

    expect(columnAutoFits).toEqual([0]);
    expect(rowAutoFits).toEqual([0]);
  });

  it("uses validation candidates consistently for formula badges", () => {
    const sheet: SheetModel = {
      ...createFixtureSheet().sheet,
      rowCount: 3,
      columnCount: 1,
      cells: {
        A1: "=5+5",
        A2: "=5+6",
        A3: '="Open"'
      },
      validations: {
        A1: { type: "number", min: 1, max: 10 },
        A2: { type: "number", min: 1, max: 10 },
        A3: { type: "list", values: ["Open", "Closed"] }
      }
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
        getCellValidation={(address) => sheet.validations[address]}
        onSelectionChange={() => undefined}
        onStartEdit={() => undefined}
        onEditValueChange={() => undefined}
        onCommitEdit={() => undefined}
        onCancelEdit={() => undefined}
        onPasteText={() => undefined}
        onKeyCommand={() => undefined}
      />
    );

    expect(screen.getByRole("gridcell", { name: "A1 10" })).not.toHaveClass("invalid-validation-cell");
    expect(screen.getByRole("gridcell", { name: "A2 11" })).toHaveClass("invalid-validation-cell");
    expect(screen.getByRole("gridcell", { name: "A3 Open" })).toHaveClass("invalid-validation-cell");
  });

  it("collects AutoFilter choices only while open and searches the full capped distinct set", () => {
    const cells: SheetModel["cells"] = { A1: "Region" };
    for (let index = 1; index <= 201; index += 1) {
      cells[`A${index + 1}`] = `Value ${String(index).padStart(3, "0")}`;
    }
    const sheet: SheetModel = {
      ...createFixtureSheet().sheet,
      rowCount: 202,
      columnCount: 1,
      cells,
      autoFilterRange: {
        start: { row: 0, column: 0 },
        end: { row: 201, column: 0 }
      }
    };
    const workbook: WorkbookModel = {
      version: 1,
      activeSheetId: sheet.id,
      sheets: [sheet],
      namedRanges: []
    };
    const formulaEngine = createFormulaEngine(workbook);
    const displayValueSpy = vi.spyOn(formulaEngine, "getDisplayValue");

    render(
      <Grid
        sheet={sheet}
        formulaEngine={formulaEngine}
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

    expect(displayValueSpy).not.toHaveBeenCalledWith(sheet.id, "A202");

    fireEvent.click(screen.getByRole("button", { name: "Open AutoFilter menu for Region" }));
    const menu = screen.getByRole("menu", { name: "AutoFilter menu for Region" });

    expect(displayValueSpy).toHaveBeenCalledWith(sheet.id, "A202");
    const scansAfterOpen = displayValueSpy.mock.calls.filter(([, address]) => address === "A202").length;
    expect(scansAfterOpen).toBe(1);
    expect(within(menu).getAllByRole("menuitemcheckbox")).toHaveLength(200);
    expect(within(menu).getByRole("status")).toHaveTextContent("Showing 200 of 201 values");
    expect(within(menu).queryByRole("menuitemcheckbox", { name: "Value 201" })).not.toBeInTheDocument();

    fireEvent.change(within(menu).getByRole("searchbox", { name: "Search Region filter values" }), {
      target: { value: "Value 201" }
    });

    expect(within(menu).getAllByRole("menuitemcheckbox")).toHaveLength(1);
    expect(within(menu).getByRole("menuitemcheckbox", { name: "Value 201" })).toBeInTheDocument();
    expect(within(menu).getByRole("status")).toHaveTextContent("Showing 1 of 1 matching values (201 total)");

    fireEvent.click(within(menu).getByRole("menuitemcheckbox", { name: "Value 201" }));
    expect(displayValueSpy.mock.calls.filter(([, address]) => address === "A202")).toHaveLength(scansAfterOpen);
  });

  it("keeps blank, formula-empty, and zero AutoFilter choices distinct", () => {
    const sheet: SheetModel = {
      ...createFixtureSheet().sheet,
      rowCount: 4,
      columnCount: 1,
      cells: { A1: "Value", A3: '=""', A4: 0 },
      autoFilterRange: {
        start: { row: 0, column: 0 },
        end: { row: 3, column: 0 }
      }
    };
    const workbook: WorkbookModel = {
      version: 1,
      activeSheetId: sheet.id,
      sheets: [sheet],
      namedRanges: []
    };
    const appliedValues: string[][] = [];

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
        onAutoFilterColumn={(_column, values) => appliedValues.push(values)}
      />
    );

    fireEvent.click(screen.getByRole("button", { name: "Open AutoFilter menu for Value" }));
    const menu = screen.getByRole("menu", { name: "AutoFilter menu for Value" });

    expect(within(menu).getByRole("menuitemcheckbox", { name: "Blank" })).toBeInTheDocument();
    expect(within(menu).getByRole("menuitemcheckbox", { name: "Empty result" })).toBeInTheDocument();
    expect(within(menu).getByRole("menuitemcheckbox", { name: "0" })).toBeInTheDocument();

    fireEvent.click(within(menu).getByRole("menuitemcheckbox", { name: "Blank" }));
    fireEvent.click(within(menu).getByRole("menuitem", { name: "Apply selected values" }));

    expect(appliedValues).toEqual([[BLANK_FILTER_VALUE]]);
  });

  it("orders AutoFilter choices deterministically without the runtime's default locale", () => {
    const sheet: SheetModel = {
      ...createFixtureSheet().sheet,
      rowCount: 3,
      columnCount: 1,
      cells: { A1: "Value", A2: "ä", A3: "z" },
      autoFilterRange: {
        start: { row: 0, column: 0 },
        end: { row: 2, column: 0 }
      }
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

    fireEvent.click(screen.getByRole("button", { name: "Open AutoFilter menu for Value" }));
    expect(
      within(screen.getByRole("menu", { name: "AutoFilter menu for Value" }))
        .getAllByRole("menuitemcheckbox")
        .map((choice) => choice.textContent)
    ).toEqual(["z", "ä"]);
  });
});

const BLANK_FILTER_VALUE = "\u0000js-spreadsheet:blank";

function createFixtureSheet(): { sheet: SheetModel; workbook: WorkbookModel } {
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
  return { sheet, workbook };
}
