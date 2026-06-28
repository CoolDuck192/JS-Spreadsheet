# JavaScript Spreadsheet Clone Design

Date: 2026-06-26

## Goal

Create a standalone browser-based spreadsheet application that feels familiar to Excel users and can be used as a JavaScript spreadsheet. It should provide a real editable grid, formulas, range references, common spreadsheet workflows, local persistence, and import/export paths.

## Scope

The first implementation will build a single-page web app with:

- A spreadsheet grid with column letters, row numbers, selected cells, active cell editing, keyboard navigation, and range selection.
- A formula bar that shows and edits the active cell's raw content.
- Formula calculation through HyperFormula, wrapped behind a local adapter so the app is not tightly coupled to one engine.
- Support for Excel-style formulas provided by the engine, including arithmetic, ranges, references, and supported built-in functions.
- Spreadsheet-style error display for invalid formulas, invalid references, circular dependencies, and calculation errors.
- Copy, paste, fill-down, fill-right, undo, redo, and clear-cell interactions.
- Sheet tabs with add, rename, duplicate, and delete.
- CSV import and export.
- Browser local-storage persistence with reset/new-workbook controls.
- Tests for the formula adapter, cell addressing, CSV handling, persistence, and important UI interactions.

This design does not promise perfect Microsoft Excel parity. Excel contains proprietary behavior and a very large function surface. The app will use HyperFormula for broad Excel-compatible formula support and will keep the engine boundary isolated so missing functions or licensing constraints can be addressed later.

## Recommended Architecture

Use Vite, React, and TypeScript for the app. Keep the spreadsheet logic in plain TypeScript modules that can be tested independently from React.

Core modules:

- `FormulaEngine`: owns the HyperFormula instance, exposes methods to set cell contents, read display values, rename sheets, add sheets, remove sheets, and serialize/restore workbook state.
- `WorkbookStore`: keeps app-level workbook state: sheets, dimensions, active sheet, selection, active cell, undo/redo history, and formatting metadata.
- `Grid`: renders row and column headers, cell values, active selection, edit mode, and keyboard/mouse interactions.
- `FormulaBar`: shows active cell address and raw cell content, commits formula or value edits.
- `Toolbar`: exposes common commands: new, import CSV, export CSV, undo, redo, clear, add sheet, duplicate sheet, delete sheet.
- `SheetTabs`: controls active sheet switching, adding, renaming, duplicating, and deleting.
- `csv`: parses and serializes CSV data with quoted cells, commas, and newlines.
- `addressing`: converts between `A1` references and zero-based row/column coordinates.
- `persistence`: saves and restores workbooks from browser storage.

## Data Model

The app stores user-entered cell content separately from displayed calculated values.

```ts
type CellContent = string | number | boolean | null;

type SheetModel = {
  id: string;
  name: string;
  rowCount: number;
  columnCount: number;
  cells: Record<string, CellContent>;
};

type WorkbookModel = {
  version: 1;
  activeSheetId: string;
  sheets: SheetModel[];
};
```

Cell keys use `A1` notation for storage readability. The formula engine adapter converts between `A1` addresses and engine coordinates.

## Formula Behavior

Cells beginning with `=` are formulas. Other values are stored as raw cell content and displayed directly. Formula results are read from HyperFormula and rendered into the grid. Errors are shown as spreadsheet-style values such as `#DIV/0!`, `#VALUE!`, `#REF!`, `#NAME?`, and circular dependency errors.

The adapter will expose a stable API:

```ts
type SpreadsheetEngine = {
  loadWorkbook(workbook: WorkbookModel): void;
  getCellDisplayValue(sheetId: string, address: string): string;
  setCellContent(sheetId: string, address: string, content: CellContent): WorkbookModel;
  addSheet(name: string): WorkbookModel;
  removeSheet(sheetId: string): WorkbookModel;
  renameSheet(sheetId: string, name: string): WorkbookModel;
};
```

## UI Behavior

The first screen is the spreadsheet itself. There is no landing page.

The layout has:

- A compact top toolbar.
- A formula bar with active address and editable raw value.
- A full-height grid area.
- A bottom sheet tab strip.

Keyboard expectations:

- Arrow keys move the active cell.
- Enter commits edits and moves down.
- Tab commits edits and moves right.
- Escape cancels edit mode.
- Delete or Backspace clears the current selection.
- Cmd/Ctrl+C copies selected cells.
- Cmd/Ctrl+V pastes tabular data.
- Cmd/Ctrl+D fills the selected range down from the top row.
- Cmd/Ctrl+R fills the selected range right from the left column.
- Cmd/Ctrl+Z and Cmd/Ctrl+Shift+Z perform undo and redo.

Mouse expectations:

- Click selects a cell.
- Double-click or typing starts edit mode.
- Drag selects a range.
- Sheet tab controls support rename, duplicate, delete, and add.

## Import And Export

CSV import replaces or creates a sheet from a selected `.csv` file. CSV export downloads the active sheet. The CSV parser and serializer must support quoted fields, embedded commas, escaped quotes, and line breaks in quoted fields.

XLSX import/export is intentionally left as a follow-up because it adds a separate file-format surface. The module boundaries leave room to add it with a library such as SheetJS later.

## Persistence

The app saves the current workbook to local storage after edits. On load, it restores the last workbook. The toolbar provides a new/reset command to start over with a blank workbook.

## Error Handling

Formula and data errors must not crash the app. Displayable spreadsheet errors should remain in cells. Invalid CSV files should produce a readable toast/status message. Persistence restore failures should fall back to a blank workbook and keep the app usable.

## Testing

Automated tests should cover:

- `addressing`: `A1`, `Z1`, `AA1`, multi-letter columns, and invalid references.
- `csv`: quotes, commas, escaped quotes, blank cells, and multiline fields.
- `FormulaEngine`: raw values, formulas, ranges, references, recalculation, and formula errors.
- `WorkbookStore`: editing, undo, redo, sheet creation, duplication, rename, deletion, and persistence serialization.
- UI smoke tests: cell edit, formula recalculation, range copy/paste, CSV export availability, and sheet tab behavior.

## Acceptance Criteria

The build is acceptable when:

- The app runs locally through the provided npm scripts.
- A user can edit cells, enter formulas, see recalculated results, and inspect raw formulas through the formula bar.
- Common formulas such as `=SUM(A1:A3)`, `=AVERAGE(A1:A3)`, `=IF(A1>5,"yes","no")`, and cross-cell references work.
- The grid supports keyboard navigation, selection, copy/paste, fill-down, fill-right, undo, redo, and clearing cells.
- Users can add, rename, duplicate, delete, and switch sheets.
- Users can import and export CSV files.
- A workbook persists across browser reloads.
- Unit tests and build verification pass.
