# Inline Pivot Drilldown Design

## Goal

Pivot table value cells can expand directly inside the pivot sheet, inserting the matching source rows below the selected summary row. Collapsing the same value cell removes those inserted detail rows.

## Behavior

- Drilldown is available on generated pivot sheets only.
- Each drillable pivot value cell has a compact expand/collapse control inside the cell.
- Expanding inserts a source-header row and the matching source data rows directly below that pivot summary row.
- Collapsing removes only that drilldown's inserted rows.
- Detail rows are rebuilt from the original selected source snapshot stored on the pivot sheet, so the pivot drilldown remains stable after switching sheets or undoing/redoing workbook history.
- Row-only pivots drill into the row group total.
- Row-and-column pivots drill into each column value and each group grand total.

## Architecture

- `src/lib/pivot.ts` builds the normal pivot matrix and a serializable drilldown metadata model in the same pass.
- `SheetModel` stores optional pivot metadata with the base pivot rows, source headers, drilldown entries, and expanded entry state.
- `App.tsx` materializes pivot sheet rows from metadata whenever a drilldown toggles, then replaces the generated pivot sheet contents.
- `Grid.tsx` renders the compact drilldown button for cells reported by the pivot metadata helper.

## Testing

- Unit tests cover metadata creation, source row filtering, and materializing expanded rows.
- Browser tests cover creating a nested pivot, expanding a summary row into source rows, and collapsing it again.
- Full build, unit tests, and Playwright tests run before the branch is pushed.
