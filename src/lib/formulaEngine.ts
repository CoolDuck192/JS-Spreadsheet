import { DetailedCellError, HyperFormula } from "hyperformula";
import type { CellContent, CellRange, NamedRange, SheetModel, WorkbookModel } from "../types";
import { parseCellAddress } from "./addressing";
import { getCellContent } from "./workbook";

export type FormulaEngine = {
  getDisplayValue: (sheetId: string, address: string) => string;
  getRawContent: (sheetId: string, address: string) => CellContent;
  /** Incrementally sync the engine to a new workbook snapshot. Cheap when only cell contents changed. */
  update: (workbook: WorkbookModel) => void;
  rebuild: (workbook: WorkbookModel) => void;
  destroy: () => void;
};

type EngineState = {
  workbook: WorkbookModel;
  hyperFormula: HyperFormula;
  sheetIds: Map<string, number>;
};

// Excel's grid limits. HyperFormula's defaults (40,000 rows) crash on larger sheets.
const EXCEL_MAX_ROWS = 1_048_576;
const EXCEL_MAX_COLUMNS = 16_384;

const ENGINE_CONFIG = {
  licenseKey: "gpl-v3",
  useColumnIndex: false,
  useStats: false,
  maxRows: EXCEL_MAX_ROWS,
  maxColumns: EXCEL_MAX_COLUMNS,
  // Excel-compatible date semantics: ISO and US formats parse as date serials.
  // leapYear1900 + a 1899-12-31 nullDate reproduce Excel's serial numbering
  // exactly, including Jan-Feb 1900 and the phantom 1900-02-29 (verified against
  // Excel ground truth: 2026-01-15=46037, 1900-01-15=15, 1900-03-01=61).
  dateFormats: ["YYYY-MM-DD", "MM/DD/YYYY", "MM/DD/YY"],
  leapYear1900: true,
  nullDate: { year: 1899, month: 12, day: 31 }
};

export function createFormulaEngine(workbook: WorkbookModel): FormulaEngine {
  let state = buildEngineState(workbook);
  let destroyed = false;

  // destroy() releases the HyperFormula instance; any later call revives it from
  // the retained workbook snapshot. This keeps unmount cleanup leak-free while
  // surviving React StrictMode's simulated unmount/remount, where the rendered
  // tree still holds this engine reference.
  function ensureAlive() {
    if (destroyed) {
      state = buildEngineState(state.workbook);
      destroyed = false;
    }
  }

  return {
    getDisplayValue(sheetId, address) {
      ensureAlive();
      const sheet = state.sheetIds.get(sheetId);
      if (sheet === undefined) {
        return "";
      }

      const coord = parseCellAddress(address);
      const value = state.hyperFormula.getCellValue({ sheet, col: coord.column, row: coord.row });

      return formatCellValue(value);
    },

    getRawContent(sheetId, address) {
      if (!state.workbook.sheets.some((sheet) => sheet.id === sheetId)) {
        return null;
      }
      return getCellContent(state.workbook, sheetId, address);
    },

    update(nextWorkbook) {
      ensureAlive();
      state = updateEngineState(state, nextWorkbook);
    },

    rebuild(nextWorkbook) {
      if (!destroyed) {
        state.hyperFormula.destroy();
      }
      destroyed = false;
      state = buildEngineState(nextWorkbook);
    },

    destroy() {
      if (!destroyed) {
        state.hyperFormula.destroy();
        destroyed = true;
      }
    }
  };
}

function updateEngineState(state: EngineState, nextWorkbook: WorkbookModel): EngineState {
  const previous = state.workbook;
  if (previous === nextWorkbook) {
    return state;
  }

  if (requiresFullRebuild(previous, nextWorkbook)) {
    state.hyperFormula.destroy();
    return buildEngineState(nextWorkbook);
  }

  const changedSheets = nextWorkbook.sheets.filter((sheet, index) => {
    const previousSheet = previous.sheets[index];
    return previousSheet !== sheet && previousSheet.cells !== sheet.cells;
  });

  if (changedSheets.length > 0) {
    state.hyperFormula.batch(() => {
      for (const sheet of changedSheets) {
        const engineSheetId = state.sheetIds.get(sheet.id);
        if (engineSheetId === undefined) {
          continue;
        }
        const previousSheet = previous.sheets.find((candidate) => candidate.id === sheet.id);
        applyCellDiff(state.hyperFormula, engineSheetId, previousSheet?.cells ?? {}, sheet.cells);
      }
    });
  }

  return { ...state, workbook: nextWorkbook };
}

function requiresFullRebuild(previous: WorkbookModel, next: WorkbookModel): boolean {
  if (previous.namedRanges !== next.namedRanges) {
    return true;
  }
  if (previous.sheets.length !== next.sheets.length) {
    return true;
  }
  return next.sheets.some((sheet, index) => {
    const previousSheet = previous.sheets[index];
    return previousSheet.id !== sheet.id || previousSheet.name !== sheet.name;
  });
}

function applyCellDiff(
  hyperFormula: HyperFormula,
  engineSheetId: number,
  previousCells: SheetModel["cells"],
  nextCells: SheetModel["cells"]
): void {
  for (const address in previousCells) {
    if (!(address in nextCells)) {
      const coord = parseCellAddress(address);
      hyperFormula.setCellContents({ sheet: engineSheetId, col: coord.column, row: coord.row }, null);
    }
  }
  for (const address in nextCells) {
    if (previousCells[address] !== nextCells[address]) {
      const coord = parseCellAddress(address);
      hyperFormula.setCellContents({ sheet: engineSheetId, col: coord.column, row: coord.row }, nextCells[address]);
    }
  }
}

function buildEngineState(workbook: WorkbookModel): EngineState {
  // Bulk-load through buildFromSheets with RAGGED arrays sized to actual content.
  // Per-cell setCellContents is ~1000x slower for initial load, and dense
  // rowCount×columnCount arrays explode with Excel-sized grid limits.
  const sheets = Object.fromEntries(workbook.sheets.map((sheet) => [sheet.name, sheetToRaggedMatrix(sheet)]));
  const hyperFormula = HyperFormula.buildFromSheets(sheets, ENGINE_CONFIG);

  const sheetIds = new Map<string, number>(
    workbook.sheets.map((sheet) => {
      const engineSheetId = hyperFormula.getSheetId(sheet.name);
      if (engineSheetId === undefined) {
        throw new Error(`Formula engine did not create sheet: ${sheet.name}`);
      }
      return [sheet.id, engineSheetId];
    })
  );

  for (const namedRange of workbook.namedRanges ?? []) {
    const sheet = workbook.sheets.find((candidate) => candidate.id === namedRange.sheetId);
    if (sheet) {
      const expression = namedRangeToExpression(namedRange, sheet);
      try {
        hyperFormula.addNamedExpression(expression.name, expression.expression);
      } catch {
        // Names that collide with cell notation (e.g. "RC", "R1C2") are rejected by
        // HyperFormula; skip them instead of failing the whole engine build.
      }
    }
  }

  return { workbook, hyperFormula, sheetIds };
}

function sheetToRaggedMatrix(sheet: SheetModel): CellContent[][] {
  const rows: CellContent[][] = [];
  for (const address in sheet.cells) {
    const coord = parseCellAddress(address);
    let row = rows[coord.row];
    if (!row) {
      row = [];
      rows[coord.row] = row;
    }
    row[coord.column] = sheet.cells[address];
  }
  // Sparse array holes confuse HyperFormula's row iteration; normalize to empty rows.
  for (let index = 0; index < rows.length; index += 1) {
    rows[index] ??= [];
  }
  return rows;
}

function columnName(index: number): string {
  let remaining = index + 1;
  let name = "";

  while (remaining > 0) {
    const remainder = (remaining - 1) % 26;
    name = String.fromCharCode(65 + remainder) + name;
    remaining = Math.floor((remaining - 1) / 26);
  }

  return name;
}

function namedRangeToExpression(namedRange: NamedRange, sheet: SheetModel) {
  return {
    name: namedRange.name,
    expression: `=${quoteSheetName(sheet.name)}!${formatAbsoluteRange(namedRange.range)}`
  };
}

function formatAbsoluteRange(range: CellRange): string {
  const start = formatAbsoluteCell(range.start);
  const end = formatAbsoluteCell(range.end);
  return start === end ? start : `${start}:${end}`;
}

function formatAbsoluteCell(coord: { row: number; column: number }): string {
  return `$${columnName(coord.column)}$${coord.row + 1}`;
}

function quoteSheetName(name: string): string {
  return `'${name.replace(/'/g, "''")}'`;
}

function formatCellValue(value: unknown): string {
  if (value === null || value === undefined) {
    return "";
  }

  if (value instanceof DetailedCellError) {
    return value.value;
  }

  if (isDetailedCellErrorLike(value)) {
    return value.value;
  }

  if (typeof value === "boolean") {
    return value ? "TRUE" : "FALSE";
  }

  return String(value);
}

function isDetailedCellErrorLike(value: unknown): value is { value: string; type: string } {
  return (
    typeof value === "object" &&
    value !== null &&
    "value" in value &&
    "type" in value &&
    typeof (value as { value: unknown }).value === "string"
  );
}
