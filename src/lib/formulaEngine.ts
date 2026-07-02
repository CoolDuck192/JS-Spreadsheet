import { DetailedCellError, HyperFormula } from "hyperformula";
import type { CellContent, CellRange, NamedRange, SheetModel, WorkbookModel } from "../types";
import { parseCellAddress } from "./addressing";
import { extractFormulaReferences } from "./formulaReferences";
import { getCellContent } from "./workbook";

export type FormulaEngine = {
  getDisplayValue: (sheetId: string, address: string) => string;
  getRawContent: (sheetId: string, address: string) => CellContent;
  rebuild: (workbook: WorkbookModel) => void;
};

type EngineState = {
  workbook: WorkbookModel;
  hyperFormula: HyperFormula;
  sheetIds: Map<string, number>;
};

export function createFormulaEngine(workbook: WorkbookModel): FormulaEngine {
  let state = buildEngineState(workbook);

  return {
    getDisplayValue(sheetId, address) {
      const sheet = state.sheetIds.get(sheetId);
      if (sheet === undefined) {
        return "";
      }

      const coord = parseCellAddress(address);
      const value = state.hyperFormula.getCellValue({ sheet, col: coord.column, row: coord.row });

      return formatCellValue(value);
    },

    getRawContent(sheetId, address) {
      return getCellContent(state.workbook, sheetId, address);
    },

    rebuild(nextWorkbook) {
      state.hyperFormula.destroy();
      state = buildEngineState(nextWorkbook);
    }
  };
}

function buildEngineState(workbook: WorkbookModel): EngineState {
  const sheets = Object.fromEntries(
    workbook.sheets.map((sheet) => [
      sheet.name,
      createFormulaSheetMatrix(
        sheet,
        (workbook.namedRanges ?? []).filter((namedRange) => namedRange.sheetId === sheet.id)
      )
    ])
  );
  const namedExpressions = (workbook.namedRanges ?? []).flatMap((namedRange) => {
    const sheet = workbook.sheets.find((candidate) => candidate.id === namedRange.sheetId);
    return sheet ? [namedRangeToExpression(namedRange, sheet)] : [];
  });
  const hyperFormula = HyperFormula.buildFromSheets(sheets, {
    licenseKey: "gpl-v3",
    useColumnIndex: false,
    useStats: false
  }, namedExpressions);

  return {
    workbook,
    hyperFormula,
    sheetIds: new Map(
      workbook.sheets.map((sheet) => {
        const hyperFormulaSheetId = hyperFormula.getSheetId(sheet.name);
        if (hyperFormulaSheetId === undefined) {
          throw new Error(`Formula engine did not create sheet: ${sheet.name}`);
        }
        return [sheet.id, hyperFormulaSheetId];
      })
    )
  };
}

export function createFormulaSheetMatrix(sheet: SheetModel, namedRanges: readonly NamedRange[] = []): CellContent[][] {
  const bounds = getFormulaMatrixBounds(sheet, namedRanges);

  return Array.from({ length: bounds.rowCount }, (_, row) => {
    const values: CellContent[] = [];
    for (let column = 0; column < bounds.columnCount; column += 1) {
      const address = `${columnName(column)}${row + 1}`;
      if (address in sheet.cells) {
        values[column] = sheet.cells[address] ?? null;
      }
    }
    return trimTrailingEmptyCells(values);
  });
}

function getFormulaMatrixBounds(sheet: SheetModel, namedRanges: readonly NamedRange[]) {
  let maxRow = 0;
  let maxColumn = 0;

  function includeRange(range: CellRange) {
    maxRow = Math.max(maxRow, range.start.row, range.end.row);
    maxColumn = Math.max(maxColumn, range.start.column, range.end.column);
  }

  for (const [address, content] of Object.entries(sheet.cells)) {
    const coord = parseCellAddress(address);
    maxRow = Math.max(maxRow, coord.row);
    maxColumn = Math.max(maxColumn, coord.column);

    if (typeof content === "string" && content.startsWith("=")) {
      for (const reference of extractFormulaReferences(content)) {
        includeRange(reference.range);
      }
    }
  }

  for (const namedRange of namedRanges) {
    includeRange(namedRange.range);
  }

  return {
    rowCount: Math.max(1, maxRow + 1),
    columnCount: Math.max(1, maxColumn + 1)
  };
}

function trimTrailingEmptyCells(values: CellContent[]): CellContent[] {
  let end = values.length;
  while (end > 0 && (values[end - 1] === null || values[end - 1] === undefined)) {
    end -= 1;
  }
  return values.slice(0, end);
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
