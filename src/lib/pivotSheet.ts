import type { CellContent, CellFormat, PivotSheetMetadata, WorkbookModel } from "../types";
import { formatCellAddress, parseCellAddress } from "./addressing";
import {
  createPivotDrilldownIndex,
  getPivotMaterializedBaseRow,
  getPivotMaterializedRowKind,
  materializePivotRows
} from "./pivot";

export function replaceGeneratedPivotSheetRows(
  workbook: WorkbookModel,
  sheetId: string,
  pivot: PivotSheetMetadata
): WorkbookModel {
  const index = createPivotDrilldownIndex(pivot);
  const matrix = materializePivotRows(pivot, index);
  const matrixColumnCount = matrixWidth(matrix);
  const columnCount = Math.max(26, matrixColumnCount);
  const rowCount = Math.max(100, matrix.length);

  return {
    ...workbook,
    sheets: workbook.sheets.map((sheet) => {
      if (sheet.id !== sheetId) {
        return sheet;
      }

      const previousPivotRows = sheet.pivot ? materializePivotRows(sheet.pivot) : [];
      const generatedRowCount = Math.max(previousPivotRows.length, matrix.length);
      const generatedCellColumnCount = Math.max(matrixWidth(previousPivotRows), matrixColumnCount);
      const generatedFormatColumnCount = Math.max(26, generatedCellColumnCount);
      const cells = withoutGeneratedRegion(sheet.cells, generatedRowCount, generatedCellColumnCount);
      const formats = withoutGeneratedRegion(sheet.formats ?? {}, generatedRowCount, generatedFormatColumnCount);

      matrix.forEach((rowValues, row) => {
        rowValues.forEach((content, column) => {
          if (content !== "") {
            cells[formatCellAddress({ row, column })] = content;
          }
        });
      });

      applyRowFormat(formats, 0, columnCount, {
        bold: true,
        textColor: "#17634a",
        backgroundColor: "#eaf7f2"
      });

      const grandTotalRow = getPivotMaterializedBaseRow(index, pivot.baseRows.length - 1);
      matrix.forEach((_, row) => {
        const rowKind = getPivotMaterializedRowKind(pivot, row, index);
        if (rowKind?.kind === "detail-header") {
          applyRowFormat(formats, row, columnCount, {
            bold: true,
            textColor: "#475569",
            backgroundColor: "#f8fafc"
          });
          return;
        }

        if (rowKind?.kind === "detail-row") {
          applyRowFormat(formats, row, columnCount, {
            textColor: "#334155",
            backgroundColor: "#fffdf7"
          });
          return;
        }

        if (row > 0 && row === grandTotalRow) {
          applyRowFormat(formats, row, columnCount, {
            bold: true,
            backgroundColor: "#f1f5f8"
          });
        }
      });

      return {
        ...sheet,
        rowCount,
        columnCount,
        cells,
        formats,
        pivot
      };
    })
  };
}

function withoutGeneratedRegion<T>(record: Record<string, T>, rowCount: number, columnCount: number): Record<string, T> {
  const nextRecord: Record<string, T> = {};
  for (const [address, value] of Object.entries(record)) {
    const coord = parseCellAddress(address);
    if (coord.row >= rowCount || coord.column >= columnCount) {
      nextRecord[address] = value;
    }
  }
  return nextRecord;
}

function applyRowFormat(formats: Record<string, CellFormat>, row: number, columnCount: number, format: CellFormat) {
  for (let column = 0; column < columnCount; column += 1) {
    formats[formatCellAddress({ row, column })] = { ...format };
  }
}

function matrixWidth(rows: CellContent[][]): number {
  return rows.reduce((width, row) => Math.max(width, row.length), 0);
}
