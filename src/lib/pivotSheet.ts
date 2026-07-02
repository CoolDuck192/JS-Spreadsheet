import type { CellContent, CellFormat, CellRange, PivotSheetMetadata, SheetChart, WorkbookModel } from "../types";
import { formatCellAddress, normalizeRange, parseCellAddress } from "./addressing";
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

  return {
    ...workbook,
    sheets: workbook.sheets.map((sheet) => {
      if (sheet.id !== sheetId) {
        return sheet;
      }

      const columnCount = Math.max(sheet.columnCount, 26, matrixColumnCount);
      const rowCount = Math.max(sheet.rowCount, 100, matrix.length);
      const previousPivotRows = sheet.pivot ? materializePivotRows(sheet.pivot) : [];
      const generatedRowCount = Math.max(previousPivotRows.length, matrix.length);
      const generatedCellColumnCount = Math.max(matrixWidth(previousPivotRows), matrixColumnCount);
      const cells = withoutGeneratedRegion(sheet.cells, generatedRowCount, generatedCellColumnCount);
      const formats = withoutGeneratedRegion(sheet.formats ?? {}, generatedRowCount, generatedCellColumnCount);
      const comments = withoutGeneratedRegion(sheet.comments ?? {}, generatedRowCount, generatedCellColumnCount);
      const hyperlinks = withoutGeneratedRegion(sheet.hyperlinks ?? {}, generatedRowCount, generatedCellColumnCount);
      const validations = withoutGeneratedRegion(sheet.validations ?? {}, generatedRowCount, generatedCellColumnCount);
      const conditionalFormats = (sheet.conditionalFormats ?? []).filter(
        (rule) => !rangeIntersectsGeneratedRegion(rule.range, generatedRowCount, generatedCellColumnCount)
      );
      const charts = (sheet.charts ?? []).filter(
        (chart) =>
          !rangeIntersectsGeneratedRegion(chart.range, generatedRowCount, generatedCellColumnCount) &&
          !coordIsInGeneratedRegion(chart.anchor, generatedRowCount, generatedCellColumnCount)
      );
      const merges = (sheet.merges ?? []).filter(
        (merge) => !rangeIntersectsGeneratedRegion(merge.range, generatedRowCount, generatedCellColumnCount)
      );

      matrix.forEach((rowValues, row) => {
        rowValues.forEach((content, column) => {
          if (content !== "") {
            cells[formatCellAddress({ row, column })] = content;
          }
        });
      });

      applyRowFormat(formats, 0, matrixColumnCount, {
        bold: true,
        textColor: "#17634a",
        backgroundColor: "#eaf7f2"
      });

      const grandTotalRow = getPivotMaterializedBaseRow(index, pivot.baseRows.length - 1);
      matrix.forEach((_, row) => {
        const rowKind = getPivotMaterializedRowKind(pivot, row, index);
        if (rowKind?.kind === "detail-header") {
          applyRowFormat(formats, row, matrixColumnCount, {
            bold: true,
            textColor: "#475569",
            backgroundColor: "#f8fafc"
          });
          return;
        }

        if (rowKind?.kind === "detail-row") {
          applyRowFormat(formats, row, matrixColumnCount, {
            textColor: "#334155",
            backgroundColor: "#fffdf7"
          });
          return;
        }

        if (row > 0 && row === grandTotalRow) {
          applyRowFormat(formats, row, matrixColumnCount, {
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
        comments,
        hyperlinks,
        validations,
        conditionalFormats,
        charts,
        merges,
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

function rangeIntersectsGeneratedRegion(range: CellRange, rowCount: number, columnCount: number): boolean {
  if (rowCount <= 0 || columnCount <= 0) {
    return false;
  }

  const normalized = normalizeRange(range);
  return normalized.start.row < rowCount && normalized.end.row >= 0 && normalized.start.column < columnCount && normalized.end.column >= 0;
}

function coordIsInGeneratedRegion(coord: SheetChart["anchor"], rowCount: number, columnCount: number): boolean {
  return coord.row >= 0 && coord.row < rowCount && coord.column >= 0 && coord.column < columnCount;
}

function applyRowFormat(formats: Record<string, CellFormat>, row: number, columnCount: number, format: CellFormat) {
  for (let column = 0; column < columnCount; column += 1) {
    formats[formatCellAddress({ row, column })] = { ...format };
  }
}

function matrixWidth(rows: CellContent[][]): number {
  return rows.reduce((width, row) => Math.max(width, row.length), 0);
}
