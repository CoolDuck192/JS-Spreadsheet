import type ExcelJS from "exceljs";
import type {
  CellBorderSide,
  CellBorders,
  CellContent,
  CellFormat,
  CellRange,
  ConditionalFormatRule,
  DataValidationRule,
  NamedRange,
  SheetModel,
  SheetProtection,
  WorkbookModel
} from "../types";
import { columnIndexToName, formatCellAddress, parseCellAddress, parseRangeAddress } from "./addressing";
import { DEFAULT_ROW_HEIGHT } from "./sheetDimensions";

const DEFAULT_ROWS = 100;
const DEFAULT_COLUMNS = 26;
const COLUMN_WIDTH_SCALE = 8;
const XLSX_AUTHOR = "JavaScript Spreadsheet";
const XLSX_NUMBER_FORMATS: Record<NonNullable<CellFormat["numberFormat"]>, string> = {
  general: "General",
  number: "#,##0.########",
  currency: "$#,##0.00",
  percent: "0.########%",
  date: "mmm d, yyyy"
};
let excelJsPromise: Promise<typeof ExcelJS> | null = null;

export async function exportWorkbookToXlsx(workbook: WorkbookModel): Promise<ArrayBuffer> {
  const ExcelJS = await loadExcelJs();
  const excelWorkbook = new ExcelJS.Workbook();
  excelWorkbook.creator = XLSX_AUTHOR;

  const worksheetNamesBySheetId = new Map<string, string>();
  for (const sheet of workbook.sheets) {
    const worksheetName = safeWorksheetName(sheet.name, excelWorkbook);
    const worksheet = excelWorkbook.addWorksheet(worksheetName);
    worksheetNamesBySheetId.set(sheet.id, worksheetName);
    await sheetToWorksheet(sheet, worksheet);
  }
  addNamedRangesToExcelWorkbook(workbook.namedRanges ?? [], excelWorkbook, worksheetNamesBySheetId);
  applyActiveSheetToExcelWorkbook(workbook, excelWorkbook);

  return toArrayBuffer(await excelWorkbook.xlsx.writeBuffer());
}

export async function importWorkbookFromXlsx(data: ArrayBuffer | Uint8Array): Promise<WorkbookModel> {
  const ExcelJS = await loadExcelJs();
  const excelWorkbook = new ExcelJS.Workbook();
  await excelWorkbook.xlsx.load(toArrayBuffer(data) as ExcelJS.Buffer);

  const worksheets = excelWorkbook.worksheets.length > 0 ? excelWorkbook.worksheets : [excelWorkbook.addWorksheet("Sheet1")];
  const sheets = worksheets.map((worksheet, index) => worksheetToSheet(worksheet, index));
  const sheetIdsByName = new Map(sheets.map((sheet) => [sheet.name, sheet.id]));
  const namedRanges = excelDefinedNamesToNamedRanges(excelWorkbook.definedNames.model, sheetIdsByName);
  const activeSheetIndex = excelWorkbookActiveSheetIndex(excelWorkbook, sheets.length);
  const activeSheet = sheets[activeSheetIndex]?.isHidden ? sheets.find((sheet) => sheet.isHidden !== true) ?? sheets[0] : sheets[activeSheetIndex];

  return {
    version: 1,
    activeSheetId: activeSheet?.id ?? sheets[0].id,
    sheets,
    namedRanges
  };
}

function applyActiveSheetToExcelWorkbook(workbook: WorkbookModel, excelWorkbook: ExcelJS.Workbook): void {
  if (workbook.sheets.length === 0) {
    return;
  }

  const activeVisibleSheetIndex = workbook.sheets.findIndex((sheet) => sheet.id === workbook.activeSheetId && sheet.isHidden !== true);
  const firstVisibleSheetIndex = workbook.sheets.findIndex((sheet) => sheet.isHidden !== true);
  const activeSheetIndex = activeVisibleSheetIndex >= 0 ? activeVisibleSheetIndex : Math.max(0, firstVisibleSheetIndex);
  excelWorkbook.views = [
    {
      x: 0,
      y: 0,
      width: 12000,
      height: 24000,
      firstSheet: 0,
      activeTab: activeSheetIndex,
      visibility: "visible"
    }
  ];
}

function excelWorkbookActiveSheetIndex(workbook: ExcelJS.Workbook, sheetCount: number): number {
  const views = Array.isArray(workbook.views) ? workbook.views : [];
  const activeTab = views
    .map((view) => (view as Partial<ExcelJS.WorkbookView>).activeTab)
    .find((tab): tab is number => Number.isInteger(tab));

  return activeTab !== undefined && activeTab >= 0 && activeTab < sheetCount ? activeTab : 0;
}

async function sheetToWorksheet(sheet: SheetModel, worksheet: ExcelJS.Worksheet): Promise<void> {
  worksheet.state = sheet.isHidden ? "hidden" : "visible";
  if (sheet.tabColor) {
    worksheet.properties.tabColor = { argb: hexToArgb(sheet.tabColor) };
  }

  for (const [address, content] of Object.entries(sheet.cells)) {
    const cell = worksheet.getCell(address);
    cell.value = cellContentToExcelValue(content, sheet.hyperlinks?.[address]);
  }

  for (const [address, hyperlink] of Object.entries(sheet.hyperlinks ?? {})) {
    const cell = worksheet.getCell(address);
    const text = cellText(cell.value) || hyperlink;
    cell.value = { text, hyperlink };
  }

  for (const [address, comment] of Object.entries(sheet.comments ?? {})) {
    const cell = worksheet.getCell(address);
    cell.note = {
      texts: [{ text: comment }],
      margins: { insetmode: "auto" },
      protection: { locked: "True", lockText: "False" },
      editAs: "absolute"
    };
  }

  for (const [address, format] of Object.entries(sheet.formats ?? {})) {
    applyCellFormatToExcelCell(worksheet.getCell(address), format);
  }

  for (const [address, rule] of Object.entries(sheet.validations ?? {})) {
    const dataValidation = dataValidationRuleToExcel(rule);
    if (dataValidation) {
      worksheet.getCell(address).dataValidation = dataValidation;
    }
  }

  for (const [ruleIndex, conditionalFormat] of (sheet.conditionalFormats ?? []).entries()) {
    const excelRule = conditionalFormatRuleToExcel(conditionalFormat, ruleIndex + 1);
    if (excelRule) {
      worksheet.addConditionalFormatting({
        ref: formatRangeReference(conditionalFormat.range),
        rules: [excelRule]
      });
    }
  }
  applyAutoFilterRangeToWorksheet(sheet, worksheet);

  for (const merge of sheet.merges ?? []) {
    const range = normalizeRange(merge.range);
    worksheet.mergeCells(range.start.row + 1, range.start.column + 1, range.end.row + 1, range.end.column + 1);
  }

  for (const [column, width] of Object.entries(sheet.columnWidths ?? {})) {
    worksheet.getColumn(Number(column) + 1).width = width / COLUMN_WIDTH_SCALE;
  }

  for (const [column, isHidden] of Object.entries(sheet.hiddenColumns ?? {})) {
    if (isHidden) {
      worksheet.getColumn(Number(column) + 1).hidden = true;
    }
  }

  for (const [row, height] of Object.entries(sheet.rowHeights ?? {})) {
    worksheet.getRow(Number(row) + 1).height = height;
  }

  for (const [row, isHidden] of Object.entries(sheet.hiddenRows ?? {})) {
    if (isHidden) {
      const worksheetRow = worksheet.getRow(Number(row) + 1);
      worksheetRow.hidden = true;
      if (typeof worksheetRow.height !== "number") {
        worksheetRow.height = DEFAULT_ROW_HEIGHT;
      }
    }
  }

  applyFreezePanesToWorksheet(sheet, worksheet);
  await applySheetProtectionToWorksheet(sheet.protection, worksheet);
}

function worksheetToSheet(worksheet: ExcelJS.Worksheet, index: number): SheetModel {
  const cells: SheetModel["cells"] = {};
  const formats: SheetModel["formats"] = {};
  const comments: SheetModel["comments"] = {};
  const hyperlinks: SheetModel["hyperlinks"] = {};
  const validations: SheetModel["validations"] = {};
  const unlockedCells: Record<string, boolean> = {};
  let maxRow = 0;
  let maxColumn = 0;

  worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    row.eachCell({ includeEmpty: false }, (cell, columnNumber) => {
      const address = formatCellAddress({ row: rowNumber - 1, column: columnNumber - 1 });
      maxRow = Math.max(maxRow, rowNumber - 1);
      maxColumn = Math.max(maxColumn, columnNumber - 1);

      const content = excelCellToCellContent(cell);
      if (content !== null) {
        cells[address] = content;
      }

      const format = excelCellToCellFormat(cell);
      if (Object.keys(format).length > 0) {
        formats[address] = format;
      }

      const validation = excelDataValidationToRule(cell.dataValidation);
      if (validation) {
        validations[address] = validation;
      }

      if (cell.protection?.locked === false) {
        unlockedCells[address] = true;
      }

      const hyperlink = excelValueToHyperlink(cell.value);
      if (hyperlink) {
        hyperlinks[address] = hyperlink;
      }

      const comment = noteToText(cell.note);
      if (comment) {
        comments[address] = comment;
      }
    });
  });

  const columnWidths: SheetModel["columnWidths"] = {};
  const hiddenColumns: NonNullable<SheetModel["hiddenColumns"]> = {};
  const worksheetColumns = Array.isArray(worksheet.columns) ? worksheet.columns : [];
  worksheetColumns.forEach((column, columnIndex) => {
    if (typeof column.width === "number") {
      columnWidths[String(columnIndex)] = Math.round(column.width * COLUMN_WIDTH_SCALE);
      maxColumn = Math.max(maxColumn, columnIndex);
    }
    if (column.hidden) {
      hiddenColumns[String(columnIndex)] = true;
      maxColumn = Math.max(maxColumn, columnIndex);
    }
  });

  const rowHeights: SheetModel["rowHeights"] = {};
  const hiddenRows: NonNullable<SheetModel["hiddenRows"]> = {};
  worksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
    if (typeof row.height === "number") {
      rowHeights[String(rowNumber - 1)] = row.height;
      maxRow = Math.max(maxRow, rowNumber - 1);
    }
    if (row.hidden) {
      hiddenRows[String(rowNumber - 1)] = true;
      maxRow = Math.max(maxRow, rowNumber - 1);
    }
  });

  const merges = ((worksheet.model.merges ?? []) as string[]).map((mergeAddress, mergeIndex) => ({
    id: `merge-${mergeIndex + 1}`,
    range: parseRangeAddress(mergeAddress)
  }));
  for (const merge of merges) {
    maxRow = Math.max(maxRow, merge.range.end.row);
    maxColumn = Math.max(maxColumn, merge.range.end.column);
  }

  const conditionalFormats = worksheetConditionalFormatsToRules(worksheet);
  for (const rule of conditionalFormats) {
    maxRow = Math.max(maxRow, rule.range.end.row);
    maxColumn = Math.max(maxColumn, rule.range.end.column);
  }
  const autoFilterRange = worksheetAutoFilterRange(worksheet);
  if (autoFilterRange) {
    maxRow = Math.max(maxRow, autoFilterRange.end.row);
    maxColumn = Math.max(maxColumn, autoFilterRange.end.column);
  }
  const freezePanes = worksheetFreezePanes(worksheet);

  return {
    id: `sheet-${index + 1}`,
    name: worksheet.name || `Sheet${index + 1}`,
    rowCount: Math.max(DEFAULT_ROWS, maxRow + 1),
    columnCount: Math.max(DEFAULT_COLUMNS, maxColumn + 1),
    isHidden: worksheet.state === "hidden" || worksheet.state === "veryHidden",
    tabColor: excelColorToHex(worksheet.properties?.tabColor),
    cells,
    formats,
    columnWidths,
    rowHeights,
    hiddenColumns,
    hiddenRows,
    freezeTopRow: freezePanes.freezeTopRow,
    freezeFirstColumn: freezePanes.freezeFirstColumn,
    comments,
    hyperlinks,
    validations,
    conditionalFormats,
    autoFilterRange: autoFilterRange ?? undefined,
    filters: [],
    charts: [],
    merges,
    protection: worksheetProtectionToSheetProtection(worksheet, unlockedCells)
  };
}

function applyAutoFilterRangeToWorksheet(sheet: SheetModel, worksheet: ExcelJS.Worksheet): void {
  const range = sheet.autoFilterRange ?? sheet.filters?.[0]?.range;
  if (!range) {
    return;
  }

  worksheet.autoFilter = formatRangeReference(range);
}

function worksheetAutoFilterRange(worksheet: ExcelJS.Worksheet): CellRange | null {
  const reference = excelAutoFilterToRangeReference(worksheet.autoFilter);
  if (!reference) {
    return null;
  }

  try {
    return parseRangeAddress(reference.replace(/\$/g, ""));
  } catch {
    return null;
  }
}

function excelAutoFilterToRangeReference(autoFilter: ExcelJS.AutoFilter | undefined): string | null {
  if (typeof autoFilter === "string") {
    return autoFilter;
  }

  if (!autoFilter) {
    return null;
  }

  const from = excelAutoFilterAddress(autoFilter.from);
  const to = excelAutoFilterAddress(autoFilter.to);
  return from && to ? `${from}:${to}` : null;
}

type ExcelAutoFilterAddress = string | { row: number; column: number };

function excelAutoFilterAddress(address: ExcelAutoFilterAddress): string | null {
  if (typeof address === "string") {
    return address;
  }

  if (!isRecord(address) || !Number.isInteger(address.row) || !Number.isInteger(address.column)) {
    return null;
  }

  return formatCellAddress({ row: address.row - 1, column: address.column - 1 });
}

function applyFreezePanesToWorksheet(sheet: SheetModel, worksheet: ExcelJS.Worksheet): void {
  const xSplit = sheet.freezeFirstColumn ? 1 : 0;
  const ySplit = sheet.freezeTopRow ? 1 : 0;
  if (xSplit === 0 && ySplit === 0) {
    return;
  }

  worksheet.views = [
    {
      state: "frozen",
      xSplit: xSplit || undefined,
      ySplit: ySplit || undefined
    }
  ];
}

type FrozenWorksheetView = Partial<ExcelJS.WorksheetViewCommon & ExcelJS.WorksheetViewFrozen>;

function worksheetFreezePanes(worksheet: ExcelJS.Worksheet): Pick<SheetModel, "freezeTopRow" | "freezeFirstColumn"> {
  const views = Array.isArray(worksheet.views) ? worksheet.views : [];
  const frozenView = views.find(isFrozenWorksheetView);
  return {
    freezeTopRow: Number(frozenView?.ySplit ?? 0) > 0,
    freezeFirstColumn: Number(frozenView?.xSplit ?? 0) > 0
  };
}

function isFrozenWorksheetView(view: Partial<ExcelJS.WorksheetView>): view is FrozenWorksheetView {
  return view.state === "frozen";
}

async function applySheetProtectionToWorksheet(
  protection: SheetProtection | undefined,
  worksheet: ExcelJS.Worksheet
): Promise<void> {
  if (!protection) {
    return;
  }

  for (const [address, isUnlocked] of Object.entries(protection.unlockedCells ?? {})) {
    if (isUnlocked) {
      worksheet.getCell(address).protection = { locked: false };
    }
  }

  for (const [address, isLocked] of Object.entries(protection.lockedCells ?? {})) {
    if (isLocked) {
      worksheet.getCell(address).protection = { locked: true };
    }
  }

  if (protection.isProtected) {
    await worksheet.protect("", {
      selectLockedCells: true,
      selectUnlockedCells: true
    });
  }
}

function worksheetProtectionToSheetProtection(
  worksheet: ExcelJS.Worksheet,
  unlockedCells: Record<string, boolean>
): SheetProtection {
  const worksheetWithProtection = worksheet as unknown as {
    sheetProtection?: unknown;
    model?: { sheetProtection?: unknown };
  };
  const protectionModel = worksheetWithProtection.sheetProtection ?? worksheetWithProtection.model?.sheetProtection;
  const isProtected = isRecord(protectionModel) && protectionModel.sheet === true;

  return {
    isProtected,
    lockedCells: {},
    unlockedCells
  };
}

function cellContentToExcelValue(content: CellContent, hyperlink: string | undefined): ExcelJS.CellValue {
  if (content === null || content === "") {
    return hyperlink ? { text: hyperlink, hyperlink } : null;
  }
  if (typeof content === "string" && content.startsWith("=")) {
    return { formula: content.slice(1) };
  }
  if (hyperlink) {
    return { text: String(content), hyperlink };
  }
  return content;
}

function addNamedRangesToExcelWorkbook(
  namedRanges: readonly NamedRange[],
  workbook: ExcelJS.Workbook,
  worksheetNamesBySheetId: ReadonlyMap<string, string>
): void {
  for (const namedRange of namedRanges) {
    const sheetName = worksheetNamesBySheetId.get(namedRange.sheetId);
    if (!sheetName) {
      continue;
    }
    workbook.definedNames.add(`${quoteWorksheetName(sheetName)}!${formatAbsoluteRange(namedRange.range)}`, namedRange.name);
  }
}

function excelDefinedNamesToNamedRanges(
  definedNames: ExcelJS.DefinedNamesModel,
  sheetIdsByName: ReadonlyMap<string, string>
): NamedRange[] {
  const namedRanges: NamedRange[] = [];
  for (const definedName of definedNames) {
    if (!isImportableNamedRangeName(definedName.name) || definedName.ranges.length !== 1) {
      continue;
    }

    const namedRange = excelDefinedNameRangeToNamedRange(definedName.name, definedName.ranges[0], sheetIdsByName);
    if (namedRange) {
      namedRanges.push(namedRange);
    }
  }
  return namedRanges;
}

function excelDefinedNameRangeToNamedRange(
  name: string,
  rangeReference: string,
  sheetIdsByName: ReadonlyMap<string, string>
): NamedRange | null {
  const parts = splitSheetRangeReference(rangeReference);
  if (!parts) {
    return null;
  }

  const sheetId = sheetIdsByName.get(parts.sheetName);
  if (!sheetId) {
    return null;
  }

  try {
    return {
      name,
      sheetId,
      range: parseRangeAddress(parts.rangeAddress.replace(/\$/g, ""))
    };
  } catch {
    return null;
  }
}

function splitSheetRangeReference(reference: string): { sheetName: string; rangeAddress: string } | null {
  let isQuoted = false;
  for (let index = 0; index < reference.length; index += 1) {
    const character = reference[index];
    if (character === "'") {
      if (isQuoted && reference[index + 1] === "'") {
        index += 1;
      } else {
        isQuoted = !isQuoted;
      }
      continue;
    }
    if (character === "!" && !isQuoted) {
      return {
        sheetName: unquoteWorksheetName(reference.slice(0, index)),
        rangeAddress: reference.slice(index + 1)
      };
    }
  }
  return null;
}

function isImportableNamedRangeName(name: string): boolean {
  if (name.startsWith("_xlnm.") || !/^[A-Za-z_][A-Za-z0-9_]*$/.test(name)) {
    return false;
  }

  try {
    parseCellAddress(name);
    return false;
  } catch {
    return true;
  }
}

function formatAbsoluteRange(range: CellRange): string {
  const normalized = normalizeRange(range);
  const start = formatAbsoluteCell(normalized.start);
  const end = formatAbsoluteCell(normalized.end);
  return start === end ? start : `${start}:${end}`;
}

function formatAbsoluteCell(coord: { row: number; column: number }): string {
  return `$${columnIndexToName(coord.column)}$${coord.row + 1}`;
}

function formatRangeReference(range: CellRange): string {
  const normalized = normalizeRange(range);
  const start = formatCellAddress(normalized.start);
  const end = formatCellAddress(normalized.end);
  return start === end ? start : `${start}:${end}`;
}

function quoteWorksheetName(name: string): string {
  return `'${name.replace(/'/g, "''")}'`;
}

function unquoteWorksheetName(name: string): string {
  const trimmed = name.trim();
  if (trimmed.startsWith("'") && trimmed.endsWith("'")) {
    return trimmed.slice(1, -1).replace(/''/g, "'");
  }
  return trimmed;
}

function applyCellFormatToExcelCell(cell: ExcelJS.Cell, format: CellFormat): void {
  if (
    format.bold !== undefined ||
    format.italic !== undefined ||
    format.textColor ||
    format.fontFamily ||
    format.fontSize !== undefined
  ) {
    cell.font = {
      ...(cell.font ?? {}),
      bold: format.bold,
      italic: format.italic,
      name: format.fontFamily ?? cell.font?.name,
      size: format.fontSize ?? cell.font?.size,
      color: format.textColor ? { argb: hexToArgb(format.textColor) } : cell.font?.color
    };
  }

  if (format.backgroundColor) {
    cell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: hexToArgb(format.backgroundColor) }
    };
  }

  if (format.wrapText !== undefined || format.horizontalAlign || format.verticalAlign) {
    cell.alignment = {
      ...(cell.alignment ?? {}),
      wrapText: format.wrapText,
      horizontal: format.horizontalAlign,
      vertical: format.verticalAlign
    };
  }

  if (format.numberFormat && format.numberFormat !== "general") {
    cell.numFmt = XLSX_NUMBER_FORMATS[format.numberFormat];
  }

  if (format.borders) {
    cell.border = {
      top: excelBorderSide(format.borders.top),
      right: excelBorderSide(format.borders.right),
      bottom: excelBorderSide(format.borders.bottom),
      left: excelBorderSide(format.borders.left)
    };
  }
}

function conditionalFormatRuleToExcel(rule: ConditionalFormatRule, priority: number): ExcelJS.ConditionalFormattingRule | null {
  const style = cellFormatToExcelStyle(rule.format);

  if (rule.condition.type === "colorScale") {
    return {
      type: "colorScale",
      priority,
      cfvo: [{ type: "min" }, { type: "max" }],
      color: [{ argb: hexToArgb(rule.condition.minColor) }, { argb: hexToArgb(rule.condition.maxColor) }]
    };
  }

  if (rule.condition.type === "duplicate" || rule.condition.type === "unique") {
    return {
      type: "expression",
      formulae: [duplicateUniqueConditionToFormula(rule)],
      priority,
      style
    };
  }

  if (rule.condition.type === "top" || rule.condition.type === "bottom") {
    return {
      type: "top10",
      rank: Math.max(1, Math.floor(rule.condition.count)),
      percent: false,
      bottom: rule.condition.type === "bottom",
      priority,
      style
    };
  }

  if (rule.condition.type === "blank" || rule.condition.type === "notBlank") {
    return {
      type: "containsText",
      operator: rule.condition.type === "blank" ? "containsBlanks" : "notContainsBlanks",
      priority,
      style
    };
  }

  if (rule.condition.type === "greaterThan" || rule.condition.type === "lessThan" || rule.condition.type === "equalTo") {
    return {
      type: "cellIs",
      operator: rule.condition.type === "equalTo" ? "equal" : rule.condition.type,
      formulae: [rule.condition.value],
      priority,
      style
    };
  }

  if (rule.condition.type === "between") {
    return {
      type: "cellIs",
      operator: "between",
      formulae: [rule.condition.value, rule.condition.secondValue],
      priority,
      style
    };
  }

  if (rule.condition.type === "textContains") {
    return {
      type: "containsText",
      operator: "containsText",
      text: rule.condition.value,
      priority,
      style
    };
  }

  return null;
}

function worksheetConditionalFormatsToRules(worksheet: ExcelJS.Worksheet): ConditionalFormatRule[] {
  const conditionalFormats: ConditionalFormatRule[] = [];
  const models = getWorksheetConditionalFormattingModels(worksheet);
  let nextId = 1;

  for (const model of models) {
    const refs = typeof model.ref === "string" ? model.ref.split(/\s+/).filter(Boolean) : [];
    const rules = Array.isArray(model.rules) ? model.rules : [];
    for (const ref of refs) {
      const range = parseConditionalFormattingRange(ref);
      if (!range) {
        continue;
      }

      for (const excelRule of rules) {
        const rule = excelConditionalRuleToConditionalFormatRule(excelRule, range, nextId);
        if (rule) {
          conditionalFormats.push(rule);
          nextId += 1;
        }
      }
    }
  }

  return conditionalFormats;
}

type WorksheetConditionalFormattingModel = {
  ref?: unknown;
  rules?: unknown;
};

function getWorksheetConditionalFormattingModels(worksheet: ExcelJS.Worksheet): WorksheetConditionalFormattingModel[] {
  const worksheetWithConditionalFormats = worksheet as unknown as {
    conditionalFormattings?: unknown;
    model?: { conditionalFormattings?: unknown };
  };
  const conditionalFormattings =
    worksheetWithConditionalFormats.conditionalFormattings ??
    worksheetWithConditionalFormats.model?.conditionalFormattings ??
    [];
  return Array.isArray(conditionalFormattings)
    ? conditionalFormattings.filter((model): model is WorksheetConditionalFormattingModel => isRecord(model))
    : [];
}

function parseConditionalFormattingRange(ref: string): CellRange | null {
  try {
    return parseRangeAddress(ref.replace(/\$/g, ""));
  } catch {
    return null;
  }
}

function excelConditionalRuleToConditionalFormatRule(
  rule: unknown,
  range: CellRange,
  id: number
): ConditionalFormatRule | null {
  if (!isRecord(rule)) {
    return null;
  }

  const condition = excelConditionalRuleToCondition(rule);
  if (!condition) {
    return null;
  }

  return {
    id: `cf-${id}`,
    range,
    condition,
    format: excelStyleToCellFormat(isRecord(rule.style) ? rule.style : undefined)
  };
}

function excelConditionalRuleToCondition(rule: Record<string, unknown>): ConditionalFormatRule["condition"] | null {
  if (rule.type === "colorScale" && Array.isArray(rule.color)) {
    const colors = rule.color.map(excelColorToHex).filter((color): color is string => Boolean(color));
    const minColor = colors[0];
    const maxColor = colors.at(-1);
    if (minColor && maxColor) {
      return { type: "colorScale", minColor, maxColor };
    }
  }

  if (rule.type === "expression") {
    const formula = Array.isArray(rule.formulae) ? formulaValueToString(rule.formulae[0]) : "";
    const condition = parseDuplicateUniqueExpression(formula);
    if (condition) {
      return condition;
    }
  }

  if (rule.type === "top10" && typeof rule.rank === "number" && Number.isFinite(rule.rank)) {
    return {
      type: rule.bottom === true ? "bottom" : "top",
      count: Math.max(1, Math.floor(rule.rank))
    };
  }

  if (rule.type === "cellIs") {
    const formulae = Array.isArray(rule.formulae) ? rule.formulae.map(formulaValueToString) : [];
    const value = formulae[0];
    if (!value) {
      return null;
    }

    if (rule.operator === "greaterThan" || rule.operator === "lessThan") {
      return { type: rule.operator, value };
    }
    if (rule.operator === "equal") {
      return { type: "equalTo", value };
    }
    if (rule.operator === "between" && formulae[1]) {
      return { type: "between", value, secondValue: formulae[1] };
    }
  }

  if (rule.type === "containsText" && rule.operator === "containsBlanks") {
    return { type: "blank" };
  }

  if (rule.type === "containsText" && rule.operator === "notContainsBlanks") {
    return { type: "notBlank" };
  }

  if (rule.type === "containsText" && (rule.operator === undefined || rule.operator === "containsText")) {
    const value =
      typeof rule.text === "string"
        ? rule.text
        : parseContainsTextFormula(Array.isArray(rule.formulae) ? formulaValueToString(rule.formulae[0]) : "");
    return value ? { type: "textContains", value } : null;
  }

  return null;
}

function duplicateUniqueConditionToFormula(rule: ConditionalFormatRule): string {
  const normalizedRange = normalizeRange(rule.range);
  const activeCell = formatCellAddress(normalizedRange.start);
  const operator = rule.condition.type === "duplicate" ? ">" : "=";
  return `AND(LEN(TRIM(${activeCell}))>0,COUNTIF(${formatAbsoluteRange(normalizedRange)},${activeCell})${operator}1)`;
}

function parseDuplicateUniqueExpression(formula: string): ConditionalFormatRule["condition"] | null {
  const normalizedFormula = formula.replace(/\s+/g, "");
  const match = normalizedFormula.match(
    /^AND\(LEN\(TRIM\(\$?[A-Z]+\$?\d+\)\)>0,COUNTIF\(\$?[A-Z]+\$?\d+(?::\$?[A-Z]+\$?\d+)?,\$?[A-Z]+\$?\d+\)([>=])1\)$/i
  );

  if (!match) {
    return null;
  }

  return match[1] === ">" ? { type: "duplicate" } : { type: "unique" };
}

function formulaValueToString(value: unknown): string {
  const text = String(value ?? "").trim();
  if (text.startsWith('"') && text.endsWith('"')) {
    return text.slice(1, -1).replace(/""/g, '"');
  }
  return text;
}

function parseContainsTextFormula(formula: string): string {
  const match = formula.match(/SEARCH\("((?:[^"]|"")*)",/i);
  return match ? match[1].replace(/""/g, '"') : "";
}

function cellFormatToExcelStyle(format: CellFormat): Partial<ExcelJS.Style> {
  const style: Partial<ExcelJS.Style> = {};
  const font: Partial<ExcelJS.Font> = {};

  if (format.bold !== undefined) {
    font.bold = format.bold;
  }
  if (format.italic !== undefined) {
    font.italic = format.italic;
  }
  if (format.fontFamily) {
    font.name = format.fontFamily;
  }
  if (format.fontSize !== undefined) {
    font.size = format.fontSize;
  }
  if (format.textColor) {
    font.color = { argb: hexToArgb(format.textColor) };
  }
  if (Object.keys(font).length > 0) {
    style.font = font;
  }

  if (format.backgroundColor) {
    style.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: hexToArgb(format.backgroundColor) }
    };
  }

  if (format.wrapText !== undefined || format.horizontalAlign || format.verticalAlign) {
    style.alignment = {
      wrapText: format.wrapText,
      horizontal: format.horizontalAlign,
      vertical: format.verticalAlign
    };
  }

  if (format.numberFormat && format.numberFormat !== "general") {
    style.numFmt = XLSX_NUMBER_FORMATS[format.numberFormat];
  }

  if (format.borders) {
    style.border = {
      top: excelBorderSide(format.borders.top),
      right: excelBorderSide(format.borders.right),
      bottom: excelBorderSide(format.borders.bottom),
      left: excelBorderSide(format.borders.left)
    };
  }

  return style;
}

function dataValidationRuleToExcel(rule: DataValidationRule): ExcelJS.DataValidation | null {
  const allowBlank = rule.allowBlank !== false;
  if (rule.type === "list") {
    return {
      type: "list",
      formulae: [`"${rule.values.map(escapeListValidationValue).join(",")}"`],
      allowBlank,
      showErrorMessage: true,
      errorStyle: "error"
    };
  }

  if (rule.type === "textLength") {
    const formulae = [rule.min, rule.max].filter((value): value is number => value !== undefined);
    return {
      type: "textLength",
      operator: boundedValidationOperator(rule),
      formulae,
      allowBlank,
      showErrorMessage: true,
      errorStyle: "error"
    };
  }

  const formulae = [rule.min, rule.max].filter((value): value is number => value !== undefined);
  return {
    type: "decimal",
    operator: boundedValidationOperator(rule),
    formulae,
    allowBlank,
    showErrorMessage: true,
    errorStyle: "error"
  };
}

function excelDataValidationToRule(validation: ExcelJS.DataValidation | undefined): DataValidationRule | null {
  if (!validation) {
    return null;
  }

  if (validation.type === "list") {
    const values = parseListValidationFormula(validation.formulae?.[0]);
    return values.length > 0
      ? {
          type: "list",
          values,
          ...dataValidationAllowBlankFlag(validation)
        }
      : null;
  }

  if (validation.type === "decimal" || validation.type === "whole") {
    const rule = numericDataValidationToRule(validation);
    return { ...rule, ...dataValidationAllowBlankFlag(validation) };
  }

  if (validation.type === "textLength") {
    const rule = textLengthDataValidationToRule(validation);
    return { ...rule, ...dataValidationAllowBlankFlag(validation) };
  }

  return null;
}

function boundedValidationOperator(
  rule: Extract<DataValidationRule, { type: "number" }> | Extract<DataValidationRule, { type: "textLength" }>
): ExcelJS.DataValidationOperator | undefined {
  if (rule.min !== undefined && rule.max !== undefined) {
    return "between";
  }
  if (rule.min !== undefined) {
    return "greaterThanOrEqual";
  }
  if (rule.max !== undefined) {
    return "lessThanOrEqual";
  }
  return undefined;
}

function dataValidationAllowBlankFlag(validation: ExcelJS.DataValidation): Pick<DataValidationRule, "allowBlank"> {
  return validation.allowBlank === true ? {} : { allowBlank: false };
}

function numericDataValidationToRule(validation: ExcelJS.DataValidation): Extract<DataValidationRule, { type: "number" }> {
  const firstValue = parseValidationNumber(validation.formulae?.[0]);
  const secondValue = parseValidationNumber(validation.formulae?.[1]);
  const rule: Extract<DataValidationRule, { type: "number" }> = { type: "number" };

  if (validation.operator === "between") {
    if (firstValue !== undefined) {
      rule.min = firstValue;
    }
    if (secondValue !== undefined) {
      rule.max = secondValue;
    }
    return rule;
  }

  if (validation.operator === "greaterThanOrEqual" && firstValue !== undefined) {
    rule.min = firstValue;
  }
  if (validation.operator === "lessThanOrEqual" && firstValue !== undefined) {
    rule.max = firstValue;
  }

  return rule;
}

function textLengthDataValidationToRule(validation: ExcelJS.DataValidation): Extract<DataValidationRule, { type: "textLength" }> {
  const firstValue = parseValidationNumber(validation.formulae?.[0]);
  const secondValue = parseValidationNumber(validation.formulae?.[1]);
  const rule: Extract<DataValidationRule, { type: "textLength" }> = { type: "textLength" };

  if (validation.operator === "between") {
    if (firstValue !== undefined) {
      rule.min = firstValue;
    }
    if (secondValue !== undefined) {
      rule.max = secondValue;
    }
    return rule;
  }

  if (validation.operator === "greaterThanOrEqual" && firstValue !== undefined) {
    rule.min = firstValue;
  }
  if (validation.operator === "lessThanOrEqual" && firstValue !== undefined) {
    rule.max = firstValue;
  }

  return rule;
}

function parseValidationNumber(value: unknown): number | undefined {
  const numberValue = typeof value === "number" ? value : Number(value);
  return Number.isFinite(numberValue) ? numberValue : undefined;
}

function escapeListValidationValue(value: string): string {
  return value.replace(/"/g, '""');
}

function parseListValidationFormula(value: unknown): string[] {
  if (typeof value !== "string") {
    return [];
  }
  const unquotedValue = value.startsWith('"') && value.endsWith('"') ? value.slice(1, -1).replace(/""/g, '"') : value;
  return unquotedValue
    .split(",")
    .map((option) => option.trim())
    .filter(Boolean);
}

type ExcelStyleParts = {
  font?: Partial<ExcelJS.Font>;
  fill?: ExcelJS.Fill;
  numFmt?: string;
  alignment?: Partial<ExcelJS.Alignment>;
  border?: Partial<ExcelJS.Borders>;
};

const DEFAULT_XLSX_FONT_NAMES = new Set(["Calibri", "Aptos", "Aptos Narrow"]);
const DEFAULT_XLSX_FONT_SIZE = 11;

function excelStyleToCellFormat(style: ExcelStyleParts | undefined): CellFormat {
  const format: CellFormat = {};

  if (style?.font?.bold === true) {
    format.bold = true;
  }
  if (style?.font?.italic === true) {
    format.italic = true;
  }
  // Excel stamps its default font (Calibri/Aptos 11) on every styled cell; importing
  // that verbatim would materialize a format entry per cell on large files. Only the
  // FULL default signature (default name AND default/absent size together) is
  // treated as the stamp — a lone deviation like "Calibri 14" or "Arial 11" is a
  // deliberate choice and both fields are kept. Known limitation: a user who
  // explicitly picks exactly Calibri 11 is indistinguishable from the stamp.
  const importedFontName =
    typeof style?.font?.name === "string" && style.font.name.trim() !== "" ? style.font.name : undefined;
  const importedFontSize =
    typeof style?.font?.size === "number" && Number.isFinite(style.font.size) && style.font.size > 0
      ? style.font.size
      : undefined;
  const isDefaultFontStamp =
    (importedFontName === undefined || DEFAULT_XLSX_FONT_NAMES.has(importedFontName)) &&
    (importedFontSize === undefined || importedFontSize === DEFAULT_XLSX_FONT_SIZE);
  if (importedFontName !== undefined && !isDefaultFontStamp) {
    format.fontFamily = importedFontName;
  }
  if (importedFontSize !== undefined && !isDefaultFontStamp) {
    format.fontSize = importedFontSize;
  }

  const textColor = excelColorToHex(style?.font?.color);
  if (textColor) {
    format.textColor = textColor;
  }

  const backgroundColor = excelFillToHex(style?.fill);
  if (backgroundColor) {
    format.backgroundColor = backgroundColor;
  }

  const numberFormat = excelNumberFormatToCellFormat(style?.numFmt);
  if (numberFormat) {
    format.numberFormat = numberFormat;
  }

  if (isHorizontalAlign(style?.alignment?.horizontal)) {
    format.horizontalAlign = style.alignment.horizontal;
  }
  if (isVerticalAlign(style?.alignment?.vertical)) {
    format.verticalAlign = style.alignment.vertical;
  }
  if (style?.alignment?.wrapText === true) {
    format.wrapText = true;
  }

  const borders = excelBordersToCellBorders(style?.border);
  if (borders) {
    format.borders = borders;
  }

  return format;
}

function excelCellToCellFormat(cell: ExcelJS.Cell): CellFormat {
  return excelStyleToCellFormat(cell);
}

function excelBorderSide(side: CellBorderSide | undefined): Partial<ExcelJS.Border> | undefined {
  return side ? { style: side.style, color: { argb: hexToArgb(side.color) } } : undefined;
}

function excelBordersToCellBorders(borders: Partial<ExcelJS.Borders> | undefined): CellBorders | undefined {
  if (!borders) {
    return undefined;
  }

  const cellBorders: CellBorders = {
    top: excelBorderToCellBorderSide(borders.top),
    right: excelBorderToCellBorderSide(borders.right),
    bottom: excelBorderToCellBorderSide(borders.bottom),
    left: excelBorderToCellBorderSide(borders.left)
  };

  return Object.values(cellBorders).some(Boolean) ? cellBorders : undefined;
}

function excelBorderToCellBorderSide(border: Partial<ExcelJS.Border> | undefined): CellBorderSide | undefined {
  if (border?.style !== "thin") {
    return undefined;
  }
  return {
    style: "thin",
    color: excelColorToHex(border.color) ?? "#64748b"
  };
}

function excelFillToHex(fill: ExcelJS.Fill | undefined): string | undefined {
  if (!isRecord(fill) || fill.type !== "pattern" || fill.pattern !== "solid") {
    return undefined;
  }
  return excelColorToHex(fill.fgColor);
}

function excelNumberFormatToCellFormat(numFmt: string | undefined): CellFormat["numberFormat"] | undefined {
  if (!numFmt || numFmt === XLSX_NUMBER_FORMATS.general) {
    return undefined;
  }
  for (const [format, excelFormat] of Object.entries(XLSX_NUMBER_FORMATS)) {
    if (numFmt === excelFormat) {
      return format as CellFormat["numberFormat"];
    }
  }
  return undefined;
}

function isHorizontalAlign(value: unknown): value is NonNullable<CellFormat["horizontalAlign"]> {
  return value === "left" || value === "center" || value === "right";
}

function isVerticalAlign(value: unknown): value is NonNullable<CellFormat["verticalAlign"]> {
  return value === "top" || value === "middle" || value === "bottom";
}

function hexToArgb(color: string): string {
  return `FF${color.replace("#", "").toUpperCase()}`;
}

function excelColorToHex(color: unknown): string | undefined {
  if (!isRecord(color) || typeof color.argb !== "string") {
    return undefined;
  }

  const value = color.argb.replace(/^#/, "").toUpperCase();
  const rgb = value.length === 8 ? value.slice(2) : value;
  return /^[0-9A-F]{6}$/.test(rgb) ? `#${rgb.toLowerCase()}` : undefined;
}

function excelCellToCellContent(cell: ExcelJS.Cell): CellContent {
  const value = cell.value;
  // Shared-formula cells: cell.value.sharedFormula is the MASTER CELL'S ADDRESS, not a
  // formula. ExcelJS's cell.formula getter translates the master formula to this cell's
  // position, which is what a 1:1 import needs.
  if (isSharedFormulaValue(value)) {
    const translated = cell.formula;
    return typeof translated === "string" && translated.length > 0
      ? `=${translated}`
      : excelValueToCellContent(value.result ?? null);
  }
  return excelValueToCellContent(value);
}

function excelValueToCellContent(value: ExcelJS.CellValue): CellContent {
  if (value === null || value === undefined || value === "") {
    return null;
  }
  if (typeof value === "string" || typeof value === "number" || typeof value === "boolean") {
    return value;
  }
  if (value instanceof Date) {
    return dateToCellContent(value);
  }
  if (isFormulaValue(value)) {
    return `=${value.formula}`;
  }
  if (isSharedFormulaValue(value)) {
    // value.sharedFormula is the master cell's ADDRESS — emitting "=<address>"
    // would misrepresent the computation. Shared formulas are translated in
    // excelCellToCellContent (which has the Cell); here fall back to the result.
    return excelValueToCellContent(value.result ?? null);
  }
  if (isHyperlinkValue(value)) {
    return value.text || value.hyperlink;
  }
  if (isRichTextValue(value)) {
    return value.richText.map((segment) => segment.text).join("");
  }
  if (isErrorValue(value)) {
    return value.error;
  }
  return String(value);
}

// Midnight-UTC dates (the common case for Excel date cells) become plain ISO dates the
// formula engine parses as date serials; anything with a time keeps the full timestamp.
function dateToCellContent(value: Date): string {
  const iso = value.toISOString();
  return iso.endsWith("T00:00:00.000Z") ? iso.slice(0, 10) : iso;
}

function excelValueToHyperlink(value: ExcelJS.CellValue): string {
  return isHyperlinkValue(value) ? value.hyperlink : "";
}

function cellText(value: ExcelJS.CellValue): string {
  const content = excelValueToCellContent(value);
  return content === null ? "" : String(content);
}

function noteToText(note: ExcelJS.Cell["note"]): string {
  if (!note) {
    return "";
  }
  if (typeof note === "string") {
    return note;
  }
  return note.texts?.map((text) => text.text).join("").trim() ?? "";
}

function isFormulaValue(value: ExcelJS.CellValue): value is ExcelJS.CellFormulaValue {
  return isRecord(value) && typeof value.formula === "string";
}

function isSharedFormulaValue(value: ExcelJS.CellValue): value is ExcelJS.CellSharedFormulaValue {
  return isRecord(value) && typeof value.sharedFormula === "string";
}

function isHyperlinkValue(value: ExcelJS.CellValue): value is ExcelJS.CellHyperlinkValue {
  return isRecord(value) && typeof value.hyperlink === "string";
}

function isRichTextValue(value: ExcelJS.CellValue): value is ExcelJS.CellRichTextValue {
  return isRecord(value) && Array.isArray(value.richText);
}

function isErrorValue(value: ExcelJS.CellValue): value is ExcelJS.CellErrorValue {
  return isRecord(value) && typeof value.error === "string";
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

function normalizeRange(range: CellRange): CellRange {
  return {
    start: {
      row: Math.min(range.start.row, range.end.row),
      column: Math.min(range.start.column, range.end.column)
    },
    end: {
      row: Math.max(range.start.row, range.end.row),
      column: Math.max(range.start.column, range.end.column)
    }
  };
}

function safeWorksheetName(name: string, workbook: ExcelJS.Workbook): string {
  const usedNames = new Set(workbook.worksheets.map((worksheet) => worksheet.name.toLowerCase()));
  const baseName = (name.replace(/[:\\/?*[\]]/g, " ").trim() || "Sheet").slice(0, 31);
  let candidate = baseName;
  let suffix = 1;

  while (usedNames.has(candidate.toLowerCase())) {
    const suffixText = ` ${suffix}`;
    candidate = `${baseName.slice(0, 31 - suffixText.length)}${suffixText}`;
    suffix += 1;
  }

  return candidate;
}

function toArrayBuffer(output: ArrayBuffer | Uint8Array): ArrayBuffer {
  if (output instanceof ArrayBuffer) {
    return output;
  }

  const copy = new Uint8Array(output.byteLength);
  copy.set(output);
  return copy.buffer;
}

function loadExcelJs(): Promise<typeof ExcelJS> {
  excelJsPromise ??= import("exceljs").then((module) => module.default);
  return excelJsPromise;
}
