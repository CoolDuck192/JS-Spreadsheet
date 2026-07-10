import type { FilterExpression, QueryScalar } from "../../table/core/query";
import type { StructuredTable, WorkbookModel } from "../../types";
import type { ComputedCellValue } from "../../lib/formulaEngine";
import { compareDeterministicText, isRowVisibleForFilter } from "../../lib/filters";
import { formatCellAddress } from "../../lib/addressing";
import { excelSerialToDate } from "../values/excelDate";
import { getStructuredTableBodyRange } from "./structuredTables";

export function matchesStructuredTableFilter(
  expression: FilterExpression,
  getValue: (columnId: string) => QueryScalar
): boolean {
  switch (expression.kind) {
    case "logical":
      return expression.operator === "and"
        ? expression.operands.every((operand) => matchesStructuredTableFilter(operand, getValue))
        : expression.operands.some((operand) => matchesStructuredTableFilter(operand, getValue));
    case "not":
      return !matchesStructuredTableFilter(expression.operand, getValue);
    case "blank": {
      const value = getValue(expression.columnId);
      const isNull = value.type === "null";
      const isEmpty = value.type === "string" && value.value === "";
      switch (expression.operator) {
        case "isNull": return isNull;
        case "isNotNull": return !isNull;
        case "isEmpty": return isEmpty;
        case "isNotEmpty": return !isEmpty;
        case "isBlank": return isNull || isEmpty;
        case "isNotBlank": return !isNull && !isEmpty;
      }
    }
    case "set": {
      const value = getValue(expression.columnId);
      const included = expression.values.some((candidate) => scalarEquals(value, candidate));
      return expression.operator === "in" ? included : !included;
    }
    case "range": {
      const value = getValue(expression.columnId);
      const included = compareScalars(value, expression.lower) >= 0
        && compareScalars(value, expression.upper) <= 0;
      return expression.operator === "between" ? included : !included;
    }
    case "comparison": {
      const value = getValue(expression.columnId);
      switch (expression.operator) {
        case "eq": return scalarEquals(value, expression.value);
        case "neq": return !scalarEquals(value, expression.value);
        case "contains": return fold(scalarText(value)).includes(fold(scalarText(expression.value)));
        case "startsWith": return fold(scalarText(value)).startsWith(fold(scalarText(expression.value)));
        case "endsWith": return fold(scalarText(value)).endsWith(fold(scalarText(expression.value)));
        case "gt": return compareScalars(value, expression.value) > 0;
        case "gte": return compareScalars(value, expression.value) >= 0;
        case "lt": return compareScalars(value, expression.value) < 0;
        case "lte": return compareScalars(value, expression.value) <= 0;
      }
    }
  }
}

export function isStructuredTableRowVisible(
  workbook: WorkbookModel,
  table: StructuredTable,
  sheetRow: number,
  getEvaluation: (sheetId: string, address: string) => ComputedCellValue
): boolean {
  const body = getStructuredTableBodyRange(table);
  if (!body || sheetRow < body.start.row || sheetRow > body.end.row) return true;
  const sheet = workbook.sheets.find((candidate) => candidate.id === table.sheetId);
  if (!sheet) return false;
  if (sheet.hiddenRows?.[String(sheetRow)] === true) return false;
  if (!sheet.filters.every((filter) => isRowVisibleForFilter(
    sheetRow,
    filter,
    (row, column) => getEvaluation(table.sheetId, formatCellAddress({ row, column }))
  ))) return false;
  if (!table.filter) return true;
  return matchesStructuredTableFilter(table.filter, (columnId) => {
    const column = table.columns.find((candidate) => candidate.id === columnId);
    if (!column) return { type: "null" };
    const value = getEvaluation(
      table.sheetId,
      formatCellAddress({ row: sheetRow, column: column.sheetColumn })
    );
    return toQueryScalar(value, column.dataType);
  });
}

function toQueryScalar(
  value: ComputedCellValue,
  dataType: StructuredTable["columns"][number]["dataType"]
): QueryScalar {
  if (value === null) return { type: "null" };
  if (typeof value === "object") return { type: "error", value: value.code };
  if (typeof value === "boolean") return { type: "boolean", value };
  if (typeof value === "number") {
    if (dataType === "date" || dataType === "datetime") {
      const date = excelSerialToDate(value);
      if (date) {
        return dataType === "date"
          ? { type: "date", value: date.toISOString().slice(0, 10) }
          : { type: "datetime", value: date.toISOString() };
      }
    }
    return { type: "number", value };
  }
  return { type: "string", value };
}

function scalarEquals(left: QueryScalar, right: QueryScalar): boolean {
  if (left.type !== right.type) return false;
  if (left.type === "null") return true;
  return left.value === (right as Exclude<QueryScalar, { type: "null" }>).value;
}

function compareScalars(left: QueryScalar, right: QueryScalar): number {
  const leftRank = scalarRank(left);
  const rightRank = scalarRank(right);
  if (leftRank !== rightRank) return leftRank - rightRank;
  if (left.type === "null" || right.type === "null") return 0;
  const leftComparable = scalarComparable(left);
  const rightComparable = scalarComparable(right);
  if (typeof leftComparable === "number" && typeof rightComparable === "number") {
    return leftComparable - rightComparable;
  }
  return compareDeterministicText(String(leftComparable), String(rightComparable));
}

function scalarRank(value: QueryScalar): number {
  switch (value.type) {
    case "number":
    case "date":
    case "datetime": return 0;
    case "boolean": return 1;
    case "string": return 2;
    case "error": return 3;
    case "null": return 4;
  }
}

function scalarComparable(value: Exclude<QueryScalar, { type: "null" }>): number | string {
  switch (value.type) {
    case "number": return value.value;
    case "date":
    case "datetime": return Date.parse(value.value);
    case "boolean": return value.value ? 1 : 0;
    case "string":
    case "error": return value.value;
  }
}

function scalarText(value: QueryScalar): string {
  if (value.type === "null") return "";
  if (value.type === "boolean") return value.value ? "TRUE" : "FALSE";
  return String(value.value);
}

function fold(value: string): string {
  return value.toLowerCase();
}
