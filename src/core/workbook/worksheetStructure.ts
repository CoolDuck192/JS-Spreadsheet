import type { WorkbookModel } from "../../types";
import { shiftSheetStructurePlanes } from "../../lib/workbook";
import type { TableIssue } from "../commands/types";
import type { IdGenerator } from "../ids";

export type WorksheetStructureCommand =
  | { type: "rows.insert" | "rows.delete"; sheetId: string; index: number; count: number }
  | {
      type: "columns.insert";
      sheetId: string;
      index: number;
      count: number;
      expandTableIds?: readonly string[];
    }
  | { type: "columns.delete"; sheetId: string; index: number; count: number };

export type WorksheetStructureServices = Readonly<{ createId: IdGenerator }>;

export type WorksheetStructureReduction =
  | { status: "committed"; workbook: WorkbookModel }
  | {
      status: "rejected";
      reason: "validation" | "permission";
      workbook: WorkbookModel;
      issues: readonly TableIssue[];
    };

type WorksheetStructureIssueCode =
  | "SHEET_STRUCTURE_INDEX_INVALID"
  | "SHEET_STRUCTURE_COUNT_INVALID"
  | "SHEET_STRUCTURE_OUT_OF_BOUNDS"
  | "TABLE_PROTECTED"
  | "TABLE_PARTIAL_STRUCTURAL_EDIT"
  | "TABLE_EXPANSION_CONTEXT_INVALID"
  | "TABLE_RANGE_OVERLAP"
  | "TABLE_COLUMN_ID_CONFLICT";

export function reduceWorksheetStructureCommand(
  workbook: WorkbookModel,
  command: WorksheetStructureCommand,
  services: WorksheetStructureServices
): WorksheetStructureReduction {
  void services;
  const sheet = workbook.sheets.find((candidate) => candidate.id === command.sheetId);
  if (!sheet) {
    return rejected(
      workbook,
      "validation",
      "SHEET_STRUCTURE_OUT_OF_BOUNDS",
      "Worksheet does not exist"
    );
  }
  if (!Number.isInteger(command.index) || command.index < 0) {
    return rejected(
      workbook,
      "validation",
      "SHEET_STRUCTURE_INDEX_INVALID",
      "Structure index must be a non-negative integer"
    );
  }
  if (!Number.isInteger(command.count) || command.count <= 0) {
    return rejected(
      workbook,
      "validation",
      "SHEET_STRUCTURE_COUNT_INVALID",
      "Structure count must be a positive integer"
    );
  }
  const dimension = command.type.startsWith("rows.") ? sheet.rowCount : sheet.columnCount;
  const withinBounds = command.type.endsWith(".insert")
    ? command.index <= dimension
    : command.index < dimension && command.index + command.count <= dimension;
  if (!withinBounds) {
    return rejected(
      workbook,
      "validation",
      "SHEET_STRUCTURE_OUT_OF_BOUNDS",
      "Structure edit is outside the worksheet"
    );
  }
  if (sheet.protection.isProtected) {
    return rejected(
      workbook,
      "permission",
      "TABLE_PROTECTED",
      "Protected sheets cannot change worksheet structure"
    );
  }
  if (workbook.tables.some((table) => table.sheetId === command.sheetId)) {
    return rejected(
      workbook,
      "validation",
      "TABLE_PARTIAL_STRUCTURAL_EDIT",
      "Structured tables require table-aware structure editing"
    );
  }
  return {
    status: "committed",
    workbook: shiftSheetStructurePlanes(workbook, command.sheetId, operationFor(command))
  };
}

export function isWorksheetStructureCommand(command: unknown): command is WorksheetStructureCommand {
  if (!command || typeof command !== "object" || !("type" in command)) {
    return false;
  }
  const type = command.type;
  return type === "rows.insert"
    || type === "rows.delete"
    || type === "columns.insert"
    || type === "columns.delete";
}

function rejected(
  workbook: WorkbookModel,
  reason: "validation" | "permission",
  code: WorksheetStructureIssueCode,
  message: string
): WorksheetStructureReduction {
  return {
    status: "rejected",
    reason,
    workbook,
    issues: [{ code, message }]
  };
}

function operationFor(command: WorksheetStructureCommand) {
  return {
    axis: command.type.startsWith("rows.") ? "row" as const : "column" as const,
    mode: command.type.endsWith(".insert") ? "insert" as const : "delete" as const,
    index: command.index,
    count: command.count
  };
}
