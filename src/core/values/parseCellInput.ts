import type { CellContent, CellFormat } from "../../types";
import { parseExcelTemporalInput } from "./excelDate";

export type ParsedCellInput = {
  raw: string;
  kind: "blank" | "formula" | "number" | "boolean" | "date" | "dateTime" | "text";
  stored: CellContent;
  formula?: string;
  inferredNumberFormat?: CellFormat["numberFormat"];
};

export function parseCellInput(raw: string): ParsedCellInput {
  if (raw === "") {
    return { raw, kind: "blank", stored: null };
  }
  if (raw.startsWith("'")) {
    return { raw, kind: "text", stored: raw.slice(1) };
  }
  if (raw.startsWith("=")) {
    return { raw, kind: "formula", stored: raw, formula: raw };
  }
  if (/^(TRUE|FALSE)$/i.test(raw)) {
    return { raw, kind: "boolean", stored: raw.toUpperCase() === "TRUE" };
  }

  const temporal = parseExcelTemporalInput(raw);
  if (temporal) {
    return {
      raw,
      kind: temporal.kind,
      stored: temporal.serial,
      inferredNumberFormat: temporal.kind
    };
  }

  const number = parseUnambiguousNumber(raw);
  if (number !== null) {
    return { raw, kind: "number", stored: number };
  }
  return { raw, kind: "text", stored: raw };
}

function parseUnambiguousNumber(raw: string): number | null {
  const value = raw.trim();
  if (!value) {
    return null;
  }

  const isPlainDecimal = /^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:e[+-]?\d+)?$/i.test(value);
  const isUsGroupedDecimal = /^[+-]?\d{1,3}(?:,\d{3})+(?:\.\d+)?(?:e[+-]?\d+)?$/i.test(value);
  if (!isPlainDecimal && !isUsGroupedDecimal) {
    return null;
  }

  const parsed = Number(value.replace(/,/g, ""));
  return Number.isFinite(parsed) ? parsed : null;
}
