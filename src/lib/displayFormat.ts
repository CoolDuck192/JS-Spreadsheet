import type { CellFormat } from "../types";
import { excelSerialToDate, parseExcelTemporalInput } from "../core/values/excelDate";

const NUMBER_FORMATTER = new Intl.NumberFormat("en-US", {
  maximumFractionDigits: 2
});
const CURRENCY_FORMATTER = new Intl.NumberFormat("en-US", {
  style: "currency",
  currency: "USD",
  minimumFractionDigits: 2,
  maximumFractionDigits: 2
});
const PERCENT_FORMATTER = new Intl.NumberFormat("en-US", {
  style: "percent",
  maximumFractionDigits: 2
});
const FINANCIAL_FORMATTER = new Intl.NumberFormat("en-US", {
  maximumFractionDigits: 0
});
const FINANCIAL2_FORMATTER = new Intl.NumberFormat("en-US", {
  minimumFractionDigits: 2,
  maximumFractionDigits: 2
});
const DATE_FORMATTER = new Intl.DateTimeFormat("en-US", {
  month: "short",
  day: "numeric",
  year: "numeric",
  timeZone: "UTC"
});
const DATE_TIME_FORMATTER = new Intl.DateTimeFormat("en-US", {
  month: "short",
  day: "numeric",
  year: "numeric",
  hour: "numeric",
  minute: "2-digit",
  timeZone: "UTC"
});
const TIME_FORMATTER = new Intl.DateTimeFormat("en-US", {
  hour: "numeric",
  minute: "2-digit",
  timeZone: "UTC"
});
const MILLISECONDS_PER_DAY = 24 * 60 * 60 * 1000;

export function formatDisplayValue(value: string, format: CellFormat | undefined): string {
  if (!value || !format?.numberFormat || format.numberFormat === "general" || value.startsWith("#")) {
    return value;
  }

  if (format.numberFormat === "date" || format.numberFormat === "dateTime") {
    const serial = parseDateSerial(value);
    if (serial !== null && serial >= 60 && serial < 61) {
      if (format.numberFormat === "date") {
        return "Feb 29, 1900";
      }
      const time = new Date(Date.UTC(1970, 0, 1) + (serial - 60) * MILLISECONDS_PER_DAY);
      return `Feb 29, 1900, ${TIME_FORMATTER.format(time)}`;
    }
    const date = serial === null ? null : excelSerialToDate(serial);
    if (!date) {
      return value;
    }
    return format.numberFormat === "date" ? DATE_FORMATTER.format(date) : DATE_TIME_FORMATTER.format(date);
  }

  const number = parseNumberValue(value);
  if (number === null) {
    return value;
  }

  if (format.numberFormat === "number") {
    return NUMBER_FORMATTER.format(number);
  }

  if (format.numberFormat === "currency") {
    return CURRENCY_FORMATTER.format(number);
  }

  if (format.numberFormat === "financial") {
    return parenthesizeNegative(FINANCIAL_FORMATTER.format(Math.abs(number)), number < 0);
  }

  if (format.numberFormat === "financial2") {
    return parenthesizeNegative(FINANCIAL2_FORMATTER.format(Math.abs(number)), number < 0);
  }

  if (format.numberFormat === "accounting") {
    // Deliberately matches Excel numFmt section-selection: zero-detection is EXACT (number === 0) and
    // parenthesization is SIGN-BASED pre-rounding — so -0.004 renders ($0.00), not "$ -", matching what
    // the exported XLSX renders in Excel/LibreOffice; on-screen and exported output agree by construction.
    if (number === 0) {
      return "$ -";
    }
    return parenthesizeNegative(`$${FINANCIAL2_FORMATTER.format(Math.abs(number))}`, number < 0);
  }

  if (format.numberFormat === "percent") {
    return PERCENT_FORMATTER.format(number);
  }

  return value;
}

function parenthesizeNegative(formatted: string, negative: boolean): string {
  return negative ? `(${formatted})` : formatted;
}

function parseNumberValue(value: string): number | null {
  let normalized = value.trim();
  if (!normalized) {
    return null;
  }

  let isNegative = false;
  if (/^\(.*\)$/.test(normalized)) {
    isNegative = true;
    normalized = normalized.slice(1, -1).trim();
  }

  const isPercent = normalized.endsWith("%");
  if (isPercent) {
    normalized = normalized.slice(0, -1).trim();
  }

  normalized = normalized.replace(/[,$£€¥₹\s]/g, "");
  if (!normalized || /[A-Za-z]/.test(normalized)) {
    return null;
  }

  const parsed = Number(normalized);
  if (!Number.isFinite(parsed)) {
    return null;
  }

  const signed = isNegative ? parsed * -1 : parsed;
  return isPercent ? signed / 100 : signed;
}

function parseDateSerial(value: string): number | null {
  const trimmed = value.trim();
  if (!trimmed) {
    return null;
  }

  const serial = parseNumberValue(trimmed);
  if (serial !== null && serial >= 0) {
    return serial;
  }

  const temporal = parseExcelTemporalInput(trimmed);
  return temporal?.serial ?? null;
}
