import type { CellFormat } from "../types";

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
const DATE_FORMATTER = new Intl.DateTimeFormat("en-US", {
  month: "short",
  day: "numeric",
  year: "numeric",
  timeZone: "UTC"
});

export function formatDisplayValue(value: string, format: CellFormat | undefined): string {
  if (!value || !format?.numberFormat || format.numberFormat === "general" || value.startsWith("#")) {
    return value;
  }

  if (format.numberFormat === "date") {
    const date = parseDateValue(value);
    return date ? DATE_FORMATTER.format(date) : value;
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

  return PERCENT_FORMATTER.format(number);
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

function parseDateValue(value: string): Date | null {
  const trimmed = value.trim();
  if (!trimmed) {
    return null;
  }

  const serial = parseNumberValue(trimmed);
  if (serial !== null && serial > 0) {
    const excelEpoch = Date.UTC(1899, 11, 30);
    return new Date(excelEpoch + serial * 24 * 60 * 60 * 1000);
  }

  const isoMatch = trimmed.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (isoMatch) {
    return new Date(Date.UTC(Number(isoMatch[1]), Number(isoMatch[2]) - 1, Number(isoMatch[3])));
  }

  const parsed = Date.parse(trimmed);
  if (!Number.isFinite(parsed)) {
    return null;
  }
  return new Date(parsed);
}
