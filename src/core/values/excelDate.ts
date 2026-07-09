const MILLISECONDS_PER_DAY = 24 * 60 * 60 * 1000;
const EXCEL_EPOCH_UTC = Date.UTC(1899, 11, 31);
const EXCEL_LEAP_DAY_SERIAL = 60;
const TWO_DIGIT_YEAR_PIVOT = 30;

export type ParsedExcelTemporalInput = {
  kind: "date" | "dateTime";
  serial: number;
};

/** Converts a real UTC date to Excel's 1900-system serial number. */
export function dateToExcelSerial(value: Date): number {
  const timestamp = value.getTime();
  if (!Number.isFinite(timestamp)) {
    return Number.NaN;
  }

  const serialWithoutCompatibilityDay = (timestamp - EXCEL_EPOCH_UTC) / MILLISECONDS_PER_DAY;
  return serialWithoutCompatibilityDay >= EXCEL_LEAP_DAY_SERIAL
    ? serialWithoutCompatibilityDay + 1
    : serialWithoutCompatibilityDay;
}

/** Converts an Excel 1900-system serial to a UTC Date. Serial 60 maps to Feb 28 because JS has no phantom Feb 29. */
export function excelSerialToDate(serial: number): Date | null {
  if (!Number.isFinite(serial)) {
    return null;
  }

  const serialWithoutCompatibilityDay = serial >= EXCEL_LEAP_DAY_SERIAL ? serial - 1 : serial;
  const timestamp = EXCEL_EPOCH_UTC + serialWithoutCompatibilityDay * MILLISECONDS_PER_DAY;
  const result = new Date(timestamp);
  return Number.isFinite(result.getTime()) ? result : null;
}

export function parseExcelTemporalInput(raw: string): ParsedExcelTemporalInput | null {
  const value = raw.trim();
  if (!value) {
    return null;
  }

  const dateTime = parseIsoDateTime(value);
  if (dateTime !== null) {
    return { kind: "dateTime", serial: dateTime };
  }

  const isoDate = value.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (isoDate) {
    const serial = civilDateToExcelSerial(Number(isoDate[1]), Number(isoDate[2]), Number(isoDate[3]));
    return serial === null ? null : { kind: "date", serial };
  }

  const usDate = value.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2}|\d{4})$/);
  if (!usDate) {
    return null;
  }

  const inputYear = Number(usDate[3]);
  const year = usDate[3].length === 2
    ? inputYear < TWO_DIGIT_YEAR_PIVOT
      ? 2000 + inputYear
      : 1900 + inputYear
    : inputYear;
  const serial = civilDateToExcelSerial(year, Number(usDate[1]), Number(usDate[2]));
  return serial === null ? null : { kind: "date", serial };
}

function parseIsoDateTime(value: string): number | null {
  const match = value.match(
    /^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2})(?::(\d{2})(?:\.(\d{1,9}))?)?(Z|[+-]\d{2}:\d{2})?$/i
  );
  if (!match) {
    return null;
  }

  const year = Number(match[1]);
  const month = Number(match[2]);
  const day = Number(match[3]);
  const hour = Number(match[4]);
  const minute = Number(match[5]);
  const second = match[6] === undefined ? 0 : Number(match[6]);
  const millisecond = match[7] === undefined ? 0 : Math.round(Number(`0.${match[7]}`) * 1000);
  const dateSerial = civilDateToExcelSerial(year, month, day);
  if (dateSerial === null || hour > 23 || minute > 59 || second > 59) {
    return null;
  }

  const offsetMilliseconds = parseTimezoneOffsetMilliseconds(match[8]);
  if (offsetMilliseconds === null) {
    return null;
  }

  const timeMilliseconds = ((hour * 60 + minute) * 60 + second) * 1000 + millisecond - offsetMilliseconds;
  return Math.round((dateSerial * MILLISECONDS_PER_DAY + timeMilliseconds)) / MILLISECONDS_PER_DAY;
}

function parseTimezoneOffsetMilliseconds(value: string | undefined): number | null {
  if (value === undefined || value.toUpperCase() === "Z") {
    return 0;
  }

  const match = value.match(/^([+-])(\d{2}):(\d{2})$/);
  if (!match) {
    return null;
  }
  const hours = Number(match[2]);
  const minutes = Number(match[3]);
  if (hours > 23 || minutes > 59) {
    return null;
  }

  const offset = (hours * 60 + minutes) * 60 * 1000;
  return match[1] === "+" ? offset : -offset;
}

function civilDateToExcelSerial(year: number, month: number, day: number): number | null {
  if (year === 1900 && month === 2 && day === 29) {
    return EXCEL_LEAP_DAY_SERIAL;
  }
  if (year < 1900 || year > 9999 || month < 1 || month > 12 || day < 1 || day > 31) {
    return null;
  }

  const timestamp = Date.UTC(year, month - 1, day);
  const date = new Date(timestamp);
  if (date.getUTCFullYear() !== year || date.getUTCMonth() !== month - 1 || date.getUTCDate() !== day) {
    return null;
  }
  return dateToExcelSerial(date);
}
