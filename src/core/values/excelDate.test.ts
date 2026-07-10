import { describe, expect, it } from "vitest";
import { dateToExcelSerial, excelSerialToDate, parseExcelTemporalInput } from "./excelDate";

describe("Excel 1900 date conversion", () => {
  it("uses Excel's leap-day-compatible serials before and after March 1900", () => {
    expect(dateToExcelSerial(new Date(Date.UTC(1899, 11, 31)))).toBe(0);
    expect(dateToExcelSerial(new Date(Date.UTC(1900, 0, 15)))).toBe(15);
    expect(dateToExcelSerial(new Date(Date.UTC(1900, 1, 28)))).toBe(59);
    expect(dateToExcelSerial(new Date(Date.UTC(1900, 2, 1)))).toBe(61);
    expect(dateToExcelSerial(new Date(Date.UTC(2026, 0, 15)))).toBe(46037);

    expect(excelSerialToDate(15)?.toISOString()).toBe("1900-01-15T00:00:00.000Z");
    expect(excelSerialToDate(61)?.toISOString()).toBe("1900-03-01T00:00:00.000Z");
    expect(excelSerialToDate(46037.5)?.toISOString()).toBe("2026-01-15T12:00:00.000Z");
  });

  it("parses the configured ISO and US civil date forms deterministically", () => {
    expect(parseExcelTemporalInput("2026-01-15")).toEqual({ kind: "date", serial: 46037 });
    expect(parseExcelTemporalInput("01/15/2026")).toEqual({ kind: "date", serial: 46037 });
    expect(parseExcelTemporalInput("01/15/26")).toEqual({ kind: "date", serial: 46037 });
    expect(parseExcelTemporalInput("01/15/29")).toEqual({ kind: "date", serial: 47133 });
    expect(parseExcelTemporalInput("01/15/30")).toEqual({ kind: "date", serial: 10973 });
    expect(parseExcelTemporalInput("1900-02-29")).toEqual({ kind: "date", serial: 60 });
  });

  it("parses ISO date-times without consulting the host locale or time zone", () => {
    expect(parseExcelTemporalInput("2026-01-15T12:00:00Z")).toEqual({ kind: "dateTime", serial: 46037.5 });
    expect(parseExcelTemporalInput("2026-01-15T12:00:00")).toEqual({ kind: "dateTime", serial: 46037.5 });
    expect(parseExcelTemporalInput("2026-01-15T07:00:00-05:00")).toEqual({ kind: "dateTime", serial: 46037.5 });
  });

  it("applies timezone offsets before crossing Excel's leap-day discontinuity", () => {
    const beforeDiscontinuity = parseExcelTemporalInput("1900-03-01T00:30:00+01:00");
    const afterDiscontinuity = parseExcelTemporalInput("1900-02-28T23:30:00-01:00");

    expect(beforeDiscontinuity?.kind).toBe("dateTime");
    expect(beforeDiscontinuity?.serial).toBeCloseTo(59 + 23.5 / 24, 10);
    expect(afterDiscontinuity?.kind).toBe("dateTime");
    expect(afterDiscontinuity?.serial).toBeCloseTo(61 + 0.5 / 24, 10);
  });

  it("rejects invalid dates and locale-dependent date text", () => {
    for (const raw of ["2026-02-29", "13/01/2026", "15/01/2026", "January 15, 2026", "2026/01/15"]) {
      expect(parseExcelTemporalInput(raw)).toBeNull();
    }
  });
});
