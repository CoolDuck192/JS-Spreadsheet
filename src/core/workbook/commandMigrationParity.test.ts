import fc from "fast-check";
import { describe, expect, it } from "vitest";
import type { ConditionalFormatRule, DataValidationRule } from "../../types";
import { createBlankWorkbook } from "../../lib/workbook";
import type { WorkbookCommand } from "./commands";
import { migrateWorkbookModel } from "./migrateWorkbook";
import { createWorkbookSession } from "./WorkbookSession";

const inBoundsRange = {
  start: { row: 0, column: 0 },
  end: { row: 1, column: 1 }
} as const;

const invalidCommands: readonly Readonly<{
  name: string;
  command: WorkbookCommand;
}>[] = [
  {
    name: "validation with reversed bounds",
    command: {
      type: "range.validation.set",
      sheetId: "sheet-1",
      range: inBoundsRange,
      rule: { type: "number", min: 10, max: 5 }
    }
  },
  {
    name: "validation with a non-finite bound",
    command: {
      type: "range.validation.set",
      sheetId: "sheet-1",
      range: inBoundsRange,
      rule: { type: "textLength", min: Number.POSITIVE_INFINITY }
    }
  },
  {
    name: "validation with an invalid runtime shape",
    command: {
      type: "range.validation.set",
      sheetId: "sheet-1",
      range: inBoundsRange,
      rule: { type: "list", values: ["valid", 1], allowBlank: "yes" } as unknown as DataValidationRule
    }
  },
  {
    name: "validation outside the sheet",
    command: {
      type: "range.validation.set",
      sheetId: "sheet-1",
      range: { start: { row: 99, column: 0 }, end: { row: 100, column: 0 } },
      rule: { type: "number" }
    }
  },
  {
    name: "named range outside the sheet",
    command: {
      type: "namedRange.define",
      namedRange: {
        name: "OutOfBounds",
        sheetId: "sheet-1",
        range: { start: { row: 99, column: 25 }, end: { row: 100, column: 25 } }
      }
    }
  },
  {
    name: "chart data range outside the sheet",
    command: {
      type: "sheet.chart.add",
      sheetId: "sheet-1",
      chart: {
        id: "chart-range",
        title: "Range",
        type: "bar",
        range: { start: { row: 0, column: 0 }, end: { row: 100, column: 1 } },
        anchor: { row: 0, column: 3 }
      }
    }
  },
  {
    name: "chart anchor outside the sheet",
    command: {
      type: "sheet.chart.add",
      sheetId: "sheet-1",
      chart: {
        id: "chart-anchor",
        title: "Anchor",
        type: "line",
        range: inBoundsRange,
        anchor: { row: 0, column: 26 }
      }
    }
  },
  {
    name: "chart with a blank id",
    command: {
      type: "sheet.chart.add",
      sheetId: "sheet-1",
      chart: {
        id: "   ",
        title: "Blank id",
        type: "pie",
        range: inBoundsRange,
        anchor: { row: 0, column: 3 }
      }
    }
  },
  {
    name: "conditional format outside the sheet",
    command: conditionalFormatCommand({
      range: { start: { row: 99, column: 0 }, end: { row: 100, column: 0 } }
    })
  },
  {
    name: "conditional format with a blank id",
    command: conditionalFormatCommand({ id: "   " })
  },
  {
    name: "conditional format with a zero top count",
    command: conditionalFormatCommand({ condition: { type: "top", count: 0 } })
  },
  {
    name: "conditional format with a non-positive font size",
    command: conditionalFormatCommand({ format: { fontSize: 0 } })
  },
  {
    name: "conditional format with an unknown number format",
    command: conditionalFormatCommand({ format: { numberFormat: "accounting" } })
  },
  {
    name: "conditional format with a non-thin border",
    command: conditionalFormatCommand({
      format: { borders: { top: { style: "double", color: "#123456" } } }
    })
  }
];

describe("workbook command migration parity", () => {
  it.each(invalidCommands)("rejects $name without changing the workbook", ({ command }) => {
    const session = createWorkbookSession({ workbook: createBlankWorkbook() });
    const before = session.getSnapshot();

    expect(session.dispatch(command)).toMatchObject({
      status: "rejected",
      reason: "validation"
    });
    expect(session.getSnapshot()).toBe(before);
    session.destroy();
  });

  it("round-trips every committed validation, conditional-format, named-range, and chart command", () => {
    const coordinate = fc.record({
      row: fc.integer({ min: 0, max: 102 }),
      column: fc.integer({ min: 0, max: 28 })
    });
    const range = fc.record({ start: coordinate, end: coordinate });
    const validationRule = fc.oneof(
      fc.record({
        type: fc.constant("list" as const),
        values: fc.array(fc.string(), { maxLength: 5 }),
        allowBlank: fc.option(fc.boolean(), { nil: undefined })
      }),
      fc.record({
        type: fc.constantFrom("number" as const, "textLength" as const),
        min: fc.option(fc.oneof(fc.integer(), fc.constant(Number.NaN), fc.constant(Number.NEGATIVE_INFINITY)), { nil: undefined }),
        max: fc.option(fc.oneof(fc.integer(), fc.constant(Number.POSITIVE_INFINITY)), { nil: undefined }),
        allowBlank: fc.option(fc.boolean(), { nil: undefined })
      })
    );

    fc.assert(fc.property(range, validationRule, range, range, coordinate,
      (validationRange, rule, namedRange, chartRange, anchor) => {
        const commands: readonly WorkbookCommand[] = [
          {
            type: "range.validation.set",
            sheetId: "sheet-1",
            range: validationRange,
            rule
          },
          {
            type: "namedRange.define",
            namedRange: { name: "GeneratedRange", sheetId: "sheet-1", range: namedRange }
          },
          {
            type: "range.conditionalFormat.add",
            sheetId: "sheet-1",
            range: validationRange,
            rule: {
              id: "generated-conditional-format",
              range: inBoundsRange,
              condition: { type: "between", value: "1", secondValue: "5" },
              format: {
                bold: true,
                fontSize: 11,
                numberFormat: "number",
                borders: { bottom: { style: "thin", color: "#123456" } }
              }
            }
          },
          {
            type: "sheet.chart.add",
            sheetId: "sheet-1",
            chart: {
              id: "generated-chart",
              title: "Generated chart",
              type: "bar",
              range: chartRange,
              anchor
            }
          }
        ];

        for (const command of commands) {
          const session = createWorkbookSession({ workbook: createBlankWorkbook() });
          const result = session.dispatch(command);
          if (result.status === "committed") {
            const serialized = JSON.parse(JSON.stringify(session.getSnapshot().workbook)) as unknown;
            expect(migrateWorkbookModel(serialized)).not.toBeNull();
          }
          session.destroy();
        }
      }), { numRuns: 75 });
  });
});

function conditionalFormatCommand(overrides: Readonly<{
  id?: unknown;
  range?: ConditionalFormatRule["range"];
  condition?: unknown;
  format?: unknown;
}>): WorkbookCommand {
  return {
    type: "range.conditionalFormat.add",
    sheetId: "sheet-1",
    range: overrides.range ?? inBoundsRange,
    rule: {
      id: overrides.id ?? "conditional-format",
      range: inBoundsRange,
      condition: overrides.condition ?? { type: "blank" },
      format: overrides.format ?? { bold: true }
    } as unknown as ConditionalFormatRule
  };
}
