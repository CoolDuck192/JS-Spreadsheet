import fc from "fast-check";
import { describe, expect, it } from "vitest";
import type {
  ConditionalFormatRule,
  DataValidationRule,
  SheetFilter,
  StructuredTable,
  TableAggregate,
  WorkbookModel
} from "../../types";
import { createBlankWorkbook } from "../../lib/workbook";
import type { FilterExpression, TableSort } from "../../table/core/query";
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
  workbook?: () => WorkbookModel;
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
  },
  {
    name: "sheet filter outside the sheet",
    command: sheetFilterCommand({
      range: { start: { row: 99, column: 0 }, end: { row: 100, column: 0 } }
    })
  },
  {
    name: "sheet filter with a column outside its range",
    command: sheetFilterCommand({ column: 2 })
  },
  {
    name: "sheet filter with a blank id",
    command: sheetFilterCommand({ id: "   " })
  },
  {
    name: "sheet filter with an invalid operator",
    command: sheetFilterCommand({ operator: "startsWith" })
  },
  {
    name: "sheet filter with a non-string value",
    command: sheetFilterCommand({ value: 5 })
  },
  {
    name: "sheet filter with non-string values",
    command: sheetFilterCommand({ values: ["Open", 5] })
  },
  {
    name: "sheet filter with a non-boolean header flag",
    command: sheetFilterCommand({ hasHeader: "yes" })
  },
  {
    name: "rich clipboard validation with reversed bounds",
    command: clipboardCommand({
      mode: "all",
      validation: { type: "number", min: 10, max: 5, allowBlank: true }
    })
  },
  {
    name: "rich clipboard with a missing runtime mode",
    command: clipboardCommand({ omitMode: true })
  },
  {
    name: "rich clipboard with non-finite content",
    command: clipboardCommand({ mode: "all", content: Number.NaN })
  },
  {
    name: "rich value paste with non-finite display content",
    command: clipboardCommand({ mode: "values", displayContent: Number.POSITIVE_INFINITY })
  },
  {
    name: "replacement workbook with duplicate table ids",
    command: {
      type: "workbook.replace",
      workbook: invalidReplacementWorkbook(),
      history: "reset"
    }
  },
  {
    name: "structured table with an invalid totals aggregate",
    command: {
      type: "table.setTotalsFunction",
      tableId: "table-parity",
      columnId: "column-parity",
      aggregate: "median" as TableAggregate
    },
    workbook: () => structuredParityWorkbook(true)
  },
  {
    name: "structured table with an invalid sort direction",
    command: {
      type: "table.sort",
      tableId: "table-parity",
      sorting: [{ columnId: "column-parity", direction: "sideways" }] as unknown as readonly TableSort[]
    },
    workbook: () => structuredParityWorkbook()
  },
  {
    name: "structured table with an invalid filter operator",
    command: {
      type: "table.setFilter",
      tableId: "table-parity",
      filter: {
        kind: "comparison",
        columnId: "column-parity",
        operator: "approximately",
        value: { type: "string", value: "Ada" }
      } as unknown as FilterExpression
    },
    workbook: () => structuredParityWorkbook()
  },
  {
    name: "structured table with a null runtime filter",
    command: {
      type: "table.setFilter",
      tableId: "table-parity",
      filter: null
    } as unknown as WorkbookCommand,
    workbook: () => structuredParityWorkbook()
  }
];

describe("workbook command migration parity", () => {
  it("enumerates every workbook command type for parity review", () => {
    expect(Object.keys(commandTypeCoverage)).toHaveLength(71);
  });

  it("returns null instead of throwing for a persisted null table filter", () => {
    const workbook = structuredParityWorkbook();
    const serialized = JSON.parse(JSON.stringify({
      ...workbook,
      tables: [{ ...workbook.tables[0], filter: null }]
    })) as unknown;

    expect(() => migrateWorkbookModel(serialized)).not.toThrow();
    expect(migrateWorkbookModel(serialized)).toBeNull();
  });

  it.each(invalidCommands)("rejects $name without changing the workbook", ({ command, workbook }) => {
    const session = createWorkbookSession({ workbook: workbook?.() ?? createBlankWorkbook() });
    const before = session.getSnapshot();

    expect(session.dispatch(command)).toMatchObject({
      status: "rejected",
      reason: "validation"
    });
    expect(session.getSnapshot()).toBe(before);
    session.destroy();
  });

  it("round-trips every committed validation, conditional-format, filter, named-range, and chart command", () => {
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
            type: "sheet.filter.set",
            sheetId: "sheet-1",
            filter: {
              id: "generated-sheet-filter",
              range: validationRange,
              column: validationRange.start.column,
              operator: "equals",
              value: "Open",
              values: ["Open", "Closed"],
              hasHeader: true
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

function sheetFilterCommand(overrides: Readonly<{
  id?: unknown;
  range?: SheetFilter["range"];
  column?: unknown;
  operator?: unknown;
  value?: unknown;
  values?: unknown;
  hasHeader?: unknown;
}>): WorkbookCommand {
  return {
    type: "sheet.filter.set",
    sheetId: "sheet-1",
    filter: {
      id: overrides.id ?? "sheet-filter",
      range: overrides.range ?? inBoundsRange,
      column: overrides.column ?? 0,
      operator: overrides.operator ?? "equals",
      value: overrides.value ?? "Open",
      ...(overrides.values === undefined ? {} : { values: overrides.values }),
      ...(overrides.hasHeader === undefined ? {} : { hasHeader: overrides.hasHeader })
    } as unknown as SheetFilter
  };
}

function clipboardCommand(overrides: Readonly<{
  mode?: unknown;
  omitMode?: boolean;
  content?: unknown;
  displayContent?: unknown;
  validation?: unknown;
}>): WorkbookCommand {
  const command = {
    type: "clipboard.paste",
    sheetId: "sheet-1",
    target: { row: 0, column: 0 },
    payload: {
      range: { start: { row: 0, column: 0 }, end: { row: 0, column: 0 } },
      cells: [[{
        sourceAddress: "A1",
        content: overrides.content ?? null,
        displayContent: overrides.displayContent ?? null,
        format: {},
        validation: overrides.validation ?? null,
        comment: null,
        hyperlink: null
      }]]
    }
  };
  return (overrides.omitMode
    ? command
    : { ...command, mode: overrides.mode ?? "all" }) as unknown as WorkbookCommand;
}

function structuredParityWorkbook(totalsRow = false): WorkbookModel {
  const workbook = createBlankWorkbook();
  const table: StructuredTable = {
    id: "table-parity",
    name: "ParityTable",
    sheetId: workbook.activeSheetId,
    range: {
      start: { row: 0, column: 0 },
      end: { row: totalsRow ? 2 : 1, column: 0 }
    },
    headerRow: true,
    totalsRow,
    columns: [{ id: "column-parity", name: "Name", sheetColumn: 0 }],
    rowIds: ["row-parity"]
  };
  return {
    ...workbook,
    tables: [table],
    sheets: [{
      ...workbook.sheets[0],
      cells: { A1: "Name", A2: "Ada" }
    }]
  };
}

function invalidReplacementWorkbook(): WorkbookModel {
  const workbook = structuredParityWorkbook();
  return { ...workbook, tables: [workbook.tables[0], { ...workbook.tables[0] }] };
}

const commandTypeCoverage = {
  "transaction": "workbook",
  "selection.set": "session",
  "cell.set": "workbook",
  "cell.comment.set": "workbook",
  "cell.hyperlink.set": "workbook",
  "range.clear": "workbook",
  "range.format": "workbook",
  "range.directFormat.clear": "workbook",
  "range.format.replace": "workbook",
  "range.borders": "workbook",
  "range.validation.set": "workbook",
  "range.validation.clear": "workbook",
  "range.conditionalFormat.add": "workbook",
  "range.conditionalFormat.remove": "workbook",
  "range.conditionalFormat.clear": "workbook",
  "range.readOnly.set": "workbook",
  "range.merge": "workbook",
  "range.unmerge": "workbook",
  "range.fill": "workbook",
  "range.autoFill": "workbook",
  "range.sort": "workbook",
  "range.removeDuplicates": "workbook",
  "clipboard.paste": "workbook",
  "clipboard.move": "workbook",
  "clipboard.pasteMatrix": "workbook",
  "rows.insert": "workbook",
  "rows.delete": "workbook",
  "columns.insert": "workbook",
  "columns.delete": "workbook",
  "rows.resize": "workbook",
  "columns.resize": "workbook",
  "rows.hidden.set": "workbook",
  "columns.hidden.set": "workbook",
  "sheet.add": "workbook",
  "sheet.createFromMatrix": "workbook",
  "sheet.replaceWithRows": "workbook",
  "sheet.rename": "workbook",
  "sheet.duplicate": "workbook",
  "sheet.delete": "workbook",
  "sheet.activate": "workbook",
  "sheet.move": "workbook",
  "sheet.hidden.set": "workbook",
  "sheet.tabColor.set": "workbook",
  "sheet.freeze.set": "workbook",
  "sheet.protection.set": "workbook",
  "sheet.filter.set": "workbook",
  "sheet.filter.clear": "workbook",
  "sheet.chart.add": "workbook",
  "sheet.chart.delete": "workbook",
  "namedRange.define": "workbook",
  "namedRange.remove": "workbook",
  "history.undo": "workbook",
  "history.redo": "workbook",
  "persistence.status": "session",
  "workbook.replace": "workbook",
  "table.create": "workbook",
  "table.rename": "workbook",
  "table.renameColumn": "workbook",
  "table.resize": "workbook",
  "table.setHeaderRow": "workbook",
  "table.setTotalsRow": "workbook",
  "table.setTotalsFunction": "workbook",
  "table.setStyle": "workbook",
  "table.setKeyColumn": "workbook",
  "table.setCalculatedColumn": "workbook",
  "table.setFilter": "workbook",
  "table.sort": "workbook",
  "table.insertRows": "workbook",
  "table.deleteRows": "workbook",
  "table.editCells": "workbook",
  "table.convertToRange": "workbook"
} as const satisfies Record<WorkbookCommand["type"], "session" | "workbook">;
