import { describe, expect, it } from "vitest";
import type { DataValidationRule } from "../types";
import { summarizeDataValidationRules } from "./dataValidationSummary";

describe("dataValidationSummary", () => {
  it("groups adjacent cells with the same validation into rectangular ranges", () => {
    const listRule: DataValidationRule = { type: "list", values: ["Open", "Closed"] };
    const numberRule: DataValidationRule = { type: "number", min: 1, max: 10 };
    const textLengthRule = { type: "textLength", min: 2, max: 5 } as DataValidationRule;

    const summaries = summarizeDataValidationRules({
      A1: listRule,
      A2: listRule,
      B1: numberRule,
      B2: numberRule,
      C3: { type: "list", values: ["High", "Low"] },
      D1: textLengthRule,
      D2: textLengthRule
    });

    expect(summaries).toMatchObject([
      {
        id: "A1:A2|list|Open,Closed",
        range: { start: { row: 0, column: 0 }, end: { row: 1, column: 0 } },
        rule: listRule
      },
      {
        id: "B1:B2|number|1|10",
        range: { start: { row: 0, column: 1 }, end: { row: 1, column: 1 } },
        rule: numberRule
      },
      {
        id: "D1:D2|textLength|2|5",
        range: { start: { row: 0, column: 3 }, end: { row: 1, column: 3 } },
        rule: textLengthRule
      },
      {
        id: "C3|list|High,Low",
        range: { start: { row: 2, column: 2 }, end: { row: 2, column: 2 } }
      }
    ]);
  });
});
