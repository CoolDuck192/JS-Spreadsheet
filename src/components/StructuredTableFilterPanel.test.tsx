import { render, screen } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import { describe, expect, it, vi } from "vitest";
import type { FilterExpression } from "../table/core/query";
import type { StructuredTableColumn } from "../types";
import { StructuredTableFilterPanel } from "./StructuredTableFilterPanel";

const columns: StructuredTableColumn[] = [
  { id: "column-region", name: "Region", sheetColumn: 0 },
  { id: "column-sales", name: "Sales", sheetColumn: 1 }
];

describe("StructuredTableFilterPanel", () => {
  it.each([
    {
      kind: "comparison",
      configure: async (user: ReturnType<typeof userEvent.setup>) => {
        await user.selectOptions(screen.getByLabelText("Filter column"), "column-sales");
        await user.selectOptions(screen.getByLabelText("Value type"), "number");
        await user.selectOptions(screen.getByLabelText("Comparison operator"), "gte");
        await user.type(screen.getByLabelText("Filter value"), "10");
      },
      expected: {
        kind: "comparison",
        columnId: "column-sales",
        operator: "gte",
        value: { type: "number", value: 10 }
      }
    },
    {
      kind: "set",
      configure: async (user: ReturnType<typeof userEvent.setup>) => {
        await user.selectOptions(screen.getByLabelText("Filter type"), "set");
        await user.selectOptions(screen.getByLabelText("Set operator"), "notIn");
        await user.type(screen.getByLabelText("Set values"), "West, East");
      },
      expected: {
        kind: "set",
        columnId: "column-region",
        operator: "notIn",
        values: [
          { type: "string", value: "West" },
          { type: "string", value: "East" }
        ]
      }
    },
    {
      kind: "range",
      configure: async (user: ReturnType<typeof userEvent.setup>) => {
        await user.selectOptions(screen.getByLabelText("Filter type"), "range");
        await user.selectOptions(screen.getByLabelText("Filter column"), "column-sales");
        await user.selectOptions(screen.getByLabelText("Value type"), "number");
        await user.selectOptions(screen.getByLabelText("Range operator"), "notBetween");
        await user.type(screen.getByLabelText("Lower value"), "5");
        await user.type(screen.getByLabelText("Upper value"), "20");
      },
      expected: {
        kind: "range",
        columnId: "column-sales",
        operator: "notBetween",
        lower: { type: "number", value: 5 },
        upper: { type: "number", value: 20 }
      }
    },
    {
      kind: "blank",
      configure: async (user: ReturnType<typeof userEvent.setup>) => {
        await user.selectOptions(screen.getByLabelText("Filter type"), "blank");
        await user.selectOptions(screen.getByLabelText("Blank operator"), "isNotBlank");
      },
      expected: {
        kind: "blank",
        columnId: "column-region",
        operator: "isNotBlank"
      }
    },
    {
      kind: "logical",
      configure: async (user: ReturnType<typeof userEvent.setup>) => {
        await user.selectOptions(screen.getByLabelText("Filter type"), "logical");
        await user.selectOptions(screen.getByLabelText("Logical operator"), "or");
        await user.type(screen.getByLabelText("First comparison value"), "West");
        await user.type(screen.getByLabelText("Second comparison value"), "East");
      },
      expected: {
        kind: "logical",
        operator: "or",
        operands: [
          {
            kind: "comparison",
            columnId: "column-region",
            operator: "eq",
            value: { type: "string", value: "West" }
          },
          {
            kind: "comparison",
            columnId: "column-region",
            operator: "eq",
            value: { type: "string", value: "East" }
          }
        ]
      }
    },
    {
      kind: "not",
      configure: async (user: ReturnType<typeof userEvent.setup>) => {
        await user.selectOptions(screen.getByLabelText("Filter type"), "not");
        await user.type(screen.getByLabelText("Comparison value"), "Closed");
      },
      expected: {
        kind: "not",
        operand: {
          kind: "comparison",
          columnId: "column-region",
          operator: "eq",
          value: { type: "string", value: "Closed" }
        }
      }
    }
  ] as Array<{
    kind: FilterExpression["kind"];
    configure(user: ReturnType<typeof userEvent.setup>): Promise<void>;
    expected: FilterExpression;
  }>)("builds an exact $kind filter AST", async ({ configure, expected }) => {
    const user = userEvent.setup();
    const onApply = vi.fn();
    render(<StructuredTableFilterPanel columns={columns} activeColumnId="column-region" onApply={onApply} />);

    await configure(user);
    await user.click(screen.getByRole("button", { name: "Apply table filter" }));

    expect(onApply).toHaveBeenCalledOnce();
    expect(onApply).toHaveBeenCalledWith(expected);
  });

  it("clears the table filter without inventing a hidden-row mutation", async () => {
    const user = userEvent.setup();
    const onApply = vi.fn();
    render(<StructuredTableFilterPanel columns={columns} activeColumnId="column-region" onApply={onApply} />);

    await user.click(screen.getByRole("button", { name: "Clear table filter" }));

    expect(onApply).toHaveBeenCalledWith(undefined);
  });
});
