import { render, screen } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import { describe, expect, it, vi } from "vitest";
import type { StructuredTableColumn } from "../types";
import { CalculatedColumnPanel } from "./CalculatedColumnPanel";

const columns: StructuredTableColumn[] = [
  { id: "column-quantity", name: "Quantity", sheetColumn: 0 },
  { id: "column-total", name: "Total", sheetColumn: 1 }
];

describe("CalculatedColumnPanel", () => {
  it("loads the selected column's committed formula when the dropdown changes", async () => {
    const user = userEvent.setup();
    const formulaColumns: StructuredTableColumn[] = [
      { ...columns[0], calculatedFormula: "=[@Quantity]+1" },
      { ...columns[1], calculatedFormula: "=[@Quantity]*2" }
    ];

    render(
      <CalculatedColumnPanel
        columns={formulaColumns}
        activeColumnId="column-total"
        onApply={vi.fn()}
      />
    );

    expect(screen.getByLabelText("Calculated column formula")).toHaveValue("=[@Quantity]*2");
    await user.selectOptions(screen.getByRole("combobox", { name: "Calculated column" }), "column-quantity");
    expect(screen.getByLabelText("Calculated column formula")).toHaveValue("=[@Quantity]+1");
  });

  it("rejects a calculated-column formula without a leading equals sign", async () => {
    const user = userEvent.setup();
    const onApply = vi.fn();
    render(<CalculatedColumnPanel columns={columns} activeColumnId="column-total" onApply={onApply} />);

    const formula = screen.getByLabelText("Calculated column formula");
    await user.type(formula, "Quantity*2");
    await user.click(screen.getByRole("button", { name: "Apply calculated column" }));

    expect(onApply).not.toHaveBeenCalled();
    expect(screen.getByRole("alert")).toHaveTextContent("Formula must start with =");
    expect(formula).toHaveFocus();
  });

  it("dispatches the formula for the active stable column id", async () => {
    const user = userEvent.setup();
    const onApply = vi.fn();
    render(<CalculatedColumnPanel columns={columns} activeColumnId="column-total" onApply={onApply} />);

    await user.type(screen.getByLabelText("Calculated column formula"), "=Quantity*2");
    await user.click(screen.getByRole("button", { name: "Apply calculated column" }));

    expect(onApply).toHaveBeenCalledOnce();
    expect(onApply).toHaveBeenCalledWith("column-total", "=Quantity*2");
  });
});
