import { describe, expect, it } from "vitest";
import { createColumnHelper, normalizeColumns } from "./columnHelper";

type Employee = { id: string; name: string; salary: number };

describe("column definitions", () => {
  it("creates immutable accessor updates with inferred values", () => {
    const column = createColumnHelper<Employee>().accessor("salary", {
      id: "salary",
      header: "Salary",
      dataType: "number"
    });
    const original: Employee = { id: "e1", name: "Ada", salary: 100 };
    const updated = column.update!(original, 125);
    expect(updated).toEqual({ id: "e1", name: "Ada", salary: 125 });
    expect(original.salary).toBe(100);
  });

  it("rejects blank and duplicate stable column ids", () => {
    const helper = createColumnHelper<Employee>();
    expect(() => normalizeColumns([
      helper.accessor("name", { id: "name", header: "Name" }),
      helper.accessor("salary", { id: "name", header: "Salary" })
    ])).toThrow("Duplicate column id: name");
    expect(() => normalizeColumns([
      helper.display({ id: " ", header: "Actions" })
    ])).toThrow("Column id must not be blank");
  });
});
