import { describe, expect, it } from "vitest";
import { createRandomId, createStableKeyRowId } from "./ids";

describe("workbook IDs", () => {
  it("creates kind-prefixed non-empty random IDs", () => {
    expect(createRandomId("table")).toMatch(/^table-[0-9a-f-]+$/);
    expect(createRandomId("table-column")).toMatch(/^table-column-[0-9a-f-]+$/);
    expect(createRandomId("table-row")).toMatch(/^table-row-[0-9a-f-]+$/);
  });

  it("derives stable typed row IDs from XLSX keys", async () => {
    const first = await createStableKeyRowId("Employees", "Employee ID", 101);
    const repeat = await createStableKeyRowId("employees", "employee id", 101);
    const stringKey = await createStableKeyRowId("Employees", "Employee ID", "101");

    expect(first).toBe(repeat);
    expect(first).not.toBe(stringKey);
    expect(first).toMatch(/^table-row-key-[0-9a-f]{32}$/);
  });

  it("normalizes Unicode deterministically without locale-sensitive casing", async () => {
    await expect(createStableKeyRowId("Cafe\u0301", "KEY", true)).resolves.toBe(
      await createStableKeyRowId("Café", "key", true)
    );
    await expect(createStableKeyRowId("I", "Key", null)).resolves.not.toBe(
      await createStableKeyRowId("İ", "Key", null)
    );
  });
});
