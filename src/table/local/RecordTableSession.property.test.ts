import fc from "fast-check";
import { expect, it } from "vitest";
import { createColumnHelper } from "../core/columnHelper";
import { createLocalRecordTableSession } from "./RecordTableSession";

type Row = { id: string; value: number };

it("restores the exact original dataset after random edits are undone", async () => {
  await fc.assert(fc.asyncProperty(
    fc.array(fc.record({ rowIndex: fc.integer({ min: 0, max: 4 }), value: fc.integer() }), { maxLength: 100 }),
    async (operations) => {
      const original: readonly Row[] = Array.from({ length: 5 }, (_, index) => ({ id: `row-${index}`, value: index }));
      const helper = createColumnHelper<Row>();
      const session = createLocalRecordTableSession({
        source: { kind: "local", rows: original, getRowId: (row) => row.id },
        columns: [helper.accessor("value", { id: "value", header: "Value", dataType: "number" })],
        historyLimit: 100
      });

      for (const operation of operations) {
        await session.dispatch({
          type: "edit-cells",
          edits: [{ rowId: `row-${operation.rowIndex}`, columnId: "value", rawText: String(operation.value) }]
        });
      }
      for (let index = 0; index < Math.min(operations.length, 100); index += 1) {
        await session.undo();
      }

      const restored = original.map((row) => ({
        id: row.id,
        value: session.getSnapshot().getCell(row.id, "value").storedValue
      }));
      expect(restored).toEqual(original);
      session.destroy();
    }
  ), { seed: 20260709, numRuns: 200 });
});
