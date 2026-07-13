import { describe, expect, it } from "vitest";
import { createLocalRecordTableSession } from "./RecordTableSession";

const LARGE_ROW_COUNT = 100_000;

type LargeRow = {
  id: string;
  value: number;
};

describe("RecordTableSession at 100,000 rows", () => {
  it("builds one row projection for a query-state dispatch", async () => {
    const rows: LargeRow[] = Array.from({ length: LARGE_ROW_COUNT }, (_, index) => ({
      id: String(index + 1),
      value: index + 1
    }));
    let rowIdReads = 0;
    let accessorReads = 0;
    const session = createLocalRecordTableSession({
      source: {
        kind: "local",
        rows,
        getRowId: (row) => {
          rowIdReads += 1;
          return row.id;
        }
      },
      columns: [{
        id: "value",
        header: "Value",
        dataType: "number",
        sortable: true,
        accessor: (row) => {
          accessorReads += 1;
          return row.value;
        }
      }]
    });

    try {
      session.getSnapshot();
      rowIdReads = 0;
      accessorReads = 0;
      let publications = 0;
      session.subscribe(() => {
        publications += 1;
        session.getSnapshot();
      });

      const started = performance.now();
      const result = await session.dispatch({
        type: "set-sorting",
        sorting: [{ columnId: "value", direction: "desc" }]
      });
      const dispatchMs = performance.now() - started;

      expect(result).toMatchObject({ status: "committed", changed: true });
      expect(publications).toBe(1);
      expect(rowIdReads).toBe(LARGE_ROW_COUNT);
      expect(accessorReads).toBe(LARGE_ROW_COUNT);
      expect(session.getSnapshot().rows[0]).toMatchObject({ kind: "data", id: "100000" });

      // eslint-disable-next-line no-console
      console.log(`record-table query-state-dispatch=${Math.round(dispatchMs)}ms projections=1`);
      if (!process.env.SKIP_PERF_ASSERT) {
        expect(dispatchMs).toBeLessThan(10_000);
      }
    } finally {
      session.destroy();
    }
  }, 60_000);
});
