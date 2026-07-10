import { expect } from "vitest";
import { createLocalRecordTableSession } from "../local/RecordTableSession";
import { createColumnHelper } from "./columnHelper";
import { defineTableSessionContract } from "./session.contract";

type Employee = { id: string; name: string; department: string; salary: number };

defineTableSessionContract<Employee>("local records", () => {
  const helper = createColumnHelper<Employee>();
  const session = createLocalRecordTableSession({
    source: {
      kind: "local",
      rows: [{ id: "e1", name: "Ada", department: "Finance", salary: 100 }],
      getRowId: (row) => row.id
    },
    columns: [
      helper.accessor("name", { id: "name", header: "Name" }),
      helper.accessor("department", { id: "department", header: "Department" }),
      helper.accessor("salary", { id: "salary", header: "Salary", dataType: "number" })
    ]
  });
  return {
    session,
    editableCell: { rowId: "e1", columnId: "name" },
    validRawText: "Grace",
    featureOperations: {
      sort: () => session.dispatch({ type: "set-sorting", sorting: [{ columnId: "name", direction: "asc" }] }),
      filter: () => session.dispatch({
        type: "set-filter",
        filter: { kind: "comparison", columnId: "department", operator: "eq", value: { type: "string", value: "Finance" } }
      }),
      group: () => session.dispatch({ type: "set-grouping", grouping: [{ columnId: "department" }] }),
      aggregate: () => session.dispatch({
        type: "set-aggregates",
        aggregates: [{ id: "salary-sum", columnId: "salary", function: "sum" }]
      }),
      pagination: () => session.dispatch({ type: "set-pagination", pagination: { kind: "offset", offset: 0, limit: 25 } }),
      edit: () => session.dispatch({
        type: "edit-cells",
        edits: [{ rowId: "e1", columnId: "name", rawText: "Katherine" }]
      }),
      formula: () => session.dispatch({
        type: "edit-cells",
        edits: [{ rowId: "e1", columnId: "salary", rawText: "=[salary]*2" }]
      }),
      undo: () => session.undo(),
      export: async () => {
        const artifact = await session.export({ format: "csv", scope: "completeDataset" });
        expect(artifact).toMatchObject({
          bytes: expect.any(Uint8Array),
          mediaType: "text/csv;charset=utf-8"
        });
        expect(artifact.fileName).toMatch(/\.csv$/);
        return { status: "committed" as const, revision: session.getSnapshot().revision };
      }
    },
    cleanup: () => session.destroy()
  };
});
