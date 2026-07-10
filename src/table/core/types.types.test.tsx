import { expectTypeOf, test } from "vitest";
import { createColumnHelper, type ColumnDef, type ExportArtifact, type TableIntent, type TableSession } from "./index";

type Employee = { id: string; name: string; salary: number };

test("column helper preserves accessor value types", () => {
  const salary = createColumnHelper<Employee>().accessor("salary", {
    id: "salary",
    header: "Salary"
  });
  expectTypeOf(salary.accessor({ id: "1", name: "Ada", salary: 42 })).toEqualTypeOf<number>();
  expectTypeOf<TableSession<Employee>["getSnapshot"]>().toBeFunction();
  expectTypeOf<Awaited<ReturnType<TableSession<Employee>["export"]>>>().toEqualTypeOf<ExportArtifact>();
  expectTypeOf<ColumnDef<Employee>>().toBeObject();

  const reload: TableIntent<Employee> = {
    type: "reload-authoritative",
    operationId: "operation-1",
    rowId: "employee-1"
  };
  const retry: TableIntent<Employee> = {
    type: "retry-with-revision",
    operationId: "operation-1",
    rowId: "employee-1",
    expectedRevision: "revision-2"
  };
  expectTypeOf(reload.operationId).toEqualTypeOf<string>();
  expectTypeOf(retry.expectedRevision).toEqualTypeOf<string>();

  // @ts-expect-error retry requires the authoritative revision being accepted
  const missingRevision: TableIntent<Employee> = {
    type: "retry-with-revision",
    operationId: "operation-1",
    rowId: "employee-1"
  };
  expectTypeOf(missingRevision).toMatchTypeOf<TableIntent<Employee>>();
});
