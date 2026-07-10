import type { ComponentProps } from "react";
import { expectTypeOf, test } from "vitest";
import type { ExportArtifact, TableSession } from "../table/core/types";
import type { ColumnDef, DataTableHandle, DataTableProps } from "./tableTypes";

type Employee = { id: string; name: string; salary: number };

test("public local props infer row and renderer values", () => {
  const columns: readonly ColumnDef<Employee>[] = [{
    id: "name",
    header: "Name",
    accessor: (row) => row.name,
    update: (row, value) => ({ ...row, name: String(value) }),
    cell: ({ row }) => row.original.name
  }];
  const props: DataTableProps<Employee> = {
    rows: [{ id: "1", name: "Ada", salary: 100 }],
    columns,
    getRowId: (row) => row.id,
    onRowsChange: (updater) => updater([{ id: "1", name: "Ada", salary: 100 }])
  };

  expectTypeOf(props.getRowId).toBeFunction();
  expectTypeOf<DataTableHandle["scrollToRow"]>().toBeFunction();
  expectTypeOf<Awaited<ReturnType<TableSession<Employee>["export"]>>>().toEqualTypeOf<ExportArtifact>();
  expectTypeOf<Awaited<ReturnType<DataTableHandle["export"]>>>().toEqualTypeOf<Blob>();
  expectTypeOf<ComponentProps<"div">["className"]>().toEqualTypeOf<string | undefined>();
});

test("simple and session modes cannot be mixed", () => {
  const columns: readonly ColumnDef<Employee>[] = [];
  // @ts-expect-error session mode may not also provide rows
  const invalid: DataTableProps<Employee> = {
    session: {} as TableSession<Employee, ColumnDef<Employee>>,
    rows: [],
    columns,
    getRowId: (row) => row.id
  };
  expectTypeOf(invalid).toMatchTypeOf<DataTableProps<Employee>>();
});
