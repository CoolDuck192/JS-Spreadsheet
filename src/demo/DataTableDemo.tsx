import { useMemo, useState } from "react";
import { DataTable } from "../react/DataTable";
import type { ColumnDef } from "../react/tableTypes";

type Employee = {
  id: string;
  name: string;
  department: "Engineering" | "Finance" | "Operations" | "People";
  role: string;
  salary: number;
  startDate: string;
  active: boolean;
  children?: readonly Employee[];
};

type TeamSummary = { id: string; team: string; headcount: number; location: string };

const initialEmployees: readonly Employee[] = [
  {
    id: "eng",
    name: "Engineering team",
    department: "Engineering",
    role: "Product engineering, platform reliability, and developer experience",
    salary: 0,
    startDate: "2021-01-04",
    active: true,
    children: [
      { id: "e-ada", name: "Ada Lovelace", department: "Engineering", role: "Staff product engineer", salary: 172000, startDate: "2022-03-14", active: true },
      { id: "e-grace", name: "Grace Hopper", department: "Engineering", role: "Principal platform engineer", salary: 188000, startDate: "2021-08-09", active: true }
    ]
  },
  { id: "e-katherine", name: "Katherine Johnson", department: "Finance", role: "Director, financial planning and analysis", salary: 164000, startDate: "2020-11-02", active: true },
  { id: "e-mary", name: "Mary Jackson", department: "Operations", role: "Senior operations program manager", salary: 138000, startDate: "2023-02-20", active: true },
  { id: "e-dorothy", name: "Dorothy Vaughan", department: "People", role: "People systems and workforce analytics lead", salary: 145000, startDate: "2022-09-12", active: false }
];

const summaryRows: readonly TeamSummary[] = [
  { id: "s-eng", team: "Engineering", headcount: 18, location: "Hybrid" },
  { id: "s-fin", team: "Finance", headcount: 7, location: "Toronto" },
  { id: "s-ops", team: "Operations", headcount: 11, location: "Remote" }
];

const summaryColumns: readonly ColumnDef<TeamSummary>[] = [
  { id: "team", header: "Team", accessor: (row) => row.team, sortable: true, filterable: true, width: 2 },
  { id: "headcount", header: "Headcount", dataType: "number", accessor: (row) => row.headcount, sortable: true, width: 1 },
  { id: "location", header: "Location", accessor: (row) => row.location, filterable: true, width: 1.4 }
];

const currency = new Intl.NumberFormat("en-CA", {
  style: "currency",
  currency: "CAD",
  maximumFractionDigits: 0
});

export function DataTableDemo() {
  const [employees, setEmployees] = useState<readonly Employee[]>(initialEmployees);
  const [selectedRowIds, setSelectedRowIds] = useState<readonly string[]>([]);
  const columns = useMemo<readonly ColumnDef<Employee>[]>(() => [
    {
      id: "name",
      header: "Employee",
      dataType: "text",
      accessor: (row) => row.name,
      update: (row, value) => ({ ...row, name: String(value) }),
      sortable: true,
      filterable: true,
      width: 1.6,
      headerActions: [{
        id: "normalize-names",
        label: "Normalize names",
        run: () => setEmployees((current) => current.map(normalizeEmployeeName))
      }]
    },
    {
      id: "department",
      header: "Department",
      dataType: "text",
      accessor: (row) => row.department,
      update: (row, value) => ({ ...row, department: String(value) as Employee["department"] }),
      sortable: true,
      filterable: true,
      groupable: true,
      width: 1.25
    },
    {
      id: "role",
      header: "Role",
      dataType: "text",
      accessor: (row) => row.role,
      update: (row, value) => ({ ...row, role: String(value) }),
      sortable: true,
      filterable: true,
      width: 2.4,
      cell: ({ row }) => <span title={row.original.role}>{row.original.role}</span>
    },
    {
      id: "salary",
      header: "Salary",
      dataType: "number",
      accessor: (row) => row.salary,
      update: (row, value) => ({ ...row, salary: Number(value) }),
      format: (value) => Number(value) === 0 ? "Team total" : currency.format(Number(value)),
      validate: ({ parsed }) => Number(parsed) < 0
        ? [{ code: "salary-negative", message: "Salary must be zero or greater" }]
        : [],
      sortable: true,
      filterable: true,
      aggregatable: ["sum", "average", "count"],
      width: 1
    },
    {
      id: "startDate",
      header: "Start date",
      dataType: "date",
      accessor: (row) => row.startDate,
      update: (row, value) => ({ ...row, startDate: String(value) }),
      sortable: true,
      filterable: true,
      width: 1
    },
    {
      id: "active",
      header: "Active",
      dataType: "boolean",
      accessor: (row) => row.active,
      update: (row, value) => ({ ...row, active: Boolean(value) }),
      filterable: true,
      width: 0.7
    }
  ], []);

  return (
    <main style={{ display: "grid", gap: 24, maxWidth: 1320, margin: "0 auto", padding: 24 }}>
      <header>
        <p style={{ margin: "0 0 4px", color: "#22795d", fontWeight: 700, letterSpacing: ".04em", textTransform: "uppercase" }}>
          Isolated React surface
        </p>
        <h1 style={{ margin: 0, color: "#172033", font: "700 28px/1.2 Inter, system-ui, sans-serif" }}>
          Embeddable employee data table
        </h1>
        <p style={{ maxWidth: 780, margin: "8px 0 0", color: "#5f6e82", font: "14px/1.5 Inter, system-ui, sans-serif" }}>
          Edit cells, expand the engineering team, filter inline, select rows, reorder columns, and open a column menu to sort, group, aggregate, or run the custom name action.
        </p>
        <p aria-live="polite" style={{ margin: "8px 0 0", color: "#5f6e82", font: "12px/1.4 Inter, system-ui, sans-serif" }}>
          Selected rows: {selectedRowIds.length === 0 ? "none" : selectedRowIds.join(", ")}
        </p>
      </header>

      <DataTable
        aria-label="Employee directory"
        rows={employees}
        columns={columns}
        getRowId={(row) => row.id}
        getSubRows={(row) => row.children}
        onRowsChange={(updater) => setEmployees((current) => updater(current))}
        onRowSelectionChange={setSelectedRowIds}
        rowSelection="multiple"
        inlineFilters
        layout="ratio"
        rowHeight="auto"
        defaultState={{
          expandedRowIds: ["eng"],
          sorting: [{ columnId: "name", direction: "asc" }]
        }}
        style={{ height: 520 }}
      />

      <section style={{ display: "grid", gap: 8 }}>
        <h2 style={{ margin: 0, color: "#172033", font: "650 18px/1.3 Inter, system-ui, sans-serif" }}>
          Independent summary instance
        </h2>
        <p style={{ margin: 0, color: "#5f6e82", font: "13px/1.4 Inter, system-ui, sans-serif" }}>
          This second table has its own selection, sort state, menus, live region, and history.
        </p>
        <DataTable
          aria-label="Team summary"
          rows={summaryRows}
          columns={summaryColumns}
          getRowId={(row) => row.id}
          rowSelection="single"
          layout="ratio"
          style={{ height: 280 }}
        />
      </section>
    </main>
  );
}

function normalizeEmployeeName(employee: Employee): Employee {
  return {
    ...employee,
    name: employee.name.replace(/\s+/g, " ").trim(),
    ...(employee.children ? { children: employee.children.map(normalizeEmployeeName) } : {})
  };
}

export default DataTableDemo;
