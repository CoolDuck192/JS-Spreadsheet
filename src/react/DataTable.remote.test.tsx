import { render, screen, waitFor } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import { describe, expect, it, vi } from "vitest";
import type { CommandResult } from "../core/commands/types";
import { createLocalRecordTableSession } from "../table/local/RecordTableSession";
import type { TableSession, TableViewSnapshot } from "../table/core/types";
import { DataTable } from "./DataTable";
import type { ColumnDef } from "./tableTypes";

type Row = { id: string; salary: number };

const rows: readonly Row[] = [{ id: "row-1", salary: 100 }];
const columns: readonly ColumnDef<Row>[] = [{
  id: "salary",
  header: "Salary",
  dataType: "number",
  accessor: (row) => row.salary,
  update: (row, value) => ({ ...row, salary: Number(value) }),
  sortable: true,
  filterable: true
}];

describe("DataTable remote lifecycle surfaces", () => {
  it("announces loading while retaining the last projected rows", () => {
    const { session } = createSession({ status: { phase: "loading" } });
    render(<DataTable aria-label="Remote salaries" session={session} />);

    expect(screen.getAllByRole("status").some((status) => status.textContent === "Loading rows")).toBe(true);
    expect(screen.getByRole("gridcell", { name: "row-1 Salary" })).toHaveTextContent("100");
  });

  it("retries an offline query and disables duplicate retry while it is in flight", async () => {
    const user = userEvent.setup();
    let resolve!: (result: CommandResult) => void;
    const pending = new Promise<CommandResult>((done) => { resolve = done; });
    const dispatch = vi.fn(() => pending);
    const { session } = createSession({
      status: { phase: "error", message: "Offline" },
      issues: [{ code: "OFFLINE", message: "Offline" }]
    }, dispatch);
    render(<DataTable aria-label="Remote salaries" session={session} />);

    const retry = screen.getByRole("button", { name: "Retry loading rows" });
    await user.click(retry);
    expect(retry).toBeDisabled();
    expect(dispatch).toHaveBeenCalledWith({ type: "refresh" });
    resolve({ status: "committed", revision: "2", changed: true });
    await waitFor(() => expect(retry).toBeEnabled());
  });

  it.each([
    ["Reload server value", {
      type: "reload-authoritative",
      operationId: "operation-1",
      rowId: "row-1"
    }],
    ["Retry my change", {
      type: "retry-with-revision",
      operationId: "operation-1",
      rowId: "row-1",
      expectedRevision: "r3"
    }]
  ] as const)("dispatches exact stable conflict action for %s", async (buttonName, intent) => {
    const user = userEvent.setup();
    const dispatch = vi.fn(async () => ({ status: "committed", revision: "4", changed: true } as const));
    const { session } = createSession({
      conflicts: [{
        operationId: "operation-1",
        rowId: "row-1",
        columnId: "salary",
        attemptedValue: 120,
        authoritativeValue: 115,
        current: { id: "row-1", salary: 115 },
        revision: "r3"
      }]
    }, dispatch);
    render(<DataTable aria-label="Remote salaries" session={session} />);

    expect(screen.getByRole("alert")).toHaveTextContent("Salary changed on the server");
    await user.click(screen.getByRole("button", { name: buttonName }));
    expect(dispatch).toHaveBeenCalledWith(intent);
  });

  it("announces a stale conflict rejection without discarding the conflict panel", async () => {
    const user = userEvent.setup();
    const { session } = createSession({
      conflicts: [{
        operationId: "operation-1", rowId: "row-1", columnId: "salary",
        current: rows[0], revision: "r3"
      }]
    }, vi.fn(async () => ({
      status: "rejected",
      reason: "validation",
      issues: [{ code: "REMOTE_CONFLICT_NOT_CURRENT", message: "Conflict is no longer current" }]
    })));
    render(<DataTable aria-label="Remote salaries" session={session} />);

    await user.click(screen.getByRole("button", { name: "Retry my change" }));
    expect(screen.getAllByRole("alert").some((alert) => alert.textContent?.includes("Conflict is no longer current"))).toBe(true);
    expect(screen.getByRole("button", { name: "Reload server value" })).toBeInTheDocument();
  });

  it("disables both conflict actions after a pending resolution is accepted", async () => {
    const user = userEvent.setup();
    const { session } = createSession({
      conflicts: [{
        operationId: "operation-1", rowId: "row-1", columnId: "salary",
        current: rows[0], revision: "r3"
      }]
    }, vi.fn(async () => ({ status: "pending", operationId: "resolution-1" })));
    render(<DataTable aria-label="Remote salaries" session={session} />);

    await user.click(screen.getByRole("button", { name: "Reload server value" }));
    expect(screen.getByRole("button", { name: "Reload server value" })).toBeDisabled();
    expect(screen.getByRole("button", { name: "Retry my change" })).toBeDisabled();
  });

  it("exposes cursor, infinite, and unknown-total navigation without guessing counts", async () => {
    const user = userEvent.setup();
    const dispatch = vi.fn(async () => ({ status: "committed", revision: "2", changed: true } as const));
    const cursorHarness = createSession({
      state: { kind: "cursor" },
      pageInfo: { kind: "cursor", nextCursor: "next", previousCursor: "previous", total: { kind: "unknown" } },
      totalRowCount: { kind: "unknown" }
    }, dispatch);
    const { unmount } = render(<DataTable aria-label="Cursor salaries" session={cursorHarness.session} />);
    expect(screen.getByLabelText("Table row total")).toHaveTextContent("Total rows unknown");
    await user.click(screen.getByRole("button", { name: "Next cursor page" }));
    expect(dispatch).toHaveBeenCalledWith({
      type: "set-pagination",
      pagination: { kind: "cursor", cursor: "next", limit: 50 }
    });
    unmount();

    const infiniteDispatch = vi.fn(async () => ({ status: "committed", revision: "2", changed: true } as const));
    const infiniteHarness = createSession({
      state: { kind: "infinite" },
      pageInfo: { kind: "infinite", nextCursor: "more", loadedCount: 1, total: { kind: "unknown" } },
      totalRowCount: { kind: "unknown" }
    }, infiniteDispatch);
    render(<DataTable aria-label="Infinite salaries" session={infiniteHarness.session} />);
    await user.click(screen.getByRole("button", { name: "Load more rows" }));
    expect(infiniteDispatch).toHaveBeenCalledWith({
      type: "set-pagination",
      pagination: { kind: "infinite", after: "more", limit: 50 }
    });
  });
});

type SnapshotOverrides = {
  status?: TableViewSnapshot<Row, ColumnDef<Row>>["status"];
  issues?: TableViewSnapshot<Row, ColumnDef<Row>>["issues"];
  conflicts?: TableViewSnapshot<Row, ColumnDef<Row>>["conflicts"];
  pendingOperations?: TableViewSnapshot<Row, ColumnDef<Row>>["pendingOperations"];
  pageInfo?: TableViewSnapshot<Row, ColumnDef<Row>>["pageInfo"];
  totalRowCount?: TableViewSnapshot<Row, ColumnDef<Row>>["totalRowCount"];
  state?: { kind: "cursor" | "infinite" };
};

function createSession(
  overrides: SnapshotOverrides,
  dispatch: (...args: any[]) => Promise<any> = vi.fn(async () => ({ status: "committed", revision: "1", changed: false } as const))
) {
  const base = createLocalRecordTableSession({
    source: { kind: "local", rows, getRowId: (row) => row.id },
    columns
  });
  const baseSnapshot = base.getSnapshot() as TableViewSnapshot<Row, ColumnDef<Row>>;
  const pagination = overrides.state?.kind === "cursor"
    ? { kind: "cursor" as const, limit: 50 }
    : overrides.state?.kind === "infinite"
      ? { kind: "infinite" as const, limit: 50 }
      : baseSnapshot.state.pagination;
  const snapshot: TableViewSnapshot<Row, ColumnDef<Row>> = {
    ...baseSnapshot,
    ...overrides,
    state: { ...baseSnapshot.state, pagination },
    status: overrides.status ?? baseSnapshot.status,
    issues: overrides.issues ?? baseSnapshot.issues,
    conflicts: overrides.conflicts ?? baseSnapshot.conflicts,
    pendingOperations: overrides.pendingOperations ?? baseSnapshot.pendingOperations,
    pageInfo: overrides.pageInfo ?? baseSnapshot.pageInfo,
    totalRowCount: overrides.totalRowCount ?? baseSnapshot.totalRowCount
  };
  const session: TableSession<Row, ColumnDef<Row>> = {
    getSnapshot: () => snapshot,
    subscribe: () => () => {},
    dispatch: (intent) => dispatch(intent) as Promise<CommandResult<Row>>,
    refresh: async () => { await dispatch({ type: "refresh" }); },
    undo: () => base.undo(),
    redo: () => base.redo(),
    export: (options) => base.export(options),
    destroy: () => base.destroy()
  };
  return { session, base };
}
