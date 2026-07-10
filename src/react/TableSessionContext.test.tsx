import { StrictMode, type ReactNode } from "react";
import { act, renderHook, waitFor } from "@testing-library/react";
import { describe, expect, it, vi } from "vitest";
import type { ColumnDef } from "./tableTypes";
import { createLocalRecordTableSession } from "../table/local/RecordTableSession";
import { TableSessionProvider, useTableSessionContext } from "./TableSessionContext";
import { useTableSnapshot } from "./useTableSnapshot";

type Row = { id: string; value: string };
const columns: readonly ColumnDef<Row>[] = [{
  id: "value",
  header: "Value",
  accessor: (row) => row.value,
  update: (row, value) => ({ ...row, value: String(value) })
}];

function createSession(value: string) {
  return createLocalRecordTableSession({
    source: { kind: "local", rows: [{ id: value, value }], getRowId: (row: Row) => row.id },
    columns
  });
}

describe("TableSessionContext", () => {
  it("throws the exact error outside a provider", () => {
    expect(() => renderHook(() => useTableSessionContext<Row>())).toThrow(
      "useTableSessionContext must be used inside TableSessionProvider"
    );
  });

  it("isolates nested providers", () => {
    const outer = createSession("outer");
    const inner = createSession("inner");
    const wrapper = ({ children }: { children: ReactNode }) => (
      <TableSessionProvider session={outer}>
        <TableSessionProvider session={inner}>{children}</TableSessionProvider>
      </TableSessionProvider>
    );

    const { result } = renderHook(() => useTableSessionContext<Row>(), { wrapper });

    expect(result.current).toBe(inner);
    outer.destroy();
    inner.destroy();
  });

  it("publishes snapshot updates through the context session", async () => {
    const session = createSession("row");
    const wrapper = ({ children }: { children: ReactNode }) => (
      <TableSessionProvider session={session}>{children}</TableSessionProvider>
    );
    const { result } = renderHook(() => useTableSnapshot(useTableSessionContext<Row>()), { wrapper });

    await act(async () => {
      await session.dispatch({
        type: "set-selection",
        selection: {
          anchor: { rowId: "row", columnId: "value" },
          focus: { rowId: "row", columnId: "value" }
        }
      });
    });

    expect(result.current.selection?.focus).toEqual({ rowId: "row", columnId: "value" });
    session.destroy();
  });

  it("never destroys a host session during StrictMode cleanup or provider unmount", async () => {
    const session = createSession("row");
    const destroy = vi.spyOn(session, "destroy");
    const wrapper = ({ children }: { children: ReactNode }) => (
      <StrictMode>
        <TableSessionProvider session={session}>{children}</TableSessionProvider>
      </StrictMode>
    );
    const { unmount } = renderHook(() => useTableSessionContext<Row>(), { wrapper });

    unmount();
    await act(async () => Promise.resolve());

    expect(destroy).not.toHaveBeenCalled();
    session.destroy();
    await waitFor(() => expect(destroy).toHaveBeenCalledTimes(1));
  });
});
