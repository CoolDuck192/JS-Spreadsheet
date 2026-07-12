import { Activity, StrictMode, type ReactNode } from "react";
import { renderToString } from "react-dom/server";
import { act, render, renderHook, waitFor } from "@testing-library/react";
import { describe, expect, it, vi } from "vitest";
import type { ColumnDef } from "./tableTypes";
import * as localSessionModule from "../table/local/RecordTableSession";
import {
  createLocalRecordTableSession,
  type RecordTableSession
} from "../table/local/RecordTableSession";
import { useTableSession } from "./useTableSession";
import { useTableSnapshot } from "./useTableSnapshot";

type Employee = { id: string; name: string };

const columns: readonly ColumnDef<Employee>[] = [{
  id: "name",
  header: "Name",
  dataType: "text",
  accessor: (row) => row.name,
  update: (row, value) => ({ ...row, name: String(value) })
}];

function options(rows: readonly Employee[], onRowsChange = vi.fn()) {
  return {
    source: { kind: "local" as const, rows, getRowId: (row: Employee) => row.id, onRowsChange },
    columns
  };
}

describe("useTableSession", () => {
  it("retains one owned session and synchronizes controlled options on rerender", () => {
    const onRowsChange = vi.fn();
    const { result, rerender } = renderHook(
      ({ rows }) => {
        const session = useTableSession(options(rows, onRowsChange));
        return { session, snapshot: useTableSnapshot(session) };
      },
      { initialProps: { rows: [{ id: "1", name: "Ada" }] } }
    );
    const first = result.current.session;

    rerender({ rows: [{ id: "1", name: "Grace" }] });

    expect(result.current.session).toBe(first);
    expect(result.current.snapshot.getCell("1", "name").displayValue).toBe("Grace");
  });

  it("survives StrictMode simulated effect cleanup", async () => {
    const wrapper = ({ children }: { children: ReactNode }) => <StrictMode>{children}</StrictMode>;
    const { result } = renderHook(() => useTableSession(options([{ id: "1", name: "Ada" }])), { wrapper });
    const destroy = vi.spyOn(result.current, "destroy");

    await act(async () => Promise.resolve());

    expect(destroy).not.toHaveBeenCalled();
    await expect(result.current.dispatch({ type: "set-selection", selection: null })).resolves.toMatchObject({
      status: "committed"
    });
  });

  it("creates only one owned session across StrictMode render replay", () => {
    const createSession = vi.spyOn(localSessionModule, "createLocalRecordTableSession");
    const wrapper = ({ children }: { children: ReactNode }) => <StrictMode>{children}</StrictMode>;

    renderHook(() => useTableSession(options([{ id: "1", name: "Ada" }])), { wrapper });

    expect(createSession).toHaveBeenCalledTimes(1);
  });

  it("does not publish controlled-state corrections during an abandoned render", () => {
    const onStateChange = vi.fn();
    function Probe() {
      useTableSession({
        ...options([{ id: "1", name: "Ada" }]),
        state: {
          grouping: [{ columnId: "name" }],
          pagination: { kind: "offset", offset: 0, limit: 25 }
        },
        onStateChange
      });
      return null;
    }

    renderToString(<Probe />);

    expect(onStateChange).not.toHaveBeenCalled();
  });

  it("publishes one controlled-state correction after a StrictMode commit", async () => {
    let rendering = false;
    const calledDuringRender: boolean[] = [];
    const onStateChange = vi.fn(() => { calledDuringRender.push(rendering); });
    const wrapper = ({ children }: { children: ReactNode }) => <StrictMode>{children}</StrictMode>;

    renderHook(() => {
      rendering = true;
      try {
        return useTableSession({
          ...options([{ id: "1", name: "Ada" }]),
          state: {
            grouping: [{ columnId: "name" }],
            pagination: { kind: "offset", offset: 0, limit: 25 }
          },
          onStateChange
        });
      } finally {
        rendering = false;
      }
    }, { wrapper });

    await waitFor(() => expect(onStateChange).toHaveBeenCalledTimes(1));
    expect(calledDuringRender).toEqual([false]);
  });

  it("destroys an owned session after real unmount", async () => {
    const { result, unmount } = renderHook(() => useTableSession(options([{ id: "1", name: "Ada" }])));
    const destroy = vi.spyOn(result.current, "destroy");

    unmount();
    await waitFor(() => expect(destroy).toHaveBeenCalledTimes(1));
  });

  it("recreates a destroyed raw table session when effects are restored", async () => {
    let current!: RecordTableSession<Employee, ColumnDef<Employee>>;
    function Probe() {
      current = useTableSession(options([{ id: "1", name: "Ada" }]));
      useTableSnapshot(current);
      return null;
    }

    const view = render(<Activity mode="visible"><Probe /></Activity>);
    const facade = current;
    view.rerender(<Activity mode="hidden"><Probe /></Activity>);
    await act(async () => Promise.resolve());
    view.rerender(<Activity mode="visible"><Probe /></Activity>);

    let dispatchResult: Awaited<ReturnType<typeof facade.dispatch>> | undefined;
    await act(async () => {
      dispatchResult = await current.dispatch({
        type: "set-selection",
        selection: {
          anchor: { rowId: "1", columnId: "name" },
          focus: { rowId: "1", columnId: "name" }
        }
      });
    });

    expect(current).toBe(facade);
    expect(dispatchResult).toMatchObject({ status: "committed" });
    expect(current.getSnapshot().selection).toEqual({
      anchor: { rowId: "1", columnId: "name" },
      focus: { rowId: "1", columnId: "name" }
    });
  });

  it("never destroys a supplied table session", async () => {
    const session = createLocalRecordTableSession(options([{ id: "1", name: "Ada" }]));
    const destroy = vi.spyOn(session, "destroy");
    const { unmount } = renderHook(() => useTableSnapshot(session));

    unmount();
    await act(async () => Promise.resolve());

    expect(destroy).not.toHaveBeenCalled();
    session.destroy();
  });

  it("subscribes through useSyncExternalStore without tearing", async () => {
    const rows = [{ id: "1", name: "Ada" }];
    const onRowsChange = vi.fn();
    const { result } = renderHook(() => {
      const session = useTableSession(options(rows, onRowsChange));
      return { session, snapshot: useTableSnapshot(session) };
    });

    await act(async () => {
      await result.current.session.dispatch({
        type: "set-selection",
        selection: {
          anchor: { rowId: "1", columnId: "name" },
          focus: { rowId: "1", columnId: "name" }
        }
      });
    });

    expect(result.current.snapshot.selection).toEqual({
      anchor: { rowId: "1", columnId: "name" },
      focus: { rowId: "1", columnId: "name" }
    });
    expect(result.current.snapshot.revision).toBe("1");
  });
});
