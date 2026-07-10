import { StrictMode, type ReactNode } from "react";
import { renderToString } from "react-dom/server";
import { act, renderHook, waitFor } from "@testing-library/react";
import { describe, expect, expectTypeOf, it, vi } from "vitest";
import type { QueryRequest, QueryResult } from "../table/core/query";
import { createTestRemoteSource } from "../table/remote/testUtils";
import type { RemoteTableSession } from "../table/remote/RemoteTableSession";
import type { ColumnDef } from "./tableTypes";
import { useTableSession } from "./useTableSession";
import { useTableSnapshot } from "./useTableSnapshot";

type Row = { id: string; name: string };

const rows: readonly Row[] = [{ id: "1", name: "Ada" }];
const columns: readonly ColumnDef<Row>[] = [{
  id: "name",
  header: "Name",
  accessor: (row) => row.name,
  update: (row, value) => ({ ...row, name: String(value) })
}];

describe("useTableSession remote overload", () => {
  it("is render-pure for an abandoned server render", () => {
    const query = vi.fn(async (request: QueryRequest) => result(request));
    const source = createTestRemoteSource<Row>({ query });
    function Probe() {
      useTableSession({ source, columns });
      return null;
    }

    renderToString(<Probe />);

    expect(query).not.toHaveBeenCalled();
  });

  it("starts after commit and retains a typed owned remote session", async () => {
    const query = vi.fn(async (request: QueryRequest) => result(request));
    const source = createTestRemoteSource<Row>({ query });
    const { result: hook } = renderHook(() => {
      const session = useTableSession({ source, columns });
      return { session, snapshot: useTableSnapshot(session) };
    });

    expectTypeOf(hook.current.session).toMatchTypeOf<RemoteTableSession<Row, ColumnDef<Row>>>();
    await waitFor(() => expect(hook.current.snapshot.status.phase).toBe("ready"));
    expect(query).toHaveBeenCalledTimes(1);
    expect(hook.current.snapshot.getCell("1", "name").displayValue).toBe("Ada");
  });

  it("applies controlled rerenders only after commit", async () => {
    const query = vi.fn(async (request: QueryRequest) => result(request));
    const source = createTestRemoteSource<Row>({ query });
    const { rerender } = renderHook(
      ({ descending }) => useTableSession({
        source,
        columns,
        state: {
          sorting: descending ? [{ columnId: "name", direction: "desc" }] : []
        }
      }),
      { initialProps: { descending: false } }
    );
    await waitFor(() => expect(query).toHaveBeenCalledTimes(1));

    rerender({ descending: true });
    await waitFor(() => expect(query).toHaveBeenCalledTimes(2));
    expect(query.mock.calls[1][0].sorting).toEqual([{ columnId: "name", direction: "desc" }]);
  });

  it("survives StrictMode stop/start replay without destroying the live session", async () => {
    const query = vi.fn(async (request: QueryRequest) => result(request));
    const source = createTestRemoteSource<Row>({ query });
    const wrapper = ({ children }: { children: ReactNode }) => <StrictMode>{children}</StrictMode>;
    const { result: hook } = renderHook(() => useTableSession({ source, columns }), { wrapper });
    const destroy = vi.spyOn(hook.current, "destroy");

    await waitFor(() => expect(query.mock.calls.length).toBeGreaterThanOrEqual(1));
    await act(async () => Promise.resolve());
    expect(destroy).not.toHaveBeenCalled();
    expect(hook.current.getDiagnostics().destroyed).toBe(false);
  });

  it("restarts the same session for a replacement remote source", async () => {
    const firstQuery = vi.fn(async (request: QueryRequest) => result(request, "r1", "Ada"));
    const secondQuery = vi.fn(async (request: QueryRequest) => result(request, "r2", "Grace"));
    const first = createTestRemoteSource<Row>({ query: firstQuery });
    const second = createTestRemoteSource<Row>({ query: secondQuery });
    const { result: hook, rerender } = renderHook(
      ({ source }) => useTableSession({ source, columns }),
      { initialProps: { source: first } }
    );
    const session = hook.current;
    await waitFor(() => expect(firstQuery).toHaveBeenCalledTimes(1));

    rerender({ source: second });
    await waitFor(() => expect(secondQuery).toHaveBeenCalledTimes(1));
    expect(hook.current).toBe(session);
    await waitFor(() => expect(hook.current.getSnapshot().getCell("1", "name").displayValue).toBe("Grace"));
  });

  it("destroys the owned remote session after final unmount", async () => {
    const source = createTestRemoteSource<Row>({ query: async (request) => result(request) });
    const { result: hook, unmount } = renderHook(() => useTableSession({ source, columns }));
    const destroy = vi.spyOn(hook.current, "destroy");

    unmount();

    await waitFor(() => expect(destroy).toHaveBeenCalledTimes(1));
  });
});

function result(
  request: QueryRequest,
  revision = "r1",
  name = "Ada"
): QueryResult<Row> {
  if (request.pagination.kind !== "offset") throw new Error("expected offset query");
  const data = [{ id: "1", name }];
  return {
    items: data.map((row) => ({ kind: "data" as const, id: row.id, original: row, depth: 0 })),
    revision,
    completeness: "loadedRows",
    pageInfo: {
      ...request.pagination,
      total: { kind: "known", value: data.length },
      hasMore: false
    }
  };
}
