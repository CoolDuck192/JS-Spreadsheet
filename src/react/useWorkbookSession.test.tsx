import { Activity, StrictMode, useSyncExternalStore, type ReactNode } from "react";
import { act, render, renderHook, waitFor } from "@testing-library/react";
import { describe, expect, it, vi } from "vitest";
import { createFormulaEngine } from "../lib/formulaEngine";
import { WORKBOOK_STORAGE_KEY } from "../lib/persistence";
import { createBlankWorkbook, getCellContent, setCellContent } from "../lib/workbook";
import type { WorkbookModel } from "../types";
import type { WorkbookStorage } from "./Spreadsheet";
import {
  BrowserWorkbookStorageCapacityError,
  createBrowserWorkbookStorage
} from "./browserWorkbookStorage";
import { useWorkbookSession } from "./useWorkbookSession";

function deferred<T>() {
  let resolve!: (value: T | PromiseLike<T>) => void;
  let reject!: (reason?: unknown) => void;
  const promise = new Promise<T>((resolvePromise, rejectPromise) => {
    resolve = resolvePromise;
    reject = rejectPromise;
  });
  return { promise, resolve, reject };
}

function withCell(value: string): WorkbookModel {
  const workbook = createBlankWorkbook();
  return setCellContent(workbook, workbook.activeSheetId, "A1", value);
}

function useLiveWorkbook(session: ReturnType<typeof useWorkbookSession>) {
  return useSyncExternalStore(session.subscribe, session.getSnapshot, session.getSnapshot).workbook;
}

describe("useWorkbookSession", () => {
  it("retains one uncontrolled session and publishes to shared subscribers", () => {
    const initial = withCell("initial");
    const { result, rerender } = renderHook(() => {
      const session = useWorkbookSession({ defaultWorkbook: initial, storage: false });
      return { session, workbook: useLiveWorkbook(session) };
    });
    const first = result.current.session;
    const listenerA = vi.fn();
    const listenerB = vi.fn();
    const unsubscribeA = first.subscribe(listenerA);
    const unsubscribeB = first.subscribe(listenerB);

    act(() => {
      first.dispatch({
        type: "cell.set",
        sheetId: initial.activeSheetId,
        address: "A1",
        input: "next"
      });
    });
    rerender();

    expect(result.current.session).toBe(first);
    expect(getCellContent(result.current.workbook, initial.activeSheetId, "A1")).toBe("next");
    expect(listenerA).toHaveBeenCalledTimes(1);
    expect(listenerB).toHaveBeenCalledTimes(1);
    unsubscribeA();
    unsubscribeB();
  });

  it("synchronizes controlled rerenders and acknowledges user revisions", async () => {
    const onWorkbookChange = vi.fn();
    const first = withCell("first");
    const second = withCell("second");
    const { result, rerender } = renderHook(
      ({ workbook }) => useWorkbookSession({ workbook, onWorkbookChange, storage: false }),
      { initialProps: { workbook: first } }
    );

    act(() => {
      result.current.dispatch({
        type: "cell.set",
        sheetId: first.activeSheetId,
        address: "A1",
        input: "user"
      });
    });
    expect(onWorkbookChange).toHaveBeenCalledTimes(1);
    expect(getCellContent(onWorkbookChange.mock.calls[0][0], first.activeSheetId, "A1")).toBe("user");
    await waitFor(() => expect(result.current.getSnapshot().workbook).toBe(first));

    rerender({ workbook: second });
    await waitFor(() => expect(result.current.getSnapshot().workbook).toBe(second));
    expect(onWorkbookChange).toHaveBeenCalledTimes(1);
  });

  it("reports every public dispatch outcome without publishing a rejected change", () => {
    const onCommandResult = vi.fn();
    const onWorkbookChangeEvent = vi.fn();
    const initial = createBlankWorkbook();
    const { result } = renderHook(() => useWorkbookSession({
      defaultWorkbook: initial,
      storage: false,
      onCommandResult,
      onWorkbookChangeEvent
    }));

    act(() => {
      result.current.dispatch({
        id: "stale-command",
        expectedRevision: "missing",
        intent: {
          type: "cell.set",
          sheetId: initial.activeSheetId,
          address: "A1",
          input: "rejected"
        }
      });
    });

    expect(onCommandResult).toHaveBeenCalledTimes(1);
    expect(onCommandResult.mock.calls[0][0].result.status).toBe("conflict");
    expect(onWorkbookChangeEvent).not.toHaveBeenCalled();
    expect(result.current.getSnapshot().workbook).toBe(initial);
  });

  it("does not touch browser storage when storage is false", async () => {
    const getItem = vi.spyOn(Storage.prototype, "getItem");
    const setItem = vi.spyOn(Storage.prototype, "setItem");
    const { unmount } = renderHook(() => useWorkbookSession({ storage: false }));
    await act(async () => Promise.resolve());
    unmount();
    await act(async () => Promise.resolve());
    expect(getItem).not.toHaveBeenCalled();
    expect(setItem).not.toHaveBeenCalled();
  });

  it("honors the default workbook after one empty browser-storage load", async () => {
    window.localStorage.removeItem(WORKBOOK_STORAGE_KEY);
    const getItem = vi.spyOn(Storage.prototype, "getItem");
    const initial = withCell("default");
    const { result, unmount } = renderHook(() => useWorkbookSession({ defaultWorkbook: initial }));

    await act(async () => Promise.resolve());

    expect(result.current.getSnapshot().workbook).toBe(initial);
    expect(
      getItem.mock.calls.filter(([key]) => key === WORKBOOK_STORAGE_KEY)
    ).toHaveLength(1);

    unmount();
    await act(async () => Promise.resolve());
    getItem.mockRestore();
  });

  it("creates a blank workbook when browser storage and the default are absent", async () => {
    window.localStorage.removeItem(WORKBOOK_STORAGE_KEY);
    const { result, unmount } = renderHook(() => useWorkbookSession());

    await act(async () => Promise.resolve());

    const workbook = result.current.getSnapshot().workbook;
    expect(workbook.sheets).toHaveLength(1);
    expect(workbook.sheets[0]?.cells).toEqual({});

    unmount();
    await act(async () => Promise.resolve());
  });

  it("discards stale async hydration after an uncontrolled edit", async () => {
    const hydration = deferred<WorkbookModel | null>();
    const storage: WorkbookStorage = { load: () => hydration.promise, save: vi.fn() };
    const initial = withCell("initial");
    const stored = withCell("stored");
    const { result } = renderHook(() => useWorkbookSession({ defaultWorkbook: initial, storage }));

    act(() => {
      result.current.dispatch({
        type: "cell.set",
        sheetId: initial.activeSheetId,
        address: "A1",
        input: "edited"
      });
    });
    await act(async () => hydration.resolve(stored));

    expect(getCellContent(result.current.getSnapshot().workbook, initial.activeSheetId, "A1")).toBe("edited");
  });

  it("never accepts storage hydration as controlled authority", async () => {
    const hydration = deferred<WorkbookModel | null>();
    const storage: WorkbookStorage = { load: () => hydration.promise, save: vi.fn() };
    const controlled = withCell("host");
    const { result } = renderHook(() => useWorkbookSession({
      workbook: controlled,
      onWorkbookChange: vi.fn(),
      storage
    }));

    await act(async () => hydration.resolve(withCell("stored")));
    expect(result.current.getSnapshot().workbook).toBe(controlled);
  });

  it("reports a load failure and blocks later autosave attempts", async () => {
    const save = vi.fn();
    const onError = vi.fn();
    const storage: WorkbookStorage = {
      load() {
        throw new Error("secret storage detail");
      },
      save
    };
    const initial = createBlankWorkbook();
    const { result, unmount } = renderHook(() => useWorkbookSession({
      defaultWorkbook: initial,
      storage,
      onError
    }));

    await waitFor(() => expect(result.current.getSnapshot().persistence).toEqual({
      status: "failed",
      operation: "load",
      message: "Workbook could not be loaded. Autosave is paused to protect stored data."
    }));
    expect(onError).toHaveBeenCalledWith({
      code: "storage.load.failed",
      message: "Workbook could not be loaded. Autosave is paused to protect stored data.",
      recoverable: true
    });

    act(() => {
      result.current.dispatch({
        type: "cell.set",
        sheetId: initial.activeSheetId,
        address: "A1",
        input: "must stay in memory"
      });
    });
    await act(async () => Promise.resolve());
    act(() => window.dispatchEvent(new Event("beforeunload")));
    unmount();
    await act(async () => Promise.resolve());
    expect(save).not.toHaveBeenCalled();
  });

  it("serializes saves and coalesces superseded queued revisions", async () => {
    const firstSave = deferred<void>();
    const secondSave = deferred<void>();
    const save = vi.fn()
      .mockImplementationOnce(() => firstSave.promise)
      .mockImplementationOnce(() => secondSave.promise);
    const storage: WorkbookStorage = { load: () => null, save };
    const initial = createBlankWorkbook();
    const { result } = renderHook(() => useWorkbookSession({ defaultWorkbook: initial, storage }));
    await act(async () => Promise.resolve());

    act(() => {
      result.current.dispatch({ type: "cell.set", sheetId: initial.activeSheetId, address: "A1", input: "one" });
    });
    await waitFor(() => expect(save).toHaveBeenCalledTimes(1));
    act(() => {
      result.current.dispatch({ type: "cell.set", sheetId: initial.activeSheetId, address: "A1", input: "two" });
      result.current.dispatch({ type: "cell.set", sheetId: initial.activeSheetId, address: "A1", input: "three" });
    });
    expect(save).toHaveBeenCalledTimes(1);

    await act(async () => firstSave.resolve());
    await waitFor(() => expect(save).toHaveBeenCalledTimes(2));
    expect(getCellContent(save.mock.calls[1][0], initial.activeSheetId, "A1")).toBe("three");
    await act(async () => secondSave.resolve());
    await waitFor(() => expect(result.current.getSnapshot().persistence.status).toBe("idle"));
  });

  it("reports quota failure and later recovery without creating workbook revisions", async () => {
    const save = vi.fn()
      .mockRejectedValueOnce(new Error("secret quota details"))
      .mockResolvedValueOnce(undefined);
    const storage: WorkbookStorage = { load: () => null, save };
    const initial = createBlankWorkbook();
    const { result } = renderHook(() => useWorkbookSession({ defaultWorkbook: initial, storage }));
    await act(async () => Promise.resolve());
    const initialRevision = result.current.getSnapshot().revision;

    act(() => {
      result.current.dispatch({ type: "cell.set", sheetId: initial.activeSheetId, address: "A1", input: "one" });
    });
    await waitFor(() => expect(result.current.getSnapshot().persistence.status).toBe("failed"));
    const failedRevision = result.current.getSnapshot().revision;
    expect(failedRevision).not.toBe(initialRevision);
    expect(result.current.getSnapshot().persistence.message).not.toContain("secret");

    act(() => {
      result.current.dispatch({ type: "cell.set", sheetId: initial.activeSheetId, address: "A2", input: "two" });
    });
    await waitFor(() => expect(result.current.getSnapshot().persistence.status).toBe("idle"));
    expect(Number(result.current.getSnapshot().revision) - Number(failedRevision)).toBe(1);
  });

  it("rejects workbooks above the 100,000-cell browser autosave cap before writing", () => {
    const workbook = createBlankWorkbook();
    const cells: Record<string, number> = {};
    for (let row = 1; row <= 100_001; row += 1) {
      cells[`A${row}`] = row;
    }
    workbook.sheets[0] = {
      ...workbook.sheets[0],
      rowCount: 100_001,
      cells
    };
    const setItem = vi.fn();
    const storage = createBrowserWorkbookStorage({ setItem } as unknown as Storage);

    expect(() => storage.save(workbook)).toThrow(BrowserWorkbookStorageCapacityError);
    expect(setItem).not.toHaveBeenCalled();
  });

  it("surfaces an actionable browser autosave capacity message", async () => {
    const onError = vi.fn();
    const storage: WorkbookStorage = {
      load: () => null,
      save() {
        throw new BrowserWorkbookStorageCapacityError();
      }
    };
    const initial = createBlankWorkbook();
    const { result } = renderHook(() => useWorkbookSession({
      defaultWorkbook: initial,
      storage,
      onError
    }));
    await act(async () => Promise.resolve());

    act(() => {
      result.current.dispatch({
        type: "cell.set",
        sheetId: initial.activeSheetId,
        address: "A1",
        input: "too large"
      });
    });

    const message = "Workbook is too large for browser autosave. Use Export XLSX to save your work.";
    await waitFor(() => expect(result.current.getSnapshot().persistence).toEqual({
      status: "failed",
      operation: "save",
      message
    }));
    expect(onError).toHaveBeenCalledWith({
      code: "storage.save.failed",
      message,
      recoverable: true
    });
  });

  it("flushes the latest edit when unmounted before the autosave effect runs", async () => {
    const save = vi.fn();
    const storage: WorkbookStorage = { load: () => null, save };
    const initial = createBlankWorkbook();
    const { result, unmount } = renderHook(() =>
      useWorkbookSession({ defaultWorkbook: initial, storage })
    );
    await act(async () => Promise.resolve());

    act(() => {
      result.current.dispatch({
        type: "cell.set",
        sheetId: initial.activeSheetId,
        address: "A1",
        input: "latest"
      });
      unmount();
    });
    await act(async () => Promise.resolve());

    expect(save).toHaveBeenCalledTimes(1);
    expect(getCellContent(save.mock.calls[0][0], initial.activeSheetId, "A1")).toBe("latest");
  });

  it("drains the latest unmounted snapshot after an in-flight save", async () => {
    const firstSave = deferred<void>();
    const save = vi.fn()
      .mockImplementationOnce(() => firstSave.promise)
      .mockResolvedValueOnce(undefined);
    const storage: WorkbookStorage = { load: () => null, save };
    const initial = createBlankWorkbook();
    const { result, unmount } = renderHook(() =>
      useWorkbookSession({ defaultWorkbook: initial, storage })
    );
    await act(async () => Promise.resolve());

    act(() => {
      result.current.dispatch({
        type: "cell.set",
        sheetId: initial.activeSheetId,
        address: "A1",
        input: "first"
      });
    });
    await waitFor(() => expect(save).toHaveBeenCalledTimes(1));

    act(() => {
      result.current.dispatch({
        type: "cell.set",
        sheetId: initial.activeSheetId,
        address: "A1",
        input: "latest"
      });
      unmount();
    });
    await act(async () => Promise.resolve());
    expect(save).toHaveBeenCalledTimes(1);

    await act(async () => firstSave.resolve());
    await waitFor(() => expect(save).toHaveBeenCalledTimes(2));
    expect(getCellContent(save.mock.calls[1][0], initial.activeSheetId, "A1")).toBe("latest");
  });

  it("flushes synchronously through beforeunload before passive effects", async () => {
    const save = vi.fn();
    const storage: WorkbookStorage = { load: () => null, save };
    const initial = createBlankWorkbook();
    const { result, unmount } = renderHook(() =>
      useWorkbookSession({ defaultWorkbook: initial, storage })
    );
    await act(async () => Promise.resolve());
    let savedDuringEvent = false;

    act(() => {
      result.current.dispatch({
        type: "cell.set",
        sheetId: initial.activeSheetId,
        address: "A1",
        input: "leaving"
      });
      window.dispatchEvent(new Event("beforeunload"));
      savedDuringEvent = save.mock.calls.length === 1;
    });

    expect(savedDuringEvent).toBe(true);
    expect(getCellContent(save.mock.calls[0][0], initial.activeSheetId, "A1")).toBe("leaving");
    unmount();
    await act(async () => Promise.resolve());
  });

  it("survives StrictMode replay and destroys an owned session after final unmount", async () => {
    const destroy = vi.fn();
    const formulaEngineFactory: typeof createFormulaEngine = (workbook) => {
      const engine = createFormulaEngine(workbook);
      return { ...engine, destroy: () => { destroy(); engine.destroy(); } };
    };
    const wrapper = ({ children }: { children: ReactNode }) => <StrictMode>{children}</StrictMode>;
    const { result, unmount } = renderHook(
      () => useWorkbookSession({ storage: false, services: { formulaEngineFactory } }),
      { wrapper }
    );
    await act(async () => Promise.resolve());
    expect(result.current.dispatch({ type: "selection.set", selection: {
      start: { row: 1, column: 1 }, end: { row: 1, column: 1 }
    } }).status).toBe("committed");
    expect(destroy).not.toHaveBeenCalled();

    unmount();
    await act(async () => Promise.resolve());
    expect(destroy).toHaveBeenCalledTimes(1);
    expect(result.current.dispatch({ type: "selection.set", selection: {
      start: { row: 2, column: 2 }, end: { row: 2, column: 2 }
    } }).status).toBe("rejected");
  });

  it("recreates a destroyed raw session when effects are restored", async () => {
    const initial = withCell("initial");
    let current!: {
      session: ReturnType<typeof useWorkbookSession>;
      workbook: WorkbookModel;
    };
    function Probe() {
      const session = useWorkbookSession({ defaultWorkbook: initial, storage: false });
      current = { session, workbook: useLiveWorkbook(session) };
      return null;
    }

    const view = render(<Activity mode="visible"><Probe /></Activity>);
    const facade = current.session;
    view.rerender(<Activity mode="hidden"><Probe /></Activity>);
    await act(async () => Promise.resolve());
    view.rerender(<Activity mode="visible"><Probe /></Activity>);

    let dispatchResult: ReturnType<typeof facade.dispatch> | undefined;
    act(() => {
      dispatchResult = current.session.dispatch({
        type: "cell.set",
        sheetId: initial.activeSheetId,
        address: "A1",
        input: "after reveal"
      });
    });

    expect(current.session).toBe(facade);
    expect(dispatchResult).toMatchObject({ status: "committed" });
    expect(getCellContent(current.workbook, initial.activeSheetId, "A1")).toBe("after reveal");
  });

  it("does not publish a late save completion after unmount", async () => {
    const pending = deferred<void>();
    const storage: WorkbookStorage = { load: () => null, save: () => pending.promise };
    const initial = createBlankWorkbook();
    const { result, unmount } = renderHook(() => useWorkbookSession({ defaultWorkbook: initial, storage }));
    await act(async () => Promise.resolve());
    act(() => {
      result.current.dispatch({ type: "cell.set", sheetId: initial.activeSheetId, address: "A1", input: "saved" });
    });
    await waitFor(() => expect(result.current.getSnapshot().persistence.status).toBe("saving"));
    const revision = result.current.getSnapshot().revision;

    unmount();
    await act(async () => Promise.resolve());
    await act(async () => pending.resolve());
    expect(result.current.getSnapshot().revision).toBe(revision);
    expect(result.current.getSnapshot().persistence.status).toBe("saving");
  });
});
