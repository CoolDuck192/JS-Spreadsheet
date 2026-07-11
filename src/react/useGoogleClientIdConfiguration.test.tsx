import { act, renderHook, waitFor } from "@testing-library/react";
import { describe, expect, it, vi } from "vitest";
import type { GoogleClientIdStorage } from "../core/workbook/services";
import { useGoogleClientIdConfiguration } from "./useGoogleClientIdConfiguration";

const CLIENT_ID = "123-abc.apps.googleusercontent.com";
const STORED_CLIENT_ID = "stored.apps.googleusercontent.com";

function deferred<T>() {
  let resolve!: (value: T) => void;
  let reject!: (reason?: unknown) => void;
  const promise = new Promise<T>((resolvePromise, rejectPromise) => {
    resolve = resolvePromise;
    reject = rejectPromise;
  });
  return { promise, resolve, reject };
}

function storage(overrides: Partial<GoogleClientIdStorage> = {}): GoogleClientIdStorage {
  return {
    load: vi.fn().mockResolvedValue(null),
    save: vi.fn().mockResolvedValue(undefined),
    clear: vi.fn().mockResolvedValue(undefined),
    ...overrides
  };
}

describe("useGoogleClientIdConfiguration", () => {
  it("loads a valid stored ID once and exposes focused configuration state", async () => {
    const load = vi.fn().mockResolvedValue(` ${STORED_CLIENT_ID} `);
    const clientIdStorage = storage({ load });
    const controller = renderHook(() => useGoogleClientIdConfiguration({
      storage: clientIdStorage,
      ignoreLoadFailure: false,
      onError: vi.fn()
    }));

    await waitFor(() => expect(controller.result.current.storageStatus).toBe("loaded"));
    expect(controller.result.current.editableClientId).toEqual({
      value: STORED_CLIENT_ID,
      source: "stored"
    });
    expect(controller.result.current.clientIdDraft).toBe(STORED_CLIENT_ID);

    controller.rerender();
    expect(load).toHaveBeenCalledTimes(1);
  });

  it("retains the usable ID when clear fails and reports an explicit retained outcome", async () => {
    const pendingClear = deferred<void>();
    const clientIdStorage = storage({
      load: vi.fn().mockResolvedValue(STORED_CLIENT_ID),
      clear: vi.fn(() => pendingClear.promise)
    });
    const controller = renderHook(() => useGoogleClientIdConfiguration({
      storage: clientIdStorage,
      ignoreLoadFailure: false,
      onError: vi.fn()
    }));
    await waitFor(() => expect(controller.result.current.editableClientId?.source).toBe("stored"));

    let forgetting!: Promise<{ status: "cleared" | "retained" }>;
    act(() => {
      forgetting = controller.result.current.forgetClientId();
    });
    expect(controller.result.current.storageBusy).toBe(true);

    let outcome!: { status: "cleared" | "retained" };
    await act(async () => {
      pendingClear.reject(new Error("private storage response"));
      outcome = await forgetting;
    });

    expect(outcome).toEqual({ status: "retained" });
    expect(controller.result.current.editableClientId?.value).toBe(STORED_CLIENT_ID);
    expect(controller.result.current.warning?.message).not.toMatch(/private|response/i);
  });

  it("serializes accepted saves and forgets through its internal queue", async () => {
    const pendingSave = deferred<void>();
    const pendingClear = deferred<void>();
    const calls: string[] = [];
    const clientIdStorage = storage({
      save: vi.fn((clientId: string) => {
        calls.push(`save:${clientId}`);
        return pendingSave.promise;
      }),
      clear: vi.fn(() => {
        calls.push("clear");
        return pendingClear.promise;
      })
    });
    const controller = renderHook(() => useGoogleClientIdConfiguration({
      storage: clientIdStorage,
      ignoreLoadFailure: false,
      onError: vi.fn()
    }));
    await waitFor(() => expect(controller.result.current.storageStatus).toBe("loaded"));

    let saving!: Promise<void>;
    act(() => {
      saving = controller.result.current.saveClientId(CLIENT_ID);
    });
    let forgetting!: Promise<{ status: "cleared" | "retained" }>;
    act(() => {
      forgetting = controller.result.current.forgetClientId();
    });
    expect(calls).toEqual([`save:${CLIENT_ID}`]);

    await act(async () => pendingSave.resolve(undefined));
    await waitFor(() => expect(calls).toEqual([`save:${CLIENT_ID}`, "clear"]));
    await act(async () => pendingClear.resolve(undefined));
    await act(async () => Promise.all([saving, forgetting]));

    expect(controller.result.current.storageBusy).toBe(false);
    expect(controller.result.current.editableClientId).toBeNull();
  });
});
