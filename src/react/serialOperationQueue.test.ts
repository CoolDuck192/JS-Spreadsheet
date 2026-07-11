import { describe, expect, it, vi } from "vitest";
import { createSerialOperationQueue } from "./serialOperationQueue";

function deferred<T>() {
  let resolve!: (value: T) => void;
  let reject!: (reason?: unknown) => void;
  const promise = new Promise<T>((resolvePromise, rejectPromise) => {
    resolve = resolvePromise;
    reject = rejectPromise;
  });
  return { promise, resolve, reject };
}

describe("serial operation queue", () => {
  it("runs tasks in issue order and continues after a rejection", async () => {
    const first = deferred<void>();
    const second = deferred<void>();
    const calls: string[] = [];
    const queue = createSerialOperationQueue();
    const firstTask = queue.enqueue(() => {
      calls.push("first");
      return first.promise;
    });
    const secondTask = queue.enqueue(() => {
      calls.push("second");
      return second.promise;
    });

    await Promise.resolve();
    expect(calls).toEqual(["first"]);

    first.reject(new Error("expected first failure"));
    await expect(firstTask).rejects.toThrow("expected first failure");
    await vi.waitFor(() => expect(calls).toEqual(["first", "second"]));

    second.resolve(undefined);
    await expect(secondTask).resolves.toBeUndefined();
  });
});
