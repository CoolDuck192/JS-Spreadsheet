import { afterEach, describe, expect, it, vi } from "vitest";
import { createCommandIdFactory } from "./commandId";

describe("createCommandIdFactory", () => {
  afterEach(() => {
    vi.unstubAllGlobals();
  });

  it("combines a collision-resistant session id with a monotonic local sequence", () => {
    const next = createCommandIdFactory(() => "session-uuid");
    expect(next()).toBe("session-uuid:1");
    expect(next()).toBe("session-uuid:2");
  });

  it("keeps independent sessions distinct", () => {
    const first = createCommandIdFactory(() => "first-uuid");
    const second = createCommandIdFactory(() => "second-uuid");
    expect(first()).not.toBe(second());
  });

  it("uses secure random bytes when randomUUID is unavailable", () => {
    let seed = 0;
    vi.stubGlobal("crypto", {
      getRandomValues<T extends ArrayBufferView | null>(array: T): T {
        if (!array) return array;
        const bytes = new Uint8Array(array.buffer, array.byteOffset, array.byteLength);
        for (let index = 0; index < bytes.length; index += 1) {
          bytes[index] = (seed + index) & 0xff;
        }
        seed += bytes.length;
        return array;
      }
    });

    const first = createCommandIdFactory();
    const second = createCommandIdFactory();

    expect(first()).toBe("00010203-0405-4607-8809-0a0b0c0d0e0f:1");
    expect(first()).toBe("00010203-0405-4607-8809-0a0b0c0d0e0f:2");
    expect(second()).toBe("10111213-1415-4617-9819-1a1b1c1d1e1f:1");
  });
});
