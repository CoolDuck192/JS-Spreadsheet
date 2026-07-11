import { describe, expect, it, vi } from "vitest";
import type { TokenProvider } from "../core/workbook/services";
import { createGoogleTokenProviderRuntime } from "./googleTokenProviderRuntime";

const CLIENT_ID = "123-abc.apps.googleusercontent.com";

function provider(overrides: Partial<TokenProvider> = {}): TokenProvider {
  return {
    prepare: vi.fn().mockResolvedValue(undefined),
    getAccessToken: vi.fn().mockResolvedValue("test-token"),
    ...overrides
  };
}

describe("Google token provider runtime", () => {
  it("caches factory providers and preparation by identity until cleared", async () => {
    const first = provider();
    const second = provider();
    const factory = vi.fn()
      .mockReturnValueOnce(first)
      .mockReturnValueOnce(second);
    const source = {
      kind: "client" as const,
      clientId: CLIENT_ID,
      source: "stored" as const,
      factory
    };
    const runtime = createGoogleTokenProviderRuntime();

    const initial = runtime.resolve(source);
    await initial.prepare();
    await runtime.resolve(source).prepare();

    expect(initial.provider).toBe(first);
    expect(factory).toHaveBeenCalledTimes(1);
    expect(first.prepare).toHaveBeenCalledTimes(1);

    runtime.clear();
    const afterClear = runtime.resolve(source);
    await afterClear.prepare();
    expect(afterClear.provider).toBe(second);
    expect(factory).toHaveBeenCalledTimes(2);
  });

  it("retries failed preparation and does not cache a throwing factory", async () => {
    const preparation = vi.fn()
      .mockRejectedValueOnce(new Error("first preparation"))
      .mockResolvedValueOnce(undefined);
    const prepared = provider({ prepare: preparation });
    const factory = vi.fn()
      .mockImplementationOnce(() => {
        throw new Error("first construction");
      })
      .mockReturnValue(prepared);
    const source = {
      kind: "client" as const,
      clientId: CLIENT_ID,
      source: "stored" as const,
      factory
    };
    const runtime = createGoogleTokenProviderRuntime();

    expect(() => runtime.resolve(source)).toThrow("first construction");
    const resolved = runtime.resolve(source);
    await expect(resolved.prepare()).rejects.toThrow("first preparation");
    await expect(runtime.resolve(source).prepare()).resolves.toBeUndefined();

    expect(factory).toHaveBeenCalledTimes(2);
    expect(preparation).toHaveBeenCalledTimes(2);
  });
});
