import { describe, expect, it, vi } from "vitest";
import type { TokenProvider } from "../core/workbook/services";
import {
  assessGoogleOAuthOrigin,
  resolveGoogleAuthSource,
  validateGoogleClientId
} from "./googleConfiguration";

const CLIENT_ID = "123-abc.apps.googleusercontent.com";

function provider(token: string): TokenProvider {
  return {
    getAccessToken: vi.fn().mockResolvedValue(token)
  };
}

const nestedFactory = vi.fn((clientId: string) => provider(`nested:${clientId}`));
const legacyFactory = vi.fn((clientId: string) => provider(`legacy:${clientId}`));
const builtInFactory = vi.fn((clientId: string) => provider(`built-in:${clientId}`));

describe("Google configuration", () => {
  it.each([
    ["https://sheets.example.com", "eligible"],
    ["https://sheets.example.com:8443", "eligible"],
    ["http://192.168.6.232:4173", "blocked"],
    ["https://192.168.6.232", "blocked"],
    ["https://[2001:db8::1]", "blocked"],
    ["file:///tmp/index.html", "blocked"],
    ["null", "blocked"],
    ["http://127.0.0.1:4173", "eligible"],
    ["http://[::1]:4173", "eligible"]
  ] as const)("assesses %s as %s", (origin, status) => {
    expect(assessGoogleOAuthOrigin(origin).status).toBe(status);
  });

  it("reports stable blocked-origin reasons", () => {
    expect(assessGoogleOAuthOrigin("http://192.168.6.232:4173")).toMatchObject({
      status: "blocked",
      reason: "ip_literal"
    });
    expect(assessGoogleOAuthOrigin("http://sheets.example.com")).toMatchObject({
      status: "blocked",
      reason: "insecure"
    });
    expect(assessGoogleOAuthOrigin("file:///tmp/index.html")).toMatchObject({
      status: "blocked",
      reason: "file"
    });
    expect(assessGoogleOAuthOrigin("null")).toMatchObject({
      status: "blocked",
      reason: "opaque"
    });
  });

  it("normalizes only public web OAuth client ids", () => {
    expect(validateGoogleClientId(" 123-abc.apps.googleusercontent.com ")).toEqual({
      valid: true,
      value: CLIENT_ID
    });
    expect(validateGoogleClientId("secret-value")).toMatchObject({ valid: false });
    expect(validateGoogleClientId("https://123-abc.apps.googleusercontent.com")).toMatchObject({
      valid: false
    });
  });

  it("resolves direct provider, managed id, stored id, then factories", () => {
    const direct = provider("direct");
    expect(
      resolveGoogleAuthSource(
        { clientId: CLIENT_ID, tokenProvider: direct, tokenProviderFactory: nestedFactory },
        { value: "stored.apps.googleusercontent.com", source: "stored" },
        legacyFactory,
        builtInFactory
      )
    ).toMatchObject({ kind: "provider", provider: direct });
    expect(
      resolveGoogleAuthSource(
        { clientId: ` ${CLIENT_ID} `, tokenProviderFactory: nestedFactory },
        { value: "stored.apps.googleusercontent.com", source: "stored" },
        legacyFactory,
        builtInFactory
      )
    ).toMatchObject({
      kind: "client",
      clientId: CLIENT_ID,
      source: "managed",
      factory: nestedFactory
    });
    expect(
      resolveGoogleAuthSource(
        {},
        { value: "stored.apps.googleusercontent.com", source: "stored" },
        legacyFactory,
        builtInFactory
      )
    ).toMatchObject({
      kind: "client",
      clientId: "stored.apps.googleusercontent.com",
      source: "stored",
      factory: legacyFactory
    });
    expect(
      resolveGoogleAuthSource(
        {},
        { value: "session.apps.googleusercontent.com", source: "session" },
        undefined,
        builtInFactory
      )
    ).toMatchObject({
      kind: "client",
      clientId: "session.apps.googleusercontent.com",
      source: "session",
      factory: builtInFactory
    });
    expect(resolveGoogleAuthSource(undefined, null, legacyFactory, builtInFactory)).toEqual({
      kind: "missing"
    });
  });
});
