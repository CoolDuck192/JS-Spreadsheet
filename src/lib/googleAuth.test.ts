import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { GoogleSheetsError } from "./googleErrors";
import { createBrowserTokenProvider, SHEETS_READONLY_SCOPE } from "./googleAuth";

const CLIENT_ID = "123-abc.apps.googleusercontent.com";
const GIS_SCRIPT_URL = "https://accounts.google.com/gsi/client";

type TokenResponse = { access_token?: string; error?: string };
type TokenFailure = { type?: string; message?: string };
type TokenClientConfiguration = {
  callback: (response: TokenResponse) => void;
  error_callback?: (error: TokenFailure) => void;
};

function installGoogleIdentityServices(
  initialize?: (configuration: TokenClientConfiguration) => { requestAccessToken(): void }
) {
  const requestAccessToken = vi.fn();
  const configurations: TokenClientConfiguration[] = [];
  const callbacks: {
    token?: TokenClientConfiguration["callback"];
    failure?: NonNullable<TokenClientConfiguration["error_callback"]>;
  } = {};
  const initTokenClient = vi.fn((configuration: TokenClientConfiguration) => {
    configurations.push(configuration);
    if (initialize) return initialize(configuration);
    callbacks.token = configuration.callback;
    callbacks.failure = configuration.error_callback;
    return { requestAccessToken };
  });

  Object.defineProperty(window, "google", {
    configurable: true,
    writable: true,
    value: { accounts: { oauth2: { initTokenClient } } }
  });

  return { callbacks, configurations, initTokenClient, requestAccessToken };
}

function expectCode(promise: Promise<unknown>, code: string) {
  return expect(promise).rejects.toMatchObject({
    name: "GoogleSheetsError",
    code
  });
}

describe("createBrowserTokenProvider", () => {
  beforeEach(() => {
    Reflect.deleteProperty(window, "google");
    document.head.querySelectorAll(`script[src="${GIS_SCRIPT_URL}"]`).forEach((script) => script.remove());
  });

  afterEach(() => {
    Reflect.deleteProperty(window, "google");
    document.head.querySelectorAll(`script[src="${GIS_SCRIPT_URL}"]`).forEach((script) => script.remove());
  });

  it("prepares GIS before requesting and reuses a valid cached token", async () => {
    const { callbacks, initTokenClient, requestAccessToken } = installGoogleIdentityServices();
    const provider = createBrowserTokenProvider(CLIENT_ID);

    await provider.prepare();
    expect(initTokenClient).not.toHaveBeenCalled();
    expect(requestAccessToken).not.toHaveBeenCalled();

    const tokenPromise = provider.getAccessToken([SHEETS_READONLY_SCOPE]);
    expect(requestAccessToken).toHaveBeenCalledTimes(1);
    callbacks.token?.({ access_token: "token" });
    await expect(tokenPromise).resolves.toBe("token");

    await expect(provider.getAccessToken([SHEETS_READONLY_SCOPE])).resolves.toBe("token");
    expect(requestAccessToken).toHaveBeenCalledTimes(1);
  });

  it("prepares lazily when a caller requests a token directly", async () => {
    const { callbacks, initTokenClient, requestAccessToken } = installGoogleIdentityServices();
    const provider = createBrowserTokenProvider(CLIENT_ID);

    const tokenPromise = provider.getAccessToken([SHEETS_READONLY_SCOPE]);
    await vi.waitFor(() => expect(requestAccessToken).toHaveBeenCalledTimes(1));
    expect(initTokenClient).toHaveBeenCalledTimes(1);

    callbacks.token?.({ access_token: "token" });
    await expect(tokenPromise).resolves.toBe("token");
    expect(requestAccessToken).toHaveBeenCalledTimes(1);
  });

  it("coalesces overlapping requests with the same canonical scopes", async () => {
    const { configurations, initTokenClient, requestAccessToken } = installGoogleIdentityServices();
    const provider = createBrowserTokenProvider(CLIENT_ID);
    await provider.prepare();

    const first = provider.getAccessToken([SHEETS_READONLY_SCOPE, "profile"]);
    const second = provider.getAccessToken(["profile", SHEETS_READONLY_SCOPE]);

    expect(second).toBe(first);
    expect(initTokenClient).toHaveBeenCalledTimes(1);
    expect(requestAccessToken).toHaveBeenCalledTimes(1);

    configurations[0]?.callback({ access_token: "shared-token" });
    await expect(first).resolves.toBe("shared-token");
    await expect(second).resolves.toBe("shared-token");
  });

  it("keeps overlapping requests for different scopes independent", async () => {
    const { configurations, initTokenClient, requestAccessToken } = installGoogleIdentityServices();
    const provider = createBrowserTokenProvider(CLIENT_ID);
    await provider.prepare();

    const readonly = provider.getAccessToken([SHEETS_READONLY_SCOPE]);
    const profile = provider.getAccessToken(["profile"]);

    expect(profile).not.toBe(readonly);
    expect(initTokenClient).toHaveBeenCalledTimes(2);
    expect(requestAccessToken).toHaveBeenCalledTimes(2);

    configurations[1]?.callback({ access_token: "profile-token" });
    configurations[0]?.callback({ access_token: "readonly-token" });
    await expect(readonly).resolves.toBe("readonly-token");
    await expect(profile).resolves.toBe("profile-token");
  });

  it("shares a same-scope rejection and clears it so the request can be retried", async () => {
    const { configurations, initTokenClient, requestAccessToken } = installGoogleIdentityServices();
    const provider = createBrowserTokenProvider(CLIENT_ID);
    await provider.prepare();

    const first = provider.getAccessToken([SHEETS_READONLY_SCOPE]);
    const second = provider.getAccessToken([SHEETS_READONLY_SCOPE]);
    const settled = Promise.allSettled([first, second]);

    expect(second).toBe(first);
    expect(initTokenClient).toHaveBeenCalledTimes(1);
    expect(requestAccessToken).toHaveBeenCalledTimes(1);
    configurations[0]?.error_callback?.({ type: "popup_closed" });
    const results = await settled;
    expect(results).toHaveLength(2);
    for (const result of results) {
      expect(result.status).toBe("rejected");
      expect(result.status === "rejected" ? result.reason : null).toMatchObject({
        name: "GoogleSheetsError",
        code: "popup_closed"
      });
    }

    const retry = provider.getAccessToken([SHEETS_READONLY_SCOPE]);
    expect(initTokenClient).toHaveBeenCalledTimes(2);
    expect(requestAccessToken).toHaveBeenCalledTimes(2);
    configurations[1]?.callback({ access_token: "retry-token" });
    await expect(retry).resolves.toBe("retry-token");
  });

  it.each([
    ["popup_failed_to_open", "popup_blocked"],
    ["popup_closed", "popup_closed"]
  ] as const)("maps %s to %s", async (failureType, code) => {
    const { callbacks } = installGoogleIdentityServices();
    const provider = createBrowserTokenProvider(CLIENT_ID);
    await provider.prepare();

    const tokenPromise = provider.getAccessToken([SHEETS_READONLY_SCOPE]);
    expect(callbacks.failure).toBeTypeOf("function");
    callbacks.failure?.({ type: failureType, message: "sensitive GIS detail" });

    await expectCode(tokenPromise, code);
    await expect(tokenPromise).rejects.not.toThrow(/sensitive GIS detail/i);
  });

  it("maps OAuth access denial to a typed safe error", async () => {
    const { callbacks } = installGoogleIdentityServices();
    const provider = createBrowserTokenProvider(CLIENT_ID);
    await provider.prepare();

    const tokenPromise = provider.getAccessToken([SHEETS_READONLY_SCOPE]);
    callbacks.token?.({ error: "access_denied" });

    await expectCode(tokenPromise, "access_denied");
  });

  it("maps GIS script load failure to a typed safe error", async () => {
    const provider = createBrowserTokenProvider(CLIENT_ID);
    const preparation = provider.prepare();
    const script = document.head.querySelector<HTMLScriptElement>(`script[src="${GIS_SCRIPT_URL}"]`);

    expect(script).not.toBeNull();
    script?.dispatchEvent(new Event("error"));

    await expectCode(preparation, "gis_load_failed");
  });

  it("maps missing GIS initialization after script load to a typed safe error", async () => {
    const provider = createBrowserTokenProvider(CLIENT_ID);
    const preparation = provider.prepare();
    const script = document.head.querySelector<HTMLScriptElement>(`script[src="${GIS_SCRIPT_URL}"]`);

    expect(script).not.toBeNull();
    script?.dispatchEvent(new Event("load"));

    await expectCode(preparation, "gis_load_failed");
  });

  it("maps lazy token client initialization exceptions without surfacing their text", async () => {
    installGoogleIdentityServices(() => {
      throw new Error("sensitive initialization detail");
    });
    const provider = createBrowserTokenProvider(CLIENT_ID);

    const error = await provider.getAccessToken([SHEETS_READONLY_SCOPE]).catch((caught: unknown) => caught);

    expect(error).toBeInstanceOf(GoogleSheetsError);
    expect(error).toMatchObject({ code: "gis_load_failed" });
    expect((error as Error).message).not.toContain("sensitive initialization detail");
  });

  it("rejects an invalid client id before loading GIS", async () => {
    const provider = createBrowserTokenProvider("client-secret-value");

    await expectCode(provider.prepare(), "invalid_client");
    expect(document.head.querySelector(`script[src="${GIS_SCRIPT_URL}"]`)).toBeNull();
  });
});
