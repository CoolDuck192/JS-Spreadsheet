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
  const callbacks: {
    token?: TokenClientConfiguration["callback"];
    failure?: NonNullable<TokenClientConfiguration["error_callback"]>;
  } = {};
  const initTokenClient = vi.fn(
    initialize ??
      ((configuration: TokenClientConfiguration) => {
        callbacks.token = configuration.callback;
        callbacks.failure = configuration.error_callback;
        return { requestAccessToken };
      })
  );

  Object.defineProperty(window, "google", {
    configurable: true,
    writable: true,
    value: { accounts: { oauth2: { initTokenClient } } }
  });

  return { callbacks, initTokenClient, requestAccessToken };
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
