import { useLayoutEffect } from "react";
import { act, renderHook, waitFor } from "@testing-library/react";
import { beforeEach, describe, expect, it, vi } from "vitest";
import type {
  GoogleClientIdStorage,
  GoogleSheetsServiceConfiguration,
  TokenProvider
} from "../core/workbook/services";
import { GoogleSheetsError } from "../lib/googleErrors";
import type { GoogleSheetsImportResult } from "../lib/googleSheets";
import { createBlankWorkbook } from "../lib/workbook";
import {
  useGoogleSheetsImport,
  type GoogleSheetsImportController
} from "./useGoogleSheetsImport";

const googleMocks = vi.hoisted(() => ({
  builtInFactory: vi.fn(),
  importWorkbook: vi.fn()
}));

vi.mock("../lib/googleAuth", async () => {
  const actual = await vi.importActual<typeof import("../lib/googleAuth")>("../lib/googleAuth");
  return { ...actual, createBrowserTokenProvider: googleMocks.builtInFactory };
});

vi.mock("../lib/googleSheets", async () => {
  const actual = await vi.importActual<typeof import("../lib/googleSheets")>("../lib/googleSheets");
  return { ...actual, importWorkbookFromGoogleSheets: googleMocks.importWorkbook };
});

const CLIENT_ID = "123-abc.apps.googleusercontent.com";
const SECOND_CLIENT_ID = "456-def.apps.googleusercontent.com";
const STORED_CLIENT_ID = "stored.apps.googleusercontent.com";
const SHEET_ID = "12345678901234567890";

function deferred<T>() {
  let resolve!: (value: T) => void;
  let reject!: (reason?: unknown) => void;
  const promise = new Promise<T>((resolvePromise, rejectPromise) => {
    resolve = resolvePromise;
    reject = rejectPromise;
  });
  return { promise, resolve, reject };
}

function provider(overrides: Partial<TokenProvider> = {}): TokenProvider {
  return {
    prepare: vi.fn().mockResolvedValue(undefined),
    getAccessToken: vi.fn().mockResolvedValue("test-token"),
    ...overrides
  };
}

function result(title = "Imported workbook"): GoogleSheetsImportResult {
  return { workbook: createBlankWorkbook(), spreadsheetTitle: title };
}

function storage(overrides: Partial<GoogleClientIdStorage> = {}): GoogleClientIdStorage {
  return {
    load: vi.fn().mockResolvedValue(null),
    save: vi.fn().mockResolvedValue(undefined),
    clear: vi.fn().mockResolvedValue(undefined),
    ...overrides
  };
}

function renderController(options: {
  configuration?: GoogleSheetsServiceConfiguration;
  deprecatedTokenProviderFactory?: (clientId: string) => TokenProvider;
  origin?: string;
  onImported?: (imported: GoogleSheetsImportResult) => void | Promise<void>;
  onError?: (error: GoogleSheetsError) => void;
}) {
  const onImported = options.onImported ?? vi.fn();
  return renderHook(() =>
    useGoogleSheetsImport({
      origin: options.origin ?? "https://sheets.example.com",
      onImported,
      onError: options.onError,
      configuration: options.configuration,
      deprecatedTokenProviderFactory: options.deprecatedTokenProviderFactory
    })
  );
}

type DynamicControllerOptions = {
  configuration?: GoogleSheetsServiceConfiguration;
  deprecatedTokenProviderFactory?: (clientId: string) => TokenProvider;
  origin?: string;
  onImported: (imported: GoogleSheetsImportResult) => void | Promise<void>;
  onError?: (error: GoogleSheetsError) => void;
};

function renderDynamicController(options: DynamicControllerOptions) {
  return renderHook((current: DynamicControllerOptions) =>
    useGoogleSheetsImport({
      origin: current.origin ?? "https://sheets.example.com",
      onImported: current.onImported,
      onError: current.onError,
      configuration: current.configuration,
      deprecatedTokenProviderFactory: current.deprecatedTokenProviderFactory
    }), { initialProps: options });
}

async function openReady(controller: { result: { current: GoogleSheetsImportController } }) {
  act(() => controller.result.current.openDialog());
  await waitFor(() => expect(controller.result.current.phase).toBe("ready"));
}

describe("useGoogleSheetsImport", () => {
  beforeEach(() => {
    googleMocks.builtInFactory.mockReset();
    googleMocks.importWorkbook.mockReset();
  });

  it("lets a direct provider bypass built-in client and LAN-origin setup", async () => {
    const direct = provider();
    const invalidateAccessToken = vi.fn();
    Object.assign(direct, { invalidateAccessToken });
    const clientIdStorage = storage();
    googleMocks.importWorkbook.mockResolvedValue(result());
    const onImported = vi.fn();
    const controller = renderController({
      configuration: { tokenProvider: direct, clientIdStorage },
      origin: "http://192.168.6.232:4173",
      onImported
    });

    await openReady(controller);
    expect(controller.result.current.clientIdSource).toBe("missing");
    expect(direct.prepare).toHaveBeenCalledTimes(1);
    expect(googleMocks.builtInFactory).not.toHaveBeenCalled();

    act(() => controller.result.current.setSheetDraft(SHEET_ID));
    act(() => controller.result.current.importSheet());

    expect(googleMocks.importWorkbook).toHaveBeenCalledWith(
      SHEET_ID,
      expect.objectContaining({ getAccessToken: expect.any(Function) })
    );
    expect(googleMocks.importWorkbook.mock.calls[0][1]).not.toBe(direct);
    const trackedProvider = googleMocks.importWorkbook.mock.calls[0][1] as TokenProvider & {
      invalidateAccessToken?(scopes: readonly string[]): void | Promise<void>;
    };
    expect(trackedProvider.invalidateAccessToken).toBeTypeOf("function");
    await trackedProvider.invalidateAccessToken?.(["scope"]);
    expect(invalidateAccessToken).toHaveBeenCalledWith(["scope"]);
    await waitFor(() => expect(onImported).toHaveBeenCalledTimes(1));
  });

  it("classifies a fresh raw-LAN instance as blocked built-in auth", async () => {
    const controller = renderController({
      origin: "http://192.168.6.232:4173"
    });

    act(() => controller.result.current.openDialog());

    await waitFor(() => expect(controller.result.current.phase).toBe("blocked"));
    expect(controller.result.current.clientIdSource).toBe("missing");
    expect(controller.result.current.hostAuthentication).toBe(false);
    expect(controller.result.current.clientIdEditable).toBe(false);
    expect(controller.result.current.error).toMatchObject({ code: "incompatible_origin" });
    expect(googleMocks.builtInFactory).not.toHaveBeenCalled();
  });

  it("prefers a managed ID over stored state and the nested factory over the deprecated one", async () => {
    const nestedProvider = provider();
    const nestedFactory = vi.fn(() => nestedProvider);
    const deprecatedFactory = vi.fn(() => provider());
    const clientIdStorage = storage({ load: vi.fn().mockResolvedValue(STORED_CLIENT_ID) });
    const controller = renderController({
      configuration: {
        clientId: ` ${CLIENT_ID} `,
        clientIdStorage,
        tokenProviderFactory: nestedFactory
      },
      deprecatedTokenProviderFactory: deprecatedFactory
    });

    await openReady(controller);

    expect(clientIdStorage.load).toHaveBeenCalledTimes(1);
    expect(controller.result.current.clientIdSource).toBe("managed");
    expect(controller.result.current.clientIdDraft).toBe(CLIENT_ID);
    expect(nestedFactory).toHaveBeenCalledTimes(1);
    expect(nestedFactory).toHaveBeenCalledWith(CLIENT_ID);
    expect(deprecatedFactory).not.toHaveBeenCalled();
    expect(nestedProvider.prepare).toHaveBeenCalledTimes(1);
  });

  it("does not construct or prepare built-in Google auth until the dialog opens", async () => {
    const builtInProvider = provider();
    googleMocks.builtInFactory.mockReturnValue(builtInProvider);
    const controller = renderController({
      configuration: { clientId: CLIENT_ID }
    });

    await act(async () => Promise.resolve());
    expect(controller.result.current.phase).toBe("closed");
    expect(googleMocks.builtInFactory).not.toHaveBeenCalled();
    expect(builtInProvider.prepare).not.toHaveBeenCalled();

    await openReady(controller);
    expect(googleMocks.builtInFactory).toHaveBeenCalledWith(CLIENT_ID);
    expect(builtInProvider.prepare).toHaveBeenCalledTimes(1);
  });

  it("blocks only the built-in provider on a raw LAN IP", async () => {
    const controller = renderController({
      configuration: { clientId: CLIENT_ID },
      origin: "http://192.168.6.232:4173"
    });

    act(() => controller.result.current.openDialog());

    await waitFor(() => expect(controller.result.current.phase).toBe("blocked"));
    expect(controller.result.current.error).toMatchObject({ code: "incompatible_origin" });
    expect(googleMocks.builtInFactory).not.toHaveBeenCalled();
  });

  it("lets a host factory own authentication on a raw LAN IP", async () => {
    const hostProvider = provider();
    googleMocks.builtInFactory.mockReturnValue(hostProvider);
    const controller = renderController({
      configuration: {
        clientId: CLIENT_ID,
        tokenProviderFactory: googleMocks.builtInFactory
      },
      origin: "http://192.168.6.232:4173"
    });

    await openReady(controller);

    expect(googleMocks.builtInFactory).toHaveBeenCalledWith(CLIENT_ID);
    expect(hostProvider.prepare).toHaveBeenCalledTimes(1);
    expect(controller.result.current.originAssessment.status).toBe("blocked");
    expect(controller.result.current.hostAuthentication).toBe(true);
  });

  it.each(["nested", "deprecated"] as const)(
    "lets a %s host factory bypass raw-LAN built-in blocking",
    async (kind) => {
      const hostProvider = provider();
      const hostFactory = vi.fn(() => hostProvider);
      const controller = renderController({
        configuration: {
          clientId: CLIENT_ID,
          ...(kind === "nested" ? { tokenProviderFactory: hostFactory } : {})
        },
        deprecatedTokenProviderFactory: kind === "deprecated" ? hostFactory : undefined,
        origin: "http://192.168.6.232:4173"
      });

      await openReady(controller);
      expect(hostFactory).toHaveBeenCalledWith(CLIENT_ID);
      expect(controller.result.current.hostAuthentication).toBe(true);
    }
  );

  it("reports load and save failures safely while continuing with session state", async () => {
    const sensitiveLoad = new Error("storage payload and token");
    const sensitiveSave = new Error("secret save response");
    const clientIdStorage = storage({
      load: vi.fn().mockRejectedValue(sensitiveLoad),
      save: vi.fn().mockRejectedValue(sensitiveSave)
    });
    const hostFactory = vi.fn(() => provider());
    const onError = vi.fn();
    const controller = renderController({
      configuration: { clientIdStorage, tokenProviderFactory: hostFactory },
      onError
    });

    act(() => controller.result.current.openDialog());
    await waitFor(() => expect(controller.result.current.phase).toBe("setup"));
    expect(controller.result.current.warning?.message).not.toMatch(/payload|token|secret|response/i);

    act(() => controller.result.current.setClientIdDraft(CLIENT_ID));
    await act(async () => controller.result.current.saveClientId());

    await waitFor(() => expect(controller.result.current.phase).toBe("ready"));
    expect(controller.result.current.clientIdSource).toBe("session");
    expect(controller.result.current.warning?.message).not.toMatch(/payload|token|secret|response/i);
    expect(onError).toHaveBeenCalled();
  });

  it("can change and forget a session-only client ID without remounting", async () => {
    const hostFactory = vi.fn(() => provider());
    const controller = renderController({
      configuration: { tokenProviderFactory: hostFactory }
    });

    act(() => controller.result.current.openDialog());
    await waitFor(() => expect(controller.result.current.phase).toBe("setup"));
    act(() => controller.result.current.setClientIdDraft(CLIENT_ID));
    await act(async () => controller.result.current.saveClientId());
    await waitFor(() => expect(controller.result.current.phase).toBe("ready"));

    expect(controller.result.current.clientIdSource).toBe("session");
    expect(controller.result.current.canChangeClientId).toBe(true);
    expect(controller.result.current.canForgetClientId).toBe(true);

    act(() => controller.result.current.changeClientId());
    await waitFor(() => expect(controller.result.current.phase).toBe("setup"));
    expect(controller.result.current.clientIdEditable).toBe(true);
    expect(controller.result.current.clientIdDraft).toBe(CLIENT_ID);

    act(() => controller.result.current.setClientIdDraft(SECOND_CLIENT_ID));
    await act(async () => controller.result.current.saveClientId());
    await waitFor(() => expect(controller.result.current.phase).toBe("ready"));
    expect(controller.result.current.clientIdDraft).toBe(SECOND_CLIENT_ID);

    await act(async () => controller.result.current.forgetClientId());
    await waitFor(() => expect(controller.result.current.phase).toBe("setup"));
    expect(controller.result.current.clientIdSource).toBe("missing");
    expect(controller.result.current.clientIdDraft).toBe("");
    expect(controller.result.current.clientIdEditable).toBe(true);
  });

  it("retains the usable stored ID when clearing storage fails", async () => {
    const pendingClear = deferred<void>();
    const clear = vi.fn(() => pendingClear.promise);
    const clientIdStorage = storage({
      load: vi.fn().mockResolvedValue(STORED_CLIENT_ID),
      clear
    });
    const hostFactory = vi.fn(() => provider());
    const controller = renderController({
      configuration: { clientIdStorage, tokenProviderFactory: hostFactory }
    });

    await openReady(controller);
    expect(controller.result.current.clientIdSource).toBe("stored");

    let forgetting!: Promise<void>;
    act(() => {
      forgetting = controller.result.current.forgetClientId();
    });

    expect(clear).toHaveBeenCalledTimes(1);
    expect(controller.result.current.storageBusy).toBe(true);
    expect(controller.result.current.clientIdSource).toBe("stored");

    await act(async () => {
      pendingClear.reject(new Error("sensitive clear failure"));
      await forgetting;
    });

    expect(controller.result.current.phase).toBe("ready");
    expect(controller.result.current.clientIdSource).toBe("stored");
    expect(controller.result.current.clientIdDraft).toBe(STORED_CLIENT_ID);
    expect(controller.result.current.warning?.message).not.toContain("sensitive clear failure");
  });

  it("serializes save, forget, and newer save so the newest stored ID wins", async () => {
    const firstSave = deferred<void>();
    const pendingClear = deferred<void>();
    const secondSave = deferred<void>();
    const calls: string[] = [];
    const clientIdStorage = storage({
      save: vi.fn((clientId: string) => {
        calls.push(`save:${clientId}`);
        return clientId === CLIENT_ID ? firstSave.promise : secondSave.promise;
      }),
      clear: vi.fn(() => {
        calls.push("clear");
        return pendingClear.promise;
      })
    });
    const controller = renderController({
      configuration: { clientIdStorage, tokenProviderFactory: vi.fn(() => provider()) }
    });
    act(() => controller.result.current.openDialog());
    await waitFor(() => expect(controller.result.current.phase).toBe("setup"));

    act(() => controller.result.current.setClientIdDraft(CLIENT_ID));
    let savingFirst!: Promise<void>;
    act(() => {
      savingFirst = controller.result.current.saveClientId();
    });
    let forgetting!: Promise<void>;
    act(() => {
      forgetting = controller.result.current.forgetClientId();
    });
    act(() => controller.result.current.setClientIdDraft(SECOND_CLIENT_ID));
    let savingSecond!: Promise<void>;
    act(() => {
      savingSecond = controller.result.current.saveClientId();
    });

    expect(calls).toEqual([`save:${CLIENT_ID}`]);
    expect(controller.result.current.storageBusy).toBe(true);
    expect(controller.result.current.canImport).toBe(false);

    await act(async () => firstSave.resolve(undefined));
    await waitFor(() => expect(calls).toEqual([`save:${CLIENT_ID}`, "clear"]));
    await act(async () => pendingClear.resolve(undefined));
    await waitFor(() => expect(calls).toEqual([
      `save:${CLIENT_ID}`,
      "clear",
      `save:${SECOND_CLIENT_ID}`
    ]));
    await act(async () => secondSave.resolve(undefined));
    await act(async () => Promise.all([savingFirst, forgetting, savingSecond]));

    expect(controller.result.current.storageBusy).toBe(false);
    expect(controller.result.current.clientIdSource).toBe("stored");
    expect(controller.result.current.clientIdDraft).toBe(SECOND_CLIENT_ID);
    expect(controller.result.current.warning).toBeUndefined();
  });

  it("keeps a failed save usable and visible after import-generation changes", async () => {
    const pendingSave = deferred<void>();
    const clientIdStorage = storage({ save: vi.fn(() => pendingSave.promise) });
    const hostProvider = provider();
    const onError = vi.fn();
    const controller = renderController({
      configuration: {
        clientIdStorage,
        tokenProviderFactory: vi.fn(() => hostProvider)
      },
      onError
    });
    act(() => controller.result.current.openDialog());
    await waitFor(() => expect(controller.result.current.phase).toBe("setup"));
    act(() => controller.result.current.setClientIdDraft(` ${CLIENT_ID} `));
    let saving!: Promise<void>;
    act(() => {
      saving = controller.result.current.saveClientId();
    });

    expect(controller.result.current.storageBusy).toBe(true);
    expect(controller.result.current.canImport).toBe(false);
    act(() => {
      controller.result.current.setSheetDraft(SHEET_ID);
      controller.result.current.importSheet();
      controller.result.current.closeDialog();
    });
    expect(googleMocks.importWorkbook).not.toHaveBeenCalled();

    await act(async () => {
      pendingSave.reject(new Error("secret storage response body"));
      await saving;
    });
    act(() => controller.result.current.openDialog());
    await waitFor(() => expect(controller.result.current.phase).toBe("ready"));

    expect(controller.result.current.clientIdSource).toBe("session");
    expect(controller.result.current.clientIdDraft).toBe(CLIENT_ID);
    expect(controller.result.current.warning?.message).toMatch(/could not be saved/i);
    expect(controller.result.current.warning?.message).not.toMatch(/secret|response body/i);
    expect(onError).toHaveBeenCalledWith(expect.objectContaining({ code: "unknown" }));
  });

  it("disables stored client-ID actions throughout provider and import work", async () => {
    const preparation = deferred<void>();
    const token = deferred<string>();
    const response = deferred<void>();
    const storedProvider = provider({
      prepare: vi.fn(() => preparation.promise),
      getAccessToken: vi.fn(() => token.promise)
    });
    googleMocks.importWorkbook.mockImplementation(async (_sheet: string, importProvider: TokenProvider) => {
      await importProvider.getAccessToken([]);
      await response.promise;
      return result("Pending actions");
    });
    const controller = renderController({
      configuration: {
        clientIdStorage: storage({ load: vi.fn().mockResolvedValue(STORED_CLIENT_ID) }),
        tokenProviderFactory: vi.fn(() => storedProvider)
      }
    });

    act(() => controller.result.current.openDialog());
    await waitFor(() => expect(controller.result.current.phase).toBe("preparing"));
    expect(controller.result.current.clientIdSource).toBe("stored");
    expect(controller.result.current.canChangeClientId).toBe(false);
    expect(controller.result.current.canForgetClientId).toBe(false);

    await act(async () => preparation.resolve(undefined));
    await waitFor(() => expect(controller.result.current.phase).toBe("ready"));
    expect(controller.result.current.canChangeClientId).toBe(true);
    expect(controller.result.current.canForgetClientId).toBe(true);

    act(() => controller.result.current.setSheetDraft(SHEET_ID));
    act(() => controller.result.current.importSheet());
    expect(controller.result.current.phase).toBe("authorizing");
    expect(controller.result.current.canChangeClientId).toBe(false);
    expect(controller.result.current.canForgetClientId).toBe(false);

    await act(async () => token.resolve("ephemeral-token"));
    await waitFor(() => expect(controller.result.current.phase).toBe("importing"));
    expect(controller.result.current.canChangeClientId).toBe(false);
    expect(controller.result.current.canForgetClientId).toBe(false);

    await act(async () => response.resolve(undefined));
  });

  it("invalidates an active import before a programmatic stored-ID forget waits on storage", async () => {
    const pendingImport = deferred<GoogleSheetsImportResult>();
    const pendingClear = deferred<void>();
    const onImported = vi.fn();
    googleMocks.importWorkbook.mockReturnValue(pendingImport.promise);
    const clientIdStorage = storage({
      load: vi.fn().mockResolvedValue(STORED_CLIENT_ID),
      clear: vi.fn(() => pendingClear.promise)
    });
    const controller = renderController({
      configuration: {
        clientIdStorage,
        tokenProviderFactory: vi.fn(() => provider())
      },
      onImported
    });

    await openReady(controller);
    act(() => controller.result.current.setSheetDraft(SHEET_ID));
    act(() => controller.result.current.importSheet());

    let forgetting!: Promise<void>;
    act(() => {
      forgetting = controller.result.current.forgetClientId();
    });
    expect(clientIdStorage.clear).toHaveBeenCalledTimes(1);

    await act(async () => pendingImport.resolve(result("Too late")));
    expect(onImported).not.toHaveBeenCalled();

    await act(async () => {
      pendingClear.resolve(undefined);
      await forgetting;
    });
    expect(controller.result.current.clientIdSource).toBe("missing");
  });

  it("caches and prepares a factory provider once per hook instance", async () => {
    const firstProvider = provider();
    const secondProvider = provider();
    const factory = vi.fn()
      .mockReturnValueOnce(firstProvider)
      .mockReturnValueOnce(secondProvider);

    const first = renderController({
      configuration: { clientId: CLIENT_ID, tokenProviderFactory: factory }
    });
    await openReady(first);
    first.rerender();
    await waitFor(() => expect(first.result.current.phase).toBe("ready"));
    expect(factory).toHaveBeenCalledTimes(1);
    expect(firstProvider.prepare).toHaveBeenCalledTimes(1);

    first.unmount();
    const second = renderController({
      configuration: { clientId: CLIENT_ID, tokenProviderFactory: factory }
    });
    await openReady(second);
    expect(factory).toHaveBeenCalledTimes(2);
    expect(secondProvider.prepare).toHaveBeenCalledTimes(1);
  });

  it("submits synchronously once, preserves failures, and permits retry", async () => {
    const firstAttempt = deferred<GoogleSheetsImportResult>();
    const secondAttempt = deferred<GoogleSheetsImportResult>();
    googleMocks.importWorkbook
      .mockReturnValueOnce(firstAttempt.promise)
      .mockReturnValueOnce(secondAttempt.promise);
    const onImported = vi.fn();
    const onError = vi.fn();
    const controller = renderController({
      configuration: { tokenProvider: provider() },
      onImported,
      onError
    });
    await openReady(controller);
    act(() => controller.result.current.setSheetDraft(SHEET_ID));

    act(() => {
      controller.result.current.importSheet();
      controller.result.current.importSheet();
    });

    expect(googleMocks.importWorkbook).toHaveBeenCalledTimes(1);
    expect(["authorizing", "importing"]).toContain(controller.result.current.phase);

    await act(async () => {
      firstAttempt.reject(new GoogleSheetsError("popup_closed", "Sign-in was closed.", true));
      await firstAttempt.promise.catch(() => undefined);
    });
    expect(controller.result.current.phase).toBe("error");
    expect(controller.result.current.sheetDraft).toBe(SHEET_ID);
    expect(onError).toHaveBeenCalledWith(expect.objectContaining({ code: "popup_closed" }));

    act(() => controller.result.current.retry());
    expect(googleMocks.importWorkbook).toHaveBeenCalledTimes(2);
    await act(async () => secondAttempt.resolve(result("Retried")));
    await waitFor(() => expect(onImported).toHaveBeenCalledTimes(1));
    expect(controller.result.current.open).toBe(false);
    expect(controller.result.current.sheetDraft).toBe("");
  });

  it("reconstructs a provider when a factory throws and Retry is selected", async () => {
    const prepared = provider();
    const factory = vi.fn()
      .mockImplementationOnce(() => {
        throw new GoogleSheetsError(
          "gis_load_failed",
          "Bearer private-token from https://private.example/response-body",
          true
        );
      })
      .mockReturnValueOnce(prepared);
    const onError = vi.fn();
    const controller = renderController({
      configuration: { clientId: CLIENT_ID, tokenProviderFactory: factory },
      onError
    });

    act(() => controller.result.current.openDialog());
    await waitFor(() => expect(controller.result.current.phase).toBe("error"));
    expect(controller.result.current.canRetry).toBe(true);
    expect(controller.result.current.showSheetInput).toBe(false);
    expect(controller.result.current.error?.message).not.toMatch(/private-token|private\.example|response-body/i);
    expect(onError.mock.calls[0][0].message).not.toMatch(/private-token|private\.example|response-body/i);

    act(() => controller.result.current.retry());

    await waitFor(() => expect(controller.result.current.phase).toBe("ready"));
    expect(factory).toHaveBeenCalledTimes(2);
    expect(prepared.prepare).toHaveBeenCalledTimes(1);
  });

  it("keeps authorizing until token acquisition resolves, then reports importing", async () => {
    const token = deferred<string>();
    const response = deferred<void>();
    const direct = provider({ getAccessToken: vi.fn(() => token.promise) });
    googleMocks.importWorkbook.mockImplementation(async (_sheet: string, importProvider: TokenProvider) => {
      await importProvider.getAccessToken([]);
      await response.promise;
      return result("Tracked phases");
    });
    const onImported = vi.fn();
    const controller = renderController({
      configuration: { tokenProvider: direct },
      onImported
    });
    await openReady(controller);
    act(() => controller.result.current.setSheetDraft(SHEET_ID));

    act(() => controller.result.current.importSheet());
    expect(controller.result.current.phase).toBe("authorizing");
    expect(direct.getAccessToken).toHaveBeenCalledTimes(1);

    await act(async () => token.resolve("ephemeral-token"));
    await waitFor(() => expect(controller.result.current.phase).toBe("importing"));
    await act(async () => response.resolve(undefined));
    await waitFor(() => expect(onImported).toHaveBeenCalledTimes(1));
  });

  it("sanitizes typed provider errors before dialog and host output", async () => {
    const direct = provider({
      getAccessToken: vi.fn().mockRejectedValue(
        new GoogleSheetsError(
          "access_denied",
          "Bearer private-token from https://private.example/response-body",
          false
        )
      )
    });
    googleMocks.importWorkbook.mockImplementation(async (_sheet: string, importProvider: TokenProvider) => {
      await importProvider.getAccessToken([]);
      return result();
    });
    const onError = vi.fn();
    const controller = renderController({ configuration: { tokenProvider: direct }, onError });
    await openReady(controller);
    act(() => controller.result.current.setSheetDraft(SHEET_ID));

    act(() => controller.result.current.importSheet());
    await waitFor(() => expect(controller.result.current.phase).toBe("error"));

    expect(controller.result.current.error).toMatchObject({ code: "access_denied", recoverable: false });
    expect(controller.result.current.error?.message).not.toMatch(/private-token|private\.example|response-body/i);
    expect(onError.mock.calls[0][0].message).not.toMatch(/private-token|private\.example|response-body/i);
  });

  it("invalidates a closed attempt and suppresses its late result", async () => {
    const staleAttempt = deferred<GoogleSheetsImportResult>();
    const currentAttempt = deferred<GoogleSheetsImportResult>();
    googleMocks.importWorkbook
      .mockReturnValueOnce(staleAttempt.promise)
      .mockReturnValueOnce(currentAttempt.promise);
    const onImported = vi.fn();
    const controller = renderController({
      configuration: { tokenProvider: provider() },
      onImported
    });
    await openReady(controller);
    act(() => controller.result.current.setSheetDraft(SHEET_ID));
    act(() => controller.result.current.importSheet());
    act(() => controller.result.current.closeDialog());

    await act(async () => staleAttempt.resolve(result("Stale")));
    expect(onImported).not.toHaveBeenCalled();

    await openReady(controller);
    act(() => controller.result.current.importSheet());
    await act(async () => currentAttempt.resolve(result("Current")));

    await waitFor(() => expect(onImported).toHaveBeenCalledWith(expect.objectContaining({
      spreadsheetTitle: "Current"
    })));
  });

  it("invalidates an active import on unmount", async () => {
    const pending = deferred<GoogleSheetsImportResult>();
    const onImported = vi.fn();
    googleMocks.importWorkbook.mockReturnValue(pending.promise);
    const controller = renderController({
      configuration: { tokenProvider: provider() },
      onImported
    });
    await openReady(controller);
    act(() => controller.result.current.setSheetDraft(SHEET_ID));
    act(() => controller.result.current.importSheet());

    controller.unmount();
    await act(async () => pending.resolve(result("Unmounted")));

    expect(onImported).not.toHaveBeenCalled();
  });

  it("binds an active import to the original target callback and invalidates a target swap", async () => {
    const pending = deferred<GoogleSheetsImportResult>();
    const firstTarget = vi.fn();
    const secondTarget = vi.fn();
    const direct = provider();
    googleMocks.importWorkbook.mockReturnValue(pending.promise);
    const controller = renderDynamicController({
      configuration: { tokenProvider: direct },
      onImported: firstTarget
    });
    await openReady(controller);
    act(() => controller.result.current.setSheetDraft(SHEET_ID));
    act(() => controller.result.current.importSheet());

    controller.rerender({
      configuration: { tokenProvider: direct },
      onImported: secondTarget
    });
    await act(async () => pending.resolve(result("Wrong target")));

    expect(firstTarget).not.toHaveBeenCalled();
    expect(secondTarget).not.toHaveBeenCalled();
  });

  it("invalidates a stale target before a later layout effect can settle the import", async () => {
    const pending = deferred<GoogleSheetsImportResult>();
    const firstTarget = vi.fn();
    const secondTarget = vi.fn();
    const direct = provider();
    googleMocks.importWorkbook.mockReturnValue(pending.promise);
    const controller = renderHook((current: DynamicControllerOptions & { settleInLayout: boolean }) => {
      const importController = useGoogleSheetsImport({
        origin: current.origin ?? "https://sheets.example.com",
        onImported: current.onImported,
        onError: current.onError,
        configuration: current.configuration,
        deprecatedTokenProviderFactory: current.deprecatedTokenProviderFactory
      });
      useLayoutEffect(() => {
        if (current.settleInLayout) {
          pending.resolve(result("Stale layout completion"));
        }
      }, [current.settleInLayout]);
      return importController;
    }, {
      initialProps: {
        configuration: { tokenProvider: direct },
        onImported: firstTarget,
        settleInLayout: false
      }
    });
    await openReady(controller);
    act(() => controller.result.current.setSheetDraft(SHEET_ID));
    act(() => controller.result.current.importSheet());

    controller.rerender({
      configuration: { tokenProvider: direct },
      onImported: secondTarget,
      settleInLayout: true
    });
    await act(async () => {
      await pending.promise;
      await Promise.resolve();
    });

    expect(firstTarget).not.toHaveBeenCalled();
    expect(secondTarget).not.toHaveBeenCalled();
  });

  it("does not start an import with the stale provider during an auth-change layout", async () => {
    const providerA = provider();
    const providerB = provider();
    const onImported = vi.fn();
    googleMocks.importWorkbook.mockImplementation(async (_sheet: string, importProvider: TokenProvider) => {
      await importProvider.getAccessToken(["scope"]);
      return result("Wrong provider");
    });
    const controller = renderHook((current: DynamicControllerOptions & { importInLayout: boolean }) => {
      const importController = useGoogleSheetsImport({
        origin: current.origin ?? "https://sheets.example.com",
        onImported: current.onImported,
        onError: current.onError,
        configuration: current.configuration,
        deprecatedTokenProviderFactory: current.deprecatedTokenProviderFactory
      });
      useLayoutEffect(() => {
        if (current.importInLayout) {
          importController.importSheet();
        }
      }, [current.importInLayout, importController.importSheet]);
      return importController;
    }, {
      initialProps: {
        configuration: { tokenProvider: providerA },
        onImported,
        importInLayout: false
      }
    });
    await openReady(controller);
    act(() => controller.result.current.setSheetDraft(SHEET_ID));

    controller.rerender({
      configuration: { tokenProvider: providerB },
      onImported,
      importInLayout: true
    });
    await act(async () => Promise.resolve());

    expect(googleMocks.importWorkbook).not.toHaveBeenCalled();
    expect(providerA.getAccessToken).not.toHaveBeenCalled();
    expect(providerB.getAccessToken).not.toHaveBeenCalled();
    expect(onImported).not.toHaveBeenCalled();
  });

  it.each([
    "managed client ID",
    "direct provider",
    "nested factory",
    "deprecated factory"
  ] as const)("invalidates an active import when the %s changes", async (sourceKind) => {
    const pending = deferred<GoogleSheetsImportResult>();
    const onImported = vi.fn();
    const firstProvider = provider();
    const secondProvider = provider();
    const firstFactory = vi.fn(() => firstProvider);
    const secondFactory = vi.fn(() => secondProvider);
    const base = sourceKind === "direct provider"
      ? { configuration: { tokenProvider: firstProvider } }
      : sourceKind === "nested factory"
        ? { configuration: { clientId: CLIENT_ID, tokenProviderFactory: firstFactory } }
        : sourceKind === "deprecated factory"
          ? { configuration: { clientId: CLIENT_ID }, deprecatedTokenProviderFactory: firstFactory }
          : { configuration: { clientId: CLIENT_ID, tokenProviderFactory: firstFactory } };
    const changed = sourceKind === "direct provider"
      ? { configuration: { tokenProvider: secondProvider } }
      : sourceKind === "nested factory"
        ? { configuration: { clientId: CLIENT_ID, tokenProviderFactory: secondFactory } }
        : sourceKind === "deprecated factory"
          ? { configuration: { clientId: CLIENT_ID }, deprecatedTokenProviderFactory: secondFactory }
          : { configuration: { clientId: SECOND_CLIENT_ID, tokenProviderFactory: firstFactory } };
    googleMocks.importWorkbook.mockReturnValue(pending.promise);
    const controller = renderDynamicController({ ...base, onImported });
    await openReady(controller);
    act(() => controller.result.current.setSheetDraft(SHEET_ID));
    act(() => controller.result.current.importSheet());

    controller.rerender({ ...changed, onImported });
    await act(async () => pending.resolve(result("Stale auth")));

    expect(onImported).not.toHaveBeenCalled();
  });

  it("fully resets client-ID state when storage is swapped or removed", async () => {
    const hostFactory = vi.fn(() => provider());
    const storageA = storage({ load: vi.fn().mockResolvedValue(STORED_CLIENT_ID) });
    const storageB = storage({ load: vi.fn().mockResolvedValue(null) });
    const onImported = vi.fn();
    const controller = renderDynamicController({
      configuration: { clientIdStorage: storageA, tokenProviderFactory: hostFactory },
      onImported
    });
    await waitFor(() => expect(controller.result.current.clientIdSource).toBe("stored"));

    controller.rerender({
      configuration: { clientIdStorage: storageB, tokenProviderFactory: hostFactory },
      onImported
    });
    await waitFor(() => expect(controller.result.current.phase).toBe("closed"));
    await waitFor(() => expect(controller.result.current.clientIdSource).toBe("missing"));
    expect(controller.result.current.clientIdDraft).toBe("");
    expect(controller.result.current.storageBusy).toBe(false);
    expect(controller.result.current.warning).toBeUndefined();

    controller.rerender({
      configuration: { clientIdStorage: false, tokenProviderFactory: hostFactory },
      onImported
    });
    await waitFor(() => expect(controller.result.current.clientIdSource).toBe("missing"));
    expect(controller.result.current.clientIdDraft).toBe("");
  });

  it("keeps a direct provider ready after the unrelated storage identity changes", async () => {
    const direct = provider();
    const storageA = storage();
    const onImported = vi.fn();
    const controller = renderDynamicController({
      configuration: { tokenProvider: direct, clientIdStorage: storageA },
      onImported
    });
    await openReady(controller);

    controller.rerender({
      configuration: { tokenProvider: direct, clientIdStorage: false },
      onImported
    });

    await waitFor(() => expect(controller.result.current.phase).toBe("ready"));
    expect(direct.prepare).toHaveBeenCalledTimes(1);
  });

  it("ignores a late load from a replaced storage adapter", async () => {
    const lateLoad = deferred<string | null>();
    const hostFactory = vi.fn(() => provider());
    const storageA = storage({ load: vi.fn(() => lateLoad.promise) });
    const storageB = storage({ load: vi.fn().mockResolvedValue(null) });
    const onImported = vi.fn();
    const controller = renderDynamicController({
      configuration: { clientIdStorage: storageA, tokenProviderFactory: hostFactory },
      onImported
    });

    controller.rerender({
      configuration: { clientIdStorage: storageB, tokenProviderFactory: hostFactory },
      onImported
    });
    await act(async () => lateLoad.resolve(STORED_CLIENT_ID));

    await waitFor(() => expect(controller.result.current.clientIdSource).toBe("missing"));
    expect(controller.result.current.clientIdDraft).toBe("");
  });

  it("resets pending storage work and ignores its completion after an adapter swap", async () => {
    const pendingSave = deferred<void>();
    const hostFactory = vi.fn(() => provider());
    const storageA = storage({ save: vi.fn(() => pendingSave.promise) });
    const storageB = storage({ load: vi.fn().mockResolvedValue(null) });
    const onImported = vi.fn();
    const controller = renderDynamicController({
      configuration: { clientIdStorage: storageA, tokenProviderFactory: hostFactory },
      onImported
    });
    await waitFor(() => expect(controller.result.current.clientIdSource).toBe("missing"));
    act(() => controller.result.current.setClientIdDraft(CLIENT_ID));
    let saving!: Promise<void>;
    act(() => {
      saving = controller.result.current.saveClientId();
    });
    expect(controller.result.current.storageBusy).toBe(true);

    controller.rerender({
      configuration: { clientIdStorage: storageB, tokenProviderFactory: hostFactory },
      onImported
    });
    await waitFor(() => expect(controller.result.current.storageBusy).toBe(false));
    expect(controller.result.current.clientIdDraft).toBe("");

    await act(async () => {
      pendingSave.resolve(undefined);
      await saving;
    });
    expect(controller.result.current.clientIdSource).toBe("missing");
    expect(controller.result.current.clientIdDraft).toBe("");
  });

  it("retains only the current provider source while reusing a stable source", async () => {
    const providerA = provider();
    const providerB = provider();
    const factoryA = vi.fn(() => providerA);
    const factoryB = vi.fn(() => providerB);
    const onImported = vi.fn();
    const controller = renderDynamicController({
      configuration: { clientId: CLIENT_ID, tokenProviderFactory: factoryA },
      onImported
    });
    await openReady(controller);

    controller.rerender({
      configuration: { clientId: CLIENT_ID, tokenProviderFactory: factoryA },
      onImported
    });
    await waitFor(() => expect(factoryA).toHaveBeenCalledTimes(1));

    controller.rerender({
      configuration: { clientId: CLIENT_ID, tokenProviderFactory: factoryB },
      onImported
    });
    await waitFor(() => expect(factoryB).toHaveBeenCalledTimes(1));
    controller.rerender({
      configuration: { clientId: CLIENT_ID, tokenProviderFactory: factoryA },
      onImported
    });

    await waitFor(() => expect(factoryA).toHaveBeenCalledTimes(2));
    expect(providerA.prepare).toHaveBeenCalledTimes(2);
  });

  it("drops the current provider when authentication becomes missing", async () => {
    const providerA = provider();
    const factoryA = vi.fn(() => providerA);
    const onImported = vi.fn();
    const controller = renderDynamicController({
      configuration: { clientId: CLIENT_ID, tokenProviderFactory: factoryA },
      onImported
    });
    await openReady(controller);

    controller.rerender({
      configuration: { tokenProviderFactory: factoryA },
      onImported
    });
    await waitFor(() => expect(controller.result.current.clientIdSource).toBe("missing"));
    controller.rerender({
      configuration: { clientId: CLIENT_ID, tokenProviderFactory: factoryA },
      onImported
    });

    await waitFor(() => expect(factoryA).toHaveBeenCalledTimes(2));
    expect(providerA.prepare).toHaveBeenCalledTimes(2);
  });

  it("drops the previous provider before resolving a failing replacement source", async () => {
    const providerA = provider();
    const factoryA = vi.fn(() => providerA);
    const failingFactory = vi.fn((): TokenProvider => {
      throw new Error("replacement factory failed");
    });
    const onImported = vi.fn();
    const controller = renderDynamicController({
      configuration: { clientId: CLIENT_ID, tokenProviderFactory: factoryA },
      onImported
    });
    await openReady(controller);

    controller.rerender({
      configuration: { clientId: CLIENT_ID, tokenProviderFactory: failingFactory },
      onImported
    });
    await waitFor(() => expect(failingFactory).toHaveBeenCalledTimes(1));
    controller.rerender({
      configuration: { clientId: CLIENT_ID, tokenProviderFactory: factoryA },
      onImported
    });

    await waitFor(() => expect(factoryA).toHaveBeenCalledTimes(2));
    expect(providerA.prepare).toHaveBeenCalledTimes(2);
  });
});
