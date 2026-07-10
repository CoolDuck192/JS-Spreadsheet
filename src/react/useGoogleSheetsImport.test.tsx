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
import { useGoogleSheetsImport } from "./useGoogleSheetsImport";

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
  return renderHook(() =>
    useGoogleSheetsImport({
      origin: options.origin ?? "https://sheets.example.com",
      onImported: options.onImported ?? vi.fn(),
      onError: options.onError,
      configuration: options.configuration,
      deprecatedTokenProviderFactory: options.deprecatedTokenProviderFactory
    })
  );
}

async function openReady(controller: ReturnType<typeof renderController>) {
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

    expect(googleMocks.importWorkbook).toHaveBeenCalledWith(SHEET_ID, direct);
    await waitFor(() => expect(onImported).toHaveBeenCalledTimes(1));
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
  });

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
    expect(controller.result.current.error?.message).not.toMatch(/payload|token|secret|response/i);

    act(() => controller.result.current.setClientIdDraft(CLIENT_ID));
    await act(async () => controller.result.current.saveClientId());

    await waitFor(() => expect(controller.result.current.phase).toBe("ready"));
    expect(controller.result.current.clientIdSource).toBe("session");
    expect(controller.result.current.error?.message).not.toMatch(/payload|token|secret|response/i);
    expect(onError).toHaveBeenCalled();
  });

  it("returns to setup and invalidates the provider cache when forgetting a stored ID", async () => {
    const clear = vi.fn().mockRejectedValue(new Error("sensitive clear failure"));
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

    await act(async () => controller.result.current.forgetClientId());

    expect(clear).toHaveBeenCalledTimes(1);
    expect(controller.result.current.phase).toBe("setup");
    expect(controller.result.current.clientIdSource).toBe("missing");
    expect(controller.result.current.error?.message).not.toContain("sensitive clear failure");
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

    act(() => controller.result.current.importSheet());
    expect(googleMocks.importWorkbook).toHaveBeenCalledTimes(2);
    await act(async () => secondAttempt.resolve(result("Retried")));
    await waitFor(() => expect(onImported).toHaveBeenCalledTimes(1));
    expect(controller.result.current.open).toBe(false);
    expect(controller.result.current.sheetDraft).toBe("");
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
});
