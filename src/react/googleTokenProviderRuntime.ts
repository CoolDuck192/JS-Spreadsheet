import type { TokenProvider } from "../core/workbook/services";
import type { GoogleAuthSource } from "../lib/googleConfiguration";

type ResolvableGoogleAuthSource = Exclude<GoogleAuthSource, { kind: "missing" }>;

export type GoogleTokenProviderPreparation = Readonly<{
  provider: TokenProvider;
  prepare(): Promise<void>;
}>;

export type GoogleTokenProviderRuntime = Readonly<{
  resolve(source: ResolvableGoogleAuthSource): GoogleTokenProviderPreparation;
  clear(): void;
}>;

type ProviderRecord = {
  kind: "direct" | "factory";
  identity: TokenProvider | ((clientId: string) => TokenProvider);
  clientId: string;
  provider: TokenProvider;
  prepared: boolean;
  preparation?: Promise<void>;
};

export function createGoogleTokenProviderRuntime(): GoogleTokenProviderRuntime {
  let records: ProviderRecord[] = [];

  return {
    resolve(source) {
      const record = resolveProviderRecord(records, source);
      return {
        provider: record.provider,
        prepare: () => prepareProvider(record)
      };
    },
    clear() {
      records = [];
    }
  };
}

function resolveProviderRecord(
  records: ProviderRecord[],
  source: ResolvableGoogleAuthSource
): ProviderRecord {
  const kind = source.kind === "provider" ? "direct" : "factory";
  const identity = source.kind === "provider" ? source.provider : source.factory;
  const clientId = source.kind === "provider" ? "" : source.clientId;
  const cached = records.find(
    (record) =>
      record.kind === kind &&
      record.identity === identity &&
      record.clientId === clientId
  );
  if (cached) {
    return cached;
  }

  const provider = source.kind === "provider"
    ? source.provider
    : source.factory(source.clientId);
  if (!provider || typeof provider.getAccessToken !== "function") {
    throw new Error("Invalid Google token provider.");
  }
  const record: ProviderRecord = {
    kind,
    identity,
    clientId,
    provider,
    prepared: false
  };
  records.splice(0, records.length, record);
  return record;
}

function prepareProvider(record: ProviderRecord): Promise<void> {
  if (record.prepared) {
    return Promise.resolve();
  }
  if (record.preparation) {
    return record.preparation;
  }
  if (!record.provider.prepare) {
    record.prepared = true;
    return Promise.resolve();
  }

  try {
    record.preparation = Promise.resolve(record.provider.prepare()).then(
      () => {
        record.prepared = true;
        record.preparation = undefined;
      },
      (error: unknown) => {
        record.preparation = undefined;
        throw error;
      }
    );
  } catch (error) {
    return Promise.reject(error);
  }
  return record.preparation;
}
