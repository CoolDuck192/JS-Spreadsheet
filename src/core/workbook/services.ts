import type { WorkbookModel } from "../../types";
import type { createFormulaEngine } from "../../lib/formulaEngine";
import type { IdGenerator } from "../ids";

export type WorkbookImportPayload =
  | { kind: "bytes"; bytes: Uint8Array; fileName?: string }
  | { kind: "text"; text: string; fileName?: string };

export type WorkbookExportArtifact = Readonly<{
  bytes: Uint8Array;
  mediaType: string;
  fileName: string;
}>;

export type WorkbookServiceOptions = Readonly<
  Record<string, string | number | boolean | null>
>;

export interface WorkbookImporter {
  import(
    payload: WorkbookImportPayload,
    options?: WorkbookServiceOptions
  ): WorkbookModel | Promise<WorkbookModel>;
}

export interface WorkbookExporter {
  export(
    workbook: WorkbookModel,
    options?: WorkbookServiceOptions
  ): WorkbookExportArtifact | Promise<WorkbookExportArtifact>;
}

/** Platform-neutral token source. Browser OAuth implementations live in optional connectors. */
export interface TokenProvider {
  prepare?(): void | Promise<void>;
  getAccessToken(scopes: readonly string[]): Promise<string>;
  invalidateAccessToken?(scopes: readonly string[]): void | Promise<void>;
}

export type GoogleClientIdStorage = Readonly<{
  load(): string | null | Promise<string | null>;
  save(clientId: string): void | Promise<void>;
  clear(): void | Promise<void>;
}>;

export type GoogleSheetsServiceConfiguration = Readonly<{
  clientId?: string;
  tokenProvider?: TokenProvider;
  tokenProviderFactory?: (clientId: string) => TokenProvider;
  clientIdStorage?: GoogleClientIdStorage | false;
}>;

export type SpreadsheetServices = Readonly<{
  formulaEngineFactory?: typeof createFormulaEngine;
  importers?: Readonly<Record<string, WorkbookImporter>>;
  exporters?: Readonly<Record<string, WorkbookExporter>>;
  now?: () => number;
  createCommandId?: () => string;
  createId?: IdGenerator;
  googleSheets?: GoogleSheetsServiceConfiguration;
  /** @deprecated Use googleSheets.tokenProviderFactory. */
  googleTokenProviderFactory?: (clientId: string) => TokenProvider;
}>;
