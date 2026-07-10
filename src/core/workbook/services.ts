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

/** Platform-neutral token source. Browser OAuth implementations live in src/react. */
export interface TokenProvider {
  getAccessToken(scopes: readonly string[]): Promise<string>;
}

export type SpreadsheetServices = Readonly<{
  formulaEngineFactory?: typeof createFormulaEngine;
  importers?: Readonly<Record<string, WorkbookImporter>>;
  exporters?: Readonly<Record<string, WorkbookExporter>>;
  now?: () => number;
  createCommandId?: () => string;
  createId?: IdGenerator;
  googleTokenProviderFactory?: (clientId: string) => TokenProvider;
}>;
