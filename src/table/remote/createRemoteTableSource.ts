import type { QueryRequest, QueryResult } from "../core/query";
import type {
  CreateRemoteTableSourceOptions,
  RemoteTableSource
} from "./types";

export function createRemoteTableSource<TRow>(
  options: CreateRemoteTableSourceOptions<TRow>
): RemoteTableSource<TRow> {
  validateDeclaration(options);

  return {
    ...options,
    kind: "remote",
    async query(request, context) {
      validateRequestPagination(request, options.paginationMode);
      const result = await options.query(request, context);
      validateResult(result, options);
      return result;
    }
  };
}

function validateDeclaration<TRow>(options: CreateRemoteTableSourceOptions<TRow>): void {
  const pagination = options.capabilities.pagination;
  if (options.paginationMode === "none") {
    if (pagination && !pagination.modes.includes("none")) {
      invalid("paginationMode none is absent from capabilities.pagination.modes");
    }
  } else if (!pagination || !pagination.modes.includes(options.paginationMode)) {
    invalid(`paginationMode ${options.paginationMode} is absent from capabilities.pagination.modes`);
  }

  const editEnabled = Boolean(options.capabilities.edit);
  const bulkEditEnabled = Boolean(options.capabilities.bulkEdit);
  const metadataEnabled = Boolean(options.capabilities.metadata);
  const serverValidation = options.capabilities.validation !== false
    && options.capabilities.validation.executor === "server";
  const serverFormula = options.capabilities.formula === "server";
  const requiresMutation = editEnabled || bulkEditEnabled || metadataEnabled || serverValidation || serverFormula;

  if (requiresMutation && !options.mutate) {
    invalid("Declared write capabilities require mutate");
  }
  if (requiresMutation && options.mutationMode === "none") {
    invalid("Declared write capabilities require versioned mutationMode");
  }
  if (metadataEnabled && !options.readCell) {
    invalid("Metadata capability requires readCell");
  }
  if (serverFormula && !options.readCell) {
    invalid("Server formula capability requires readCell");
  }
  if (options.capabilities.export && !options.export) {
    invalid("Export capability requires export");
  }
  if (options.capabilities.subscription && !options.subscribe) {
    invalid("Subscription capability requires subscribe");
  }

  if (options.undoMode === "compensating") {
    if (
      options.mutationMode !== "versioned"
      || !options.mutate
      || !options.capabilities.undo
    ) {
      invalid("Compensating undo requires versioned mutations, mutate, and undo capability");
    }
  } else if (options.capabilities.undo) {
    invalid("Undo capability requires compensating undoMode");
  }
}

function validateRequestPagination(
  request: QueryRequest,
  mode: CreateRemoteTableSourceOptions<unknown>["paginationMode"]
): void {
  if (request.pagination.kind !== mode) {
    invalid(`Query pagination kind ${request.pagination.kind} does not match source mode ${mode}`);
  }
}

function validateResult<TRow>(
  result: QueryResult<TRow>,
  options: CreateRemoteTableSourceOptions<TRow>
): void {
  if (result.pageInfo.kind !== options.paginationMode) {
    invalid(`Result page kind ${result.pageInfo.kind} does not match source mode ${options.paginationMode}`);
  }
  if (options.paginationMode !== "none" && result.completeness === "completeDataset") {
    invalid("Paginated results cannot claim completeDataset completeness");
  }

  const ids = new Set<string>();
  for (const item of result.items) {
    if (!item.id.trim()) invalid("Remote query rows require non-blank IDs");
    if (ids.has(item.id)) invalid(`Remote query returned duplicate row ID: ${item.id}`);
    ids.add(item.id);
    if (item.kind !== "data") continue;
    const sourceId = options.getRowId(item.original);
    if (!sourceId.trim()) invalid("Remote source getRowId returned a blank ID");
    if (sourceId !== item.id) {
      invalid(`Remote query row ID ${item.id} does not match source row ID ${sourceId}`);
    }
  }
}

function invalid(message: string): never {
  throw new Error(`Invalid remote table source: ${message}`);
}
