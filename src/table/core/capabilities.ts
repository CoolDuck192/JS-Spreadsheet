import type { PaginationRequest } from "./query";

export type OperationCapability = {
  executor: "client" | "server";
  scope: "loadedRows" | "completeDataset";
};

export type FormulaCapability = "none" | "loadedRows" | "fullLocalDataset" | "server";
export type TableFeature =
  | "sort"
  | "filter"
  | "group"
  | "aggregate"
  | "pagination"
  | "edit"
  | "bulkEdit"
  | "metadata"
  | "validation"
  | "formula"
  | "subscription"
  | "undo"
  | "export";
export type TableFeatureConfiguration = Partial<Record<
  TableFeature,
  boolean | { enabled?: boolean; requiredScope?: "loadedRows" | "completeDataset" }
>>;
export type TableOperationState = {
  enabled: boolean;
  scopeLabel?: "Loaded rows" | "Complete dataset";
  reason?: string;
};

export type TableCapabilities = {
  sort: OperationCapability | false;
  filter: OperationCapability | false;
  group: OperationCapability | false;
  aggregate: OperationCapability | false;
  pagination: (OperationCapability & { modes: readonly PaginationRequest["kind"][] }) | false;
  edit: OperationCapability | false;
  bulkEdit: OperationCapability | false;
  metadata: OperationCapability | false;
  validation: OperationCapability | false;
  formula: FormulaCapability;
  subscription: boolean;
  undo: OperationCapability | false;
  export: OperationCapability | false;
};

const FEATURES: readonly TableFeature[] = [
  "sort",
  "filter",
  "group",
  "aggregate",
  "pagination",
  "edit",
  "bulkEdit",
  "metadata",
  "validation",
  "formula",
  "subscription",
  "undo",
  "export"
];

const COMPLETE_DATASET: OperationCapability = { executor: "client", scope: "completeDataset" };

export function createLocalTableCapabilities(
  options: { formula?: FormulaCapability; undo?: boolean } = {}
): TableCapabilities {
  return {
    sort: { ...COMPLETE_DATASET },
    filter: { ...COMPLETE_DATASET },
    group: { ...COMPLETE_DATASET },
    aggregate: { ...COMPLETE_DATASET },
    pagination: { ...COMPLETE_DATASET, modes: ["none", "offset"] },
    edit: { ...COMPLETE_DATASET },
    bulkEdit: { ...COMPLETE_DATASET },
    metadata: { ...COMPLETE_DATASET },
    validation: { ...COMPLETE_DATASET },
    formula: options.formula ?? "none",
    subscription: false,
    undo: options.undo === false ? false : { ...COMPLETE_DATASET },
    export: { ...COMPLETE_DATASET }
  };
}

export function resolveTableOperationStates(
  capabilities: TableCapabilities,
  configuration: TableFeatureConfiguration
): Readonly<Record<TableFeature, TableOperationState>> {
  return Object.fromEntries(FEATURES.map((feature) => [
    feature,
    resolveFeature(feature, capabilities, configuration[feature])
  ])) as Record<TableFeature, TableOperationState>;
}

function resolveFeature(
  feature: TableFeature,
  capabilities: TableCapabilities,
  configuration: TableFeatureConfiguration[TableFeature]
): TableOperationState {
  if (configuration === false || (typeof configuration === "object" && configuration.enabled === false)) {
    return { enabled: false, reason: "Disabled by host configuration" };
  }

  const capability = capabilities[feature];
  if (feature === "formula") {
    return resolveFormula(capabilities.formula, configuration);
  }
  if (feature === "subscription") {
    return capability
      ? { enabled: true }
      : { enabled: false, reason: "Subscriptions are not supported by this table" };
  }
  if (!capability) {
    return { enabled: false, reason: `${featureLabel(feature)} is not supported by this table` };
  }

  const scope = (capability as OperationCapability).scope;
  const scopeLabel = scope === "completeDataset" ? "Complete dataset" : "Loaded rows";
  const requiredScope = typeof configuration === "object" ? configuration.requiredScope : undefined;
  if (requiredScope === "completeDataset" && scope === "loadedRows") {
    return {
      enabled: false,
      scopeLabel,
      reason: `This table only supports ${featureLabel(feature).toLowerCase()} loaded rows.`
    };
  }
  return { enabled: true, scopeLabel };
}

function resolveFormula(
  capability: FormulaCapability,
  configuration: TableFeatureConfiguration[TableFeature]
): TableOperationState {
  if (capability === "none") {
    return { enabled: false, reason: "Formula service is not configured" };
  }
  const scope = capability === "loadedRows" ? "loadedRows" : "completeDataset";
  const scopeLabel = scope === "loadedRows" ? "Loaded rows" : "Complete dataset";
  const requiredScope = typeof configuration === "object" ? configuration.requiredScope : undefined;
  if (requiredScope === "completeDataset" && scope === "loadedRows") {
    return {
      enabled: false,
      scopeLabel,
      reason: "This table only supports formulas for loaded rows."
    };
  }
  return { enabled: true, scopeLabel };
}

function featureLabel(feature: TableFeature): string {
  const labels: Record<TableFeature, string> = {
    sort: "Sorting",
    filter: "Filtering",
    group: "Grouping",
    aggregate: "Aggregation",
    pagination: "Pagination",
    edit: "Editing",
    bulkEdit: "Bulk editing",
    metadata: "Metadata",
    validation: "Validation",
    formula: "Formulas",
    subscription: "Subscriptions",
    undo: "Undo",
    export: "Export"
  };
  return labels[feature];
}
