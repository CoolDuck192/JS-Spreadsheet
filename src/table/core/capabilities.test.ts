import { describe, expect, it } from "vitest";
import { createLocalTableCapabilities, resolveTableOperationStates } from "./capabilities";

describe("table capability resolution", () => {
  it("advertises local operations as client-owned complete-dataset work", () => {
    const states = resolveTableOperationStates(createLocalTableCapabilities(), {});
    expect(states.sort).toEqual({ enabled: true, scopeLabel: "Complete dataset" });
    expect(states.filter).toEqual({ enabled: true, scopeLabel: "Complete dataset" });
    expect(states.formula).toEqual({ enabled: false, reason: "Formula service is not configured" });
  });

  it("intersects source support with host feature configuration", () => {
    const states = resolveTableOperationStates(createLocalTableCapabilities({ formula: "fullLocalDataset" }), {
      sort: false,
      formula: true
    });
    expect(states.sort).toEqual({ enabled: false, reason: "Disabled by host configuration" });
    expect(states.formula).toEqual({ enabled: true, scopeLabel: "Complete dataset" });
  });

  it("rejects a complete-dataset requirement for loaded-row support", () => {
    const capabilities = createLocalTableCapabilities();
    capabilities.sort = { executor: "client", scope: "loadedRows" };
    const states = resolveTableOperationStates(capabilities, {
      sort: { requiredScope: "completeDataset" }
    });
    expect(states.sort).toEqual({
      enabled: false,
      scopeLabel: "Loaded rows",
      reason: "This table only supports sorting loaded rows."
    });
  });
});
