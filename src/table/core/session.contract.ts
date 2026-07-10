import { describe, expect, it } from "vitest";
import type { CommandResult } from "../../core/commands/types";
import type { TableFeature } from "./capabilities";
import type { TableCellRef, TableIntent, TableSession } from "./types";

export type TableSessionContractHarness<TRow> = {
  session: TableSession<TRow>;
  editableCell: TableCellRef;
  validRawText: string;
  featureOperations: Partial<Record<TableFeature, () => Promise<CommandResult>>>;
  conflictResolution?: {
    reload: Extract<TableIntent<TRow>, { type: "reload-authoritative" }>;
    retry: Extract<TableIntent<TRow>, { type: "retry-with-revision" }>;
  };
  cleanup(): void;
};

export function defineTableSessionContract<TRow>(
  name: string,
  createHarness: () => TableSessionContractHarness<TRow>
): void {
  describe(`${name} TableSession contract`, () => {
    it("keeps snapshot identity stable until one publication", async () => {
      const harness = createHarness();
      const before = harness.session.getSnapshot();
      expect(harness.session.getSnapshot()).toBe(before);
      let publications = 0;
      harness.session.subscribe(() => { publications += 1; });
      const result = await harness.session.dispatch({
        type: "edit-cells",
        edits: [{ ...harness.editableCell, rawText: harness.validRawText }]
      });
      if (before.operationStates.edit.enabled) {
        expect(result.status).toBe("committed");
        expect(publications).toBe(1);
        expect(harness.session.getSnapshot()).not.toBe(before);
      } else {
        expect(result).toMatchObject({ status: "rejected", reason: "unsupported" });
        expect(publications).toBe(0);
      }
      harness.cleanup();
    });

    it("matches every advertised feature state to command behavior", async () => {
      const harness = createHarness();
      const snapshot = harness.session.getSnapshot();
      for (const [feature, operation] of Object.entries(harness.featureOperations) as Array<[TableFeature, () => Promise<CommandResult>]>) {
        const result = await operation();
        if (snapshot.operationStates[feature].enabled) {
          expect(result).not.toMatchObject({ status: "rejected", reason: "unsupported" });
        } else {
          expect(result).toMatchObject({ status: "rejected", reason: "unsupported" });
        }
      }
      harness.cleanup();
    });

    for (const kind of ["reload", "retry"] as const) {
      it(`${kind} conflict resolution is supported only by a configured source harness`, async () => {
        const harness = createHarness();
        const configured = harness.conflictResolution;
        const intent: TableIntent<TRow> = configured?.[kind] ?? (kind === "reload"
          ? { type: "reload-authoritative", operationId: "contract-conflict", rowId: "contract-row" }
          : {
              type: "retry-with-revision",
              operationId: "contract-conflict",
              rowId: "contract-row",
              expectedRevision: "contract-revision"
            });
        const result = await harness.session.dispatch(intent);
        if (configured) {
          expect(result).not.toMatchObject({ status: "rejected", reason: "unsupported" });
        } else {
          expect(result).toMatchObject({ status: "rejected", reason: "unsupported" });
        }
        harness.cleanup();
      });
    }
  });
}
