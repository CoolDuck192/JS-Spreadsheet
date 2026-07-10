import { describe, expect, it } from "vitest";
import { RemoteOperationJournal } from "./RemoteOperationJournal";

describe("RemoteOperationJournal", () => {
  it("claims a pending batch by aborting it and retaining a reconciliation tombstone", () => {
    const journal = new RemoteOperationJournal();
    const abortController = new AbortController();
    journal.begin("operation-1", abortController, [{
      kind: "value",
      rowId: "row-1",
      columnId: "amount",
      originalValue: 10,
      committedValue: 20
    }]);

    const claim = journal.claimLatestForUndo();

    expect(claim).toMatchObject({ kind: "pending", entry: { operationId: "operation-1" } });
    expect(abortController.signal.aborted).toBe(true);
    expect(journal.isTombstoned("operation-1")).toBe(true);
    expect(journal.canUndo).toBe(false);
    expect(journal.size).toBe(0);
  });

  it("retains at most the latest 100 acknowledged batches by default", () => {
    const journal = new RemoteOperationJournal();
    for (let index = 1; index <= 101; index += 1) {
      const operationId = `operation-${index}`;
      journal.begin(operationId, new AbortController(), [{
        kind: "value",
        rowId: "row-1",
        columnId: "amount",
        originalValue: index - 1,
        committedValue: index
      }]);
      journal.acknowledge(operationId, `revision-${index}`, []);
    }

    expect(journal.size).toBe(100);
    expect(journal.claimLatestForUndo()).toMatchObject({
      kind: "committed",
      entry: { operationId: "operation-101" }
    });
  });

  it("bounds acknowledged batches and preserves complete cloned compensation state", () => {
    const journal = new RemoteOperationJournal(2);
    const originalMetadata = { comment: "original", format: { bold: true } };
    for (let index = 1; index <= 3; index += 1) {
      const operationId = `operation-${index}`;
      journal.begin(operationId, new AbortController(), [{
        kind: "metadata",
        rowId: "row-1",
        columnId: "name",
        originalMetadata,
        committedMetadata: { comment: `changed-${index}` }
      }]);
      journal.acknowledge(operationId, `revision-${index}`, [{
        rowId: "row-1",
        columnId: "name",
        rowVersion: `row-version-${index}`
      }]);
    }
    originalMetadata.comment = "mutated later";

    expect(journal.size).toBe(2);
    const claim = journal.claimLatestForUndo();
    expect(claim).toMatchObject({
      kind: "committed",
      entry: {
        operationId: "operation-3",
        acknowledgedRevision: "revision-3",
        rowVersions: [{ rowId: "row-1", columnId: "name", rowVersion: "row-version-3" }],
        changes: [{
          kind: "metadata",
          originalMetadata: { comment: "original", format: { bold: true } },
          committedMetadata: { comment: "changed-3" }
        }]
      }
    });
    expect(journal.canUndo).toBe(true);
  });

  it("keeps a failed compensation retryable and removes it only after acknowledgement", () => {
    const journal = new RemoteOperationJournal();
    journal.begin("operation-1", new AbortController(), [{
      kind: "value",
      rowId: "row-1",
      columnId: "amount",
      originalValue: 10,
      committedValue: 20
    }]);
    journal.acknowledge("operation-1", "revision-2", []);

    expect(journal.claimLatestForUndo()).toMatchObject({ kind: "committed" });
    expect(journal.canUndo).toBe(false);
    journal.releaseCompensation("operation-1");
    expect(journal.canUndo).toBe(true);
    expect(journal.reserveCompensation("operation-1")).toBe(true);
    expect(journal.canUndo).toBe(false);
    journal.completeCompensation("operation-1");
    expect(journal.canUndo).toBe(false);
    expect(journal.size).toBe(0);
  });
});
