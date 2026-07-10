import type { TableCellMetadata } from "../core/types";

export type RemoteJournalRowVersion = {
  rowId: string;
  columnId: string;
  rowVersion?: string;
};

export type RemoteOperationJournalChange = {
  rowId: string;
  columnId: string;
} & (
  | { kind: "value"; originalValue: unknown; committedValue: unknown }
  | {
      kind: "metadata";
      originalMetadata: TableCellMetadata;
      committedMetadata: TableCellMetadata;
    }
);

export type RemoteOperationJournalEntry = {
  operationId: string;
  acknowledgedRevision: string;
  rowVersions: readonly RemoteJournalRowVersion[];
  changes: readonly RemoteOperationJournalChange[];
};

export type RemoteUndoClaim =
  | {
      kind: "pending";
      entry: { operationId: string; changes: readonly RemoteOperationJournalChange[] };
    }
  | { kind: "committed"; entry: RemoteOperationJournalEntry };

type PendingEntry = {
  operationId: string;
  abortController: AbortController;
  changes: readonly RemoteOperationJournalChange[];
  sequence: number;
};

type AcknowledgedEntry = RemoteOperationJournalEntry & {
  sequence: number;
  compensationPending: boolean;
};

const DEFAULT_MAX_ACKNOWLEDGED = 100;

export class RemoteOperationJournal {
  private readonly pending = new Map<string, PendingEntry>();
  private readonly acknowledged: AcknowledgedEntry[] = [];
  private readonly tombstones = new Set<string>();
  private readonly maxAcknowledged: number;
  private sequence = 0;

  constructor(maxAcknowledged: number = DEFAULT_MAX_ACKNOWLEDGED) {
    if (!Number.isInteger(maxAcknowledged) || maxAcknowledged < 1) {
      throw new RangeError("maxAcknowledged must be a positive integer");
    }
    this.maxAcknowledged = Math.min(maxAcknowledged, DEFAULT_MAX_ACKNOWLEDGED);
  }

  get size(): number {
    return this.acknowledged.length;
  }

  get canUndo(): boolean {
    return this.pending.size > 0 || this.acknowledged.some((entry) => !entry.compensationPending);
  }

  begin(
    operationId: string,
    abortController: AbortController,
    changes: readonly RemoteOperationJournalChange[]
  ): void {
    if (!operationId.trim() || this.pending.has(operationId) || this.findAcknowledged(operationId)) {
      throw new Error("Remote journal operation ID must be unique and non-blank");
    }
    this.pending.set(operationId, {
      operationId,
      abortController,
      changes: cloneChanges(changes),
      sequence: ++this.sequence
    });
  }

  acknowledge(
    operationId: string,
    acknowledgedRevision: string,
    rowVersions: readonly RemoteJournalRowVersion[]
  ): boolean {
    if (this.tombstones.has(operationId)) return false;
    const entry = this.pending.get(operationId);
    if (!entry) return false;
    this.pending.delete(operationId);
    this.acknowledged.push({
      operationId,
      acknowledgedRevision,
      rowVersions: rowVersions.map((rowVersion) => ({ ...rowVersion })),
      changes: cloneChanges(entry.changes),
      sequence: entry.sequence,
      compensationPending: false
    });
    while (this.acknowledged.length > this.maxAcknowledged) this.acknowledged.shift();
    return true;
  }

  discardPending(operationId: string): void {
    this.pending.delete(operationId);
  }

  claimLatestForUndo(): RemoteUndoClaim | null {
    const pending = [...this.pending.values()].sort((left, right) => right.sequence - left.sequence)[0];
    const committed = [...this.acknowledged]
      .filter((entry) => !entry.compensationPending)
      .sort((left, right) => right.sequence - left.sequence)[0];
    if (!pending && !committed) return null;
    if (pending && (!committed || pending.sequence > committed.sequence)) {
      this.pending.delete(pending.operationId);
      this.tombstones.add(pending.operationId);
      pending.abortController.abort();
      return {
        kind: "pending",
        entry: { operationId: pending.operationId, changes: cloneChanges(pending.changes) }
      };
    }
    committed!.compensationPending = true;
    return { kind: "committed", entry: publicEntry(committed!) };
  }

  releaseCompensation(operationId: string): void {
    const entry = this.findAcknowledged(operationId);
    if (entry) entry.compensationPending = false;
  }

  reserveCompensation(operationId: string): boolean {
    const entry = this.findAcknowledged(operationId);
    if (!entry || entry.compensationPending) return false;
    entry.compensationPending = true;
    return true;
  }

  completeCompensation(operationId: string): void {
    const index = this.acknowledged.findIndex((entry) => entry.operationId === operationId);
    if (index >= 0) this.acknowledged.splice(index, 1);
  }

  isTombstoned(operationId: string): boolean {
    return this.tombstones.has(operationId);
  }

  reconcileTombstone(operationId: string): void {
    this.tombstones.delete(operationId);
  }

  clear(): void {
    for (const entry of this.pending.values()) entry.abortController.abort();
    this.pending.clear();
    this.acknowledged.length = 0;
    this.tombstones.clear();
  }

  private findAcknowledged(operationId: string): AcknowledgedEntry | undefined {
    return this.acknowledged.find((entry) => entry.operationId === operationId);
  }
}

function publicEntry(entry: AcknowledgedEntry): RemoteOperationJournalEntry {
  return {
    operationId: entry.operationId,
    acknowledgedRevision: entry.acknowledgedRevision,
    rowVersions: entry.rowVersions.map((rowVersion) => ({ ...rowVersion })),
    changes: cloneChanges(entry.changes)
  };
}

function cloneChanges(changes: readonly RemoteOperationJournalChange[]): RemoteOperationJournalChange[] {
  return changes.map((change) => change.kind === "value"
    ? { ...change }
    : {
        ...change,
        originalMetadata: cloneMetadata(change.originalMetadata),
        committedMetadata: cloneMetadata(change.committedMetadata)
      });
}

function cloneMetadata(metadata: TableCellMetadata): TableCellMetadata {
  return {
    ...metadata,
    ...(metadata.format ? { format: { ...metadata.format } } : {}),
    ...(metadata.validation
      ? {
          validation: metadata.validation.kind === "list"
            ? { ...metadata.validation, values: [...metadata.validation.values] }
            : { ...metadata.validation }
        }
      : {})
  };
}
