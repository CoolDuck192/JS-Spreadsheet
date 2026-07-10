import type { TableCellIssue, TableCellMetadata } from "../core/types";

export type OptimisticCell = {
  storedValue: unknown;
  evaluatedValue: unknown;
  displayValue: string;
  formula?: string;
  metadata: TableCellMetadata;
  issues?: readonly TableCellIssue[];
};

export type OptimisticOverlay = {
  clientMutationId: string;
  operationId: string;
  rowId: string;
  columnId: string;
  baseRevision: string;
  rowVersion?: string;
  cell: OptimisticCell;
  status: "pending" | "uncertain" | "conflict";
  sequence: number;
};

export class OptimisticOverlayStore {
  private readonly byMutationId = new Map<string, OptimisticOverlay>();
  private readonly byCell = new Map<string, string[]>();
  private sequence = 0;

  get size(): number {
    return this.byMutationId.size;
  }

  get uniqueCellCount(): number {
    return this.byCell.size;
  }

  add(overlay: Omit<OptimisticOverlay, "sequence">): OptimisticOverlay {
    if (this.byMutationId.has(overlay.clientMutationId)) {
      throw new Error(`Duplicate optimistic mutation ID: ${overlay.clientMutationId}`);
    }
    const value: OptimisticOverlay = { ...overlay, sequence: ++this.sequence };
    this.byMutationId.set(value.clientMutationId, value);
    const key = cellKey(value.rowId, value.columnId);
    this.byCell.set(key, [...(this.byCell.get(key) ?? []), value.clientMutationId]);
    return value;
  }

  get(clientMutationId: string): OptimisticOverlay | undefined {
    return this.byMutationId.get(clientMutationId);
  }

  getLatest(rowId: string, columnId: string): OptimisticOverlay | undefined {
    const ids = this.byCell.get(cellKey(rowId, columnId));
    if (!ids || ids.length === 0) return undefined;
    return this.byMutationId.get(ids[ids.length - 1]);
  }

  updateStatus(
    clientMutationId: string,
    status: OptimisticOverlay["status"]
  ): OptimisticOverlay | undefined {
    const current = this.byMutationId.get(clientMutationId);
    if (!current) return undefined;
    const next = { ...current, status };
    this.byMutationId.set(clientMutationId, next);
    return next;
  }

  remove(clientMutationId: string): OptimisticOverlay | undefined {
    const current = this.byMutationId.get(clientMutationId);
    if (!current) return undefined;
    this.byMutationId.delete(clientMutationId);
    const key = cellKey(current.rowId, current.columnId);
    const remaining = (this.byCell.get(key) ?? []).filter((id) => id !== clientMutationId);
    if (remaining.length === 0) this.byCell.delete(key);
    else this.byCell.set(key, remaining);
    return current;
  }

  values(): readonly OptimisticOverlay[] {
    return [...this.byMutationId.values()].sort((left, right) => left.sequence - right.sequence);
  }

  clear(): void {
    this.byMutationId.clear();
    this.byCell.clear();
  }
}

function cellKey(rowId: string, columnId: string): string {
  return `${rowId.length}:${rowId}${columnId}`;
}
