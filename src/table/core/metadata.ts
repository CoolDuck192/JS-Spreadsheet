import type { TableCellMetadataUpdate } from "./types";

export function coalesceTableCellMetadataUpdates(
  updates: readonly TableCellMetadataUpdate[]
): readonly TableCellMetadataUpdate[] {
  const byTarget = new Map<string, TableCellMetadataUpdate>();
  const coalesced: TableCellMetadataUpdate[] = [];
  for (const update of updates) {
    const key = `${update.rowId.length}:${update.rowId}${update.columnId}`;
    const existing = byTarget.get(key);
    if (existing) {
      existing.patch = { ...existing.patch, ...update.patch };
      continue;
    }
    const candidate = { ...update, patch: { ...update.patch } };
    byTarget.set(key, candidate);
    coalesced.push(candidate);
  }
  return coalesced;
}
