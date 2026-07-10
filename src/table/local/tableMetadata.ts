export function createTableMetadataKey(rowId: string, columnId: string): string {
  return JSON.stringify([rowId, columnId]);
}

export function parseTableMetadataKey(key: string): { rowId: string; columnId: string } {
  const parsed: unknown = JSON.parse(key);
  if (!Array.isArray(parsed) || parsed.length !== 2 || parsed.some((value) => typeof value !== "string")) {
    throw new Error("Invalid table metadata key");
  }
  return { rowId: parsed[0], columnId: parsed[1] };
}
