import type { CellContent } from "../types";

export type IdKind = "table" | "table-column" | "table-row";
export type IdGenerator = (kind: IdKind) => string;

export function createRandomId(kind: IdKind): string {
  const crypto = globalThis.crypto;
  if (!crypto) {
    throw new Error("Secure random ID generation is unavailable");
  }
  if (typeof crypto.randomUUID === "function") {
    return `${kind}-${crypto.randomUUID()}`;
  }
  const bytes = new Uint8Array(16);
  crypto.getRandomValues(bytes);
  bytes[6] = (bytes[6] & 0x0f) | 0x40;
  bytes[8] = (bytes[8] & 0x3f) | 0x80;
  const hex = Array.from(bytes, (byte) => byte.toString(16).padStart(2, "0"));
  return `${kind}-${hex.slice(0, 4).join("")}-${hex.slice(4, 6).join("")}-${hex.slice(6, 8).join("")}-${hex.slice(8, 10).join("")}-${hex.slice(10).join("")}`;
}

export async function createStableKeyRowId(
  tableName: string,
  columnName: string,
  value: CellContent
): Promise<string> {
  const subtle = globalThis.crypto?.subtle;
  if (!subtle) {
    throw new Error("SHA-256 row ID generation is unavailable");
  }
  const identity = [normalizeName(tableName), normalizeName(columnName), encodeTypedValue(value)].join("\u0000");
  const digest = new Uint8Array(await subtle.digest("SHA-256", new TextEncoder().encode(identity)));
  const key = Array.from(digest.slice(0, 16), (byte) => byte.toString(16).padStart(2, "0")).join("");
  return `table-row-key-${key}`;
}

function normalizeName(value: string): string {
  return value.normalize("NFKC").trim().toLowerCase();
}

function encodeTypedValue(value: CellContent): string {
  if (value === null) return "null:";
  if (typeof value === "string") return `string:${value}`;
  if (typeof value === "boolean") return `boolean:${value ? "true" : "false"}`;
  if (Number.isNaN(value)) return "number:NaN";
  if (Object.is(value, -0)) return "number:-0";
  return `number:${String(value)}`;
}
