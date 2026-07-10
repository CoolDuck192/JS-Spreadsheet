export type ExcelTableNameValidation =
  | { valid: true; normalizedKey: string }
  | {
      valid: false;
      reason: "empty" | "tooLong" | "invalidCharacter" | "reserved" | "cellReference";
    };

const EXCEL_MAX_ROW = 1_048_576;
const EXCEL_MAX_COLUMN = 16_384;

export function validateExcelTableName(name: string): ExcelTableNameValidation {
  const normalized = name.normalize("NFKC");
  const characters = [...normalized];
  if (normalized.trim().length === 0) {
    return { valid: false, reason: "empty" };
  }
  if (characters.length > 255) {
    return { valid: false, reason: "tooLong" };
  }
  if (normalized !== normalized.trim() || !/^(?:\p{L}|_|\\)$/u.test(characters[0] ?? "")) {
    return { valid: false, reason: "invalidCharacter" };
  }
  if (!characters.slice(1).every((character) => /^(?:\p{L}|\p{N}|_|\.)$/u.test(character))) {
    return { valid: false, reason: "invalidCharacter" };
  }

  const normalizedKey = normalizeExcelTableNameKey(normalized);
  if (normalizedKey === "r" || normalizedKey === "c") {
    return { valid: false, reason: "reserved" };
  }
  if (isR1C1Reference(normalizedKey) || isInGridA1Reference(normalizedKey)) {
    return { valid: false, reason: "cellReference" };
  }
  return { valid: true, normalizedKey };
}

export function normalizeExcelTableNameKey(name: string): string {
  return name.normalize("NFKC").toLowerCase();
}

function isR1C1Reference(name: string): boolean {
  return /^r[1-9]\d*c[1-9]\d*$/.test(name);
}

function isInGridA1Reference(name: string): boolean {
  const match = /^([a-z]+)([1-9]\d*)$/.exec(name);
  if (!match) return false;
  const row = Number(match[2]);
  if (!Number.isSafeInteger(row) || row > EXCEL_MAX_ROW) return false;
  let column = 0;
  for (const character of match[1]) {
    column = column * 26 + character.charCodeAt(0) - 96;
    if (column > EXCEL_MAX_COLUMN) return false;
  }
  return column >= 1 && column <= EXCEL_MAX_COLUMN;
}
