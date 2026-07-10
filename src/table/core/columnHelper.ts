import {
  DEFAULT_COLUMN_WIDTH,
  MAX_COLUMN_WIDTH,
  MIN_COLUMN_WIDTH,
  clampColumnWidth
} from "../../lib/sheetDimensions";
import type { ColumnBaseDef, ColumnDef, ColumnValueContext } from "./types";

type AccessorColumnDef<TRow, TValue> = ColumnBaseDef<TRow, TValue> & {
  kind: "accessor";
  accessor(row: TRow): TValue;
  update(row: TRow, value: TValue): TRow;
};
type ComputedColumnDef<TRow, TValue> = ColumnBaseDef<TRow, TValue> & {
  kind: "computed";
  calculate(context: ColumnValueContext<TRow>): TValue;
};
type DisplayColumnDef<TRow> = ColumnBaseDef<TRow> & { kind: "display" };

export type ColumnHelper<TRow> = {
  accessor<TKey extends keyof TRow>(
    key: TKey,
    options: ColumnBaseDef<TRow, TRow[TKey]>
  ): AccessorColumnDef<TRow, TRow[TKey]>;
  computed<TValue>(
    options: ColumnBaseDef<TRow, TValue> & { calculate(context: ColumnValueContext<TRow>): TValue }
  ): ComputedColumnDef<TRow, TValue>;
  display(options: ColumnBaseDef<TRow>): DisplayColumnDef<TRow>;
};

export function createColumnHelper<TRow>(): ColumnHelper<TRow> {
  return {
    accessor(key, options) {
      return {
        ...options,
        kind: "accessor",
        accessor: (row) => row[key],
        update: (row, value) => ({ ...row, [key]: value })
      };
    },
    computed(options) {
      return { ...options, kind: "computed" };
    },
    display(options) {
      return { ...options, kind: "display" };
    }
  };
}

export function normalizeColumns<TColumn extends ColumnDef<any, any>>(
  columns: readonly TColumn[]
): readonly TColumn[] {
  const ids = new Set<string>();
  const normalized = columns.map((column) => {
    const id = column.id.trim();
    if (!id) {
      throw new Error("Column id must not be blank");
    }
    if (ids.has(id)) {
      throw new Error(`Duplicate column id: ${id}`);
    }
    ids.add(id);

    const minWidth = clampColumnWidth(column.minWidth ?? MIN_COLUMN_WIDTH);
    const maxWidth = Math.max(minWidth, clampColumnWidth(column.maxWidth ?? MAX_COLUMN_WIDTH));
    const width = Math.min(maxWidth, Math.max(minWidth, clampColumnWidth(column.width ?? DEFAULT_COLUMN_WIDTH)));
    return Object.freeze({ ...column, id, width, minWidth, maxWidth }) as TColumn;
  });
  return Object.freeze(normalized);
}
