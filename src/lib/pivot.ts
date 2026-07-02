import type { PivotDrilldownEntry, PivotMaterializedRowKind, PivotSheetMetadata } from "../types";

export type PivotAggregator = "SUM" | "COUNT" | "AVERAGE" | "MIN" | "MAX";

export type PivotConfig = {
  rowFields: string[];
  columnField?: string;
  valueField: string;
  aggregator: PivotAggregator;
};

type AggregateState = {
  nonEmptyCount: number;
  numericCount: number;
  sum: number;
  min: number | null;
  max: number | null;
};

type PivotGroup = {
  key: string[];
  total: AggregateState;
  columns: Map<string, AggregateState>;
  records: number[];
  columnRecords: Map<string, number[]>;
};

type PivotBuildOptions = {
  includeDrilldowns: boolean;
};

export type PivotDrilldownCell = {
  entry: PivotDrilldownEntry;
  expanded: boolean;
  sourceRowCount: number;
};

export type PivotDrilldownIndex = {
  entriesByBaseRow: Map<number, PivotDrilldownEntry[]>;
  entriesByMaterializedRow: Map<number, Map<number, PivotDrilldownEntry>>;
  rowKindsByMaterializedRow: Map<number, PivotMaterializedRowKind>;
  baseRowToMaterializedRow: Map<number, number>;
};

const PIVOT_LABEL_COLLATOR = new Intl.Collator("en", { numeric: true, sensitivity: "base" });

export function createPivotTable(rows: string[][], config: PivotConfig): string[][] {
  return buildPivotTable(rows, config, { includeDrilldowns: false }).rows;
}

export function createPivotTableWithDrilldowns(rows: string[][], config: PivotConfig): {
  rows: string[][];
  metadata: PivotSheetMetadata;
} {
  const result = buildPivotTable(rows, config, { includeDrilldowns: true });
  return { rows: result.rows, metadata: result.metadata! };
}

function buildPivotTable(
  rows: string[][],
  config: PivotConfig,
  options: PivotBuildOptions
): {
  rows: string[][];
  metadata?: PivotSheetMetadata;
} {
  if (rows.length < 2) {
    throw new Error("Pivot tables need a header row and at least one data row.");
  }

  const headers = normalizeHeaders(rows[0]);
  validateConfig(headers, config);

  const fieldIndexes = Object.fromEntries(headers.map((header, index) => [header, index]));
  const valueIndex = fieldIndexes[config.valueField];
  const columnIndex = config.columnField ? fieldIndexes[config.columnField] : undefined;
  const rowFieldIndexes = config.rowFields.map((field) => fieldIndexes[field]);
  const groups = new Map<string, PivotGroup>();
  const columnTotals = new Map<string, AggregateState>();
  const columnTotalRecords = new Map<string, number[]>();
  const columnValues: string[] = [];
  const grandTotal = createAggregateState();
  const sourceRows: string[][] = [];
  const allSourceRowIndexes: number[] = [];

  for (const row of rows.slice(1)) {
    const sourceRow = headers.map((_, index) => String(row[index] ?? ""));
    const sourceRowIndex = options.includeDrilldowns ? sourceRows.push(sourceRow) - 1 : -1;
    const rowKey = rowFieldIndexes.map((index) => normalizeDimensionValue(row[index]));
    const rowKeyId = serializeKey(rowKey);
    const group = getOrCreateGroup(groups, rowKeyId, rowKey);
    const rawValue = sourceRow[valueIndex] ?? "";

    if (options.includeDrilldowns) {
      group.records.push(sourceRowIndex);
      allSourceRowIndexes.push(sourceRowIndex);
    }
    addAggregateValue(group.total, rawValue);
    addAggregateValue(grandTotal, rawValue);

    if (columnIndex !== undefined && config.columnField) {
      const columnValue = normalizeDimensionValue(row[columnIndex]);
      let columnState = group.columns.get(columnValue);
      if (!columnState) {
        columnState = createAggregateState();
        group.columns.set(columnValue, columnState);
      }
      addAggregateValue(columnState, rawValue);
      if (options.includeDrilldowns) {
        pushRecord(group.columnRecords, columnValue, sourceRowIndex);
      }

      let columnTotal = columnTotals.get(columnValue);
      if (!columnTotal) {
        columnTotal = createAggregateState();
        columnTotals.set(columnValue, columnTotal);
        columnValues.push(columnValue);
      }
      addAggregateValue(columnTotal, rawValue);
      if (options.includeDrilldowns) {
        pushRecord(columnTotalRecords, columnValue, sourceRowIndex);
      }
    }
  }

  const sortedGroups = Array.from(groups.values()).sort(comparePivotGroups);
  const sortedColumnValues = [...columnValues].sort(comparePivotLabels);
  const drilldowns: Record<string, PivotDrilldownEntry> = {};

  if (!config.columnField) {
    const header = [...config.rowFields, `${config.aggregator} of ${config.valueField}`];
    const body = sortedGroups.map((group) => [...group.key, formatAggregate(group.total, config.aggregator)]);
    const pivotRows = [
      header,
      ...body,
      [...grandTotalCells(config.rowFields.length), formatAggregate(grandTotal, config.aggregator)]
    ];

    if (!options.includeDrilldowns) {
      return { rows: pivotRows };
    }

    sortedGroups.forEach((group, groupIndex) => {
      addDrilldownEntry(drilldowns, groupIndex + 1, config.rowFields.length, rowFilters(config.rowFields, group.key), group.records);
    });
    addDrilldownEntry(drilldowns, pivotRows.length - 1, config.rowFields.length, {}, allSourceRowIndexes);

    return createPivotResult(pivotRows, rows[0], sourceRows, config, drilldowns);
  }

  const header = [...config.rowFields, ...sortedColumnValues, "Grand Total"];
  const body = sortedGroups.map((group) => [
    ...group.key,
    ...sortedColumnValues.map((columnValue) => formatAggregate(group.columns.get(columnValue), config.aggregator)),
    formatAggregate(group.total, config.aggregator)
  ]);

  const grandTotalRow = [
    ...grandTotalCells(config.rowFields.length),
    ...sortedColumnValues.map((columnValue) => formatAggregate(columnTotals.get(columnValue), config.aggregator)),
    formatAggregate(grandTotal, config.aggregator)
  ];

  const pivotRows = [header, ...body, grandTotalRow];
  if (!options.includeDrilldowns) {
    return { rows: pivotRows };
  }

  sortedGroups.forEach((group, groupIndex) => {
    const baseRow = groupIndex + 1;
    sortedColumnValues.forEach((columnValue, columnIndexOffset) => {
      const filters = { ...rowFilters(config.rowFields, group.key), [config.columnField!]: columnValue };
      addDrilldownEntry(
        drilldowns,
        baseRow,
        config.rowFields.length + columnIndexOffset,
        filters,
        group.columnRecords.get(columnValue) ?? []
      );
    });
    addDrilldownEntry(
      drilldowns,
      baseRow,
      config.rowFields.length + sortedColumnValues.length,
      rowFilters(config.rowFields, group.key),
      group.records
    );
  });

  sortedColumnValues.forEach((columnValue, columnIndexOffset) => {
    addDrilldownEntry(
      drilldowns,
      pivotRows.length - 1,
      config.rowFields.length + columnIndexOffset,
      { [config.columnField!]: columnValue },
      columnTotalRecords.get(columnValue) ?? []
    );
  });
  addDrilldownEntry(drilldowns, pivotRows.length - 1, header.length - 1, {}, allSourceRowIndexes);

  return createPivotResult(pivotRows, rows[0], sourceRows, config, drilldowns);
}

export function materializePivotRows(
  metadata: PivotSheetMetadata,
  index = createPivotDrilldownIndex(metadata)
): string[][] {
  const width = Math.max(matrixWidth(metadata.baseRows), metadata.sourceHeaders.length);
  const rows: string[][] = [];

  metadata.baseRows.forEach((baseRow, baseRowIndex) => {
    rows.push(padRow(baseRow, width));

    for (const entry of expandedEntriesForBaseRow(metadata, baseRowIndex, index)) {
      rows.push(padRow(metadata.sourceHeaders, width));
      for (const sourceRow of sourceRowsForEntry(metadata, entry)) {
        rows.push(padRow(sourceRow, width));
      }
    }
  });

  return rows;
}

export function getPivotDrilldownCell(
  metadata: PivotSheetMetadata | undefined,
  row: number,
  column: number,
  index = metadata ? createPivotDrilldownIndex(metadata) : null
): PivotDrilldownCell | null {
  if (!metadata) {
    return null;
  }

  const entry = index?.entriesByMaterializedRow.get(row)?.get(column);
  if (!entry) {
    return null;
  }

  return { entry, expanded: Boolean(metadata.expanded[entry.id]), sourceRowCount: sourceRowCount(metadata, entry) };
}

export function getPivotMaterializedRowKind(
  metadata: PivotSheetMetadata | undefined,
  row: number,
  index = metadata ? createPivotDrilldownIndex(metadata) : null
): PivotMaterializedRowKind | null {
  if (!metadata) {
    return null;
  }

  return index?.rowKindsByMaterializedRow.get(row) ?? null;
}

export function getPivotMaterializedBaseRow(index: PivotDrilldownIndex, baseRow: number): number | null {
  return index.baseRowToMaterializedRow.get(baseRow) ?? null;
}

export function createPivotDrilldownIndex(metadata: PivotSheetMetadata): PivotDrilldownIndex {
  const entriesByBaseRow = new Map<number, PivotDrilldownEntry[]>();
  const entriesByMaterializedRow = new Map<number, Map<number, PivotDrilldownEntry>>();
  const rowKindsByMaterializedRow = new Map<number, PivotMaterializedRowKind>();
  const baseRowToMaterializedRow = new Map<number, number>();

  for (const entry of Object.values(metadata.drilldowns)) {
    const entries = entriesByBaseRow.get(entry.baseRow) ?? [];
    entries.push(entry);
    entriesByBaseRow.set(entry.baseRow, entries);
  }

  for (const entries of entriesByBaseRow.values()) {
    entries.sort((left, right) => left.column - right.column || left.id.localeCompare(right.id));
  }

  let materializedRow = 0;
  for (let baseRow = 0; baseRow < metadata.baseRows.length; baseRow += 1) {
    baseRowToMaterializedRow.set(baseRow, materializedRow);
    const entries = entriesByBaseRow.get(baseRow) ?? [];
    if (entries.length > 0) {
      entriesByMaterializedRow.set(
        materializedRow,
        new Map(entries.map((entry) => [entry.column, entry]))
      );
    }
    materializedRow += 1;

    for (const entry of entries) {
      if (!metadata.expanded[entry.id]) {
        continue;
      }

      rowKindsByMaterializedRow.set(materializedRow, { kind: "detail-header", entryId: entry.id });
      materializedRow += 1;

      for (let index = 0; index < sourceRowCount(metadata, entry); index += 1) {
        rowKindsByMaterializedRow.set(materializedRow, { kind: "detail-row", entryId: entry.id });
        materializedRow += 1;
      }
    }
  }

  return {
    entriesByBaseRow,
    entriesByMaterializedRow,
    rowKindsByMaterializedRow,
    baseRowToMaterializedRow
  };
}

export function togglePivotDrilldown(metadata: PivotSheetMetadata, entryId: string): PivotSheetMetadata {
  if (!metadata.drilldowns[entryId]) {
    return metadata;
  }

  const expanded = { ...metadata.expanded };
  if (expanded[entryId]) {
    delete expanded[entryId];
  } else {
    expanded[entryId] = true;
  }

  return { ...metadata, expanded };
}

function normalizeHeaders(headerRow: string[]): string[] {
  const headers = headerRow.map((header) => String(header ?? "").trim());
  const seen = new Set<string>();
  for (const header of headers) {
    if (!header) {
      throw new Error("Pivot tables need a non-empty header for every selected column.");
    }
    if (seen.has(header)) {
      throw new Error("Pivot table headers must be unique.");
    }
    seen.add(header);
  }
  return headers;
}

function grandTotalCells(rowFieldCount: number): string[] {
  return ["Grand Total", ...Array(Math.max(rowFieldCount - 1, 0)).fill("")];
}

function validateConfig(headers: string[], config: PivotConfig) {
  const requiredFields = [...config.rowFields, config.valueField, config.columnField].filter(Boolean) as string[];
  for (const field of requiredFields) {
    if (!headers.includes(field)) {
      throw new Error(`Unknown pivot field: ${field}`);
    }
  }

  if (config.rowFields.length === 0) {
    throw new Error("Choose at least one row field.");
  }
}

function getOrCreateGroup(groups: Map<string, PivotGroup>, rowKeyId: string, rowKey: string[]): PivotGroup {
  let group = groups.get(rowKeyId);
  if (!group) {
    group = {
      key: rowKey,
      total: createAggregateState(),
      columns: new Map(),
      records: [],
      columnRecords: new Map()
    };
    groups.set(rowKeyId, group);
  }
  return group;
}

function comparePivotGroups(left: PivotGroup, right: PivotGroup): number {
  const length = Math.max(left.key.length, right.key.length);
  for (let index = 0; index < length; index += 1) {
    const comparison = comparePivotLabels(left.key[index] ?? "", right.key[index] ?? "");
    if (comparison !== 0) {
      return comparison;
    }
  }
  return 0;
}

function comparePivotLabels(left: string, right: string): number {
  if (left === right) {
    return 0;
  }
  if (left === "") {
    return 1;
  }
  if (right === "") {
    return -1;
  }
  return PIVOT_LABEL_COLLATOR.compare(left, right);
}

function createAggregateState(): AggregateState {
  return {
    nonEmptyCount: 0,
    numericCount: 0,
    sum: 0,
    min: null,
    max: null
  };
}

function addAggregateValue(state: AggregateState, rawValue: string) {
  if (String(rawValue).trim() !== "") {
    state.nonEmptyCount += 1;
  }

  const numericValue = parseNumericValue(rawValue);
  if (numericValue === null) {
    return;
  }

  state.numericCount += 1;
  state.sum += numericValue;
  state.min = state.min === null ? numericValue : Math.min(state.min, numericValue);
  state.max = state.max === null ? numericValue : Math.max(state.max, numericValue);
}

function formatAggregate(state: AggregateState | undefined, aggregator: PivotAggregator): string {
  if (!state) {
    return aggregator === "COUNT" ? "0" : "";
  }

  if (aggregator === "COUNT") {
    return String(state.nonEmptyCount);
  }

  if (state.numericCount === 0) {
    return "";
  }

  if (aggregator === "SUM") {
    return formatNumber(state.sum);
  }

  if (aggregator === "AVERAGE") {
    return formatNumber(state.sum / state.numericCount);
  }

  if (aggregator === "MIN") {
    return formatNumber(state.min ?? 0);
  }

  return formatNumber(state.max ?? 0);
}

function parseNumericValue(rawValue: string): number | null {
  let value = String(rawValue).trim();
  if (!value) {
    return null;
  }

  let isNegative = false;
  if (/^\(.*\)$/.test(value)) {
    isNegative = true;
    value = value.slice(1, -1).trim();
  }

  const isPercent = value.endsWith("%");
  if (isPercent) {
    value = value.slice(0, -1).trim();
  }

  value = value.replace(/[,$£€¥₹\s]/g, "");
  if (!value || /[A-Za-z]/.test(value)) {
    return null;
  }

  const parsed = Number(value);
  if (!Number.isFinite(parsed)) {
    return null;
  }

  const signed = isNegative ? parsed * -1 : parsed;
  return isPercent ? signed / 100 : signed;
}

function normalizeDimensionValue(value: string | undefined): string {
  return String(value ?? "").trim();
}

function serializeKey(key: string[]): string {
  return JSON.stringify(key);
}

function createPivotResult(
  rows: string[][],
  sourceHeaderRow: string[],
  sourceRows: string[][],
  config: PivotConfig,
  drilldowns: Record<string, PivotDrilldownEntry>
) {
  const metadata: PivotSheetMetadata = {
    sourceHeaders: normalizeHeaders(sourceHeaderRow),
    sourceRows: sourceRows.map((row) => [...row]),
    baseRows: rows.map((row) => [...row]),
    config: {
      rowFields: [...config.rowFields],
      columnField: config.columnField,
      valueField: config.valueField,
      aggregator: config.aggregator
    },
    drilldowns,
    expanded: {}
  };

  return { rows, metadata };
}

function addDrilldownEntry(
  drilldowns: Record<string, PivotDrilldownEntry>,
  baseRow: number,
  column: number,
  filters: Record<string, string>,
  sourceRowIndexes: number[]
) {
  if (sourceRowIndexes.length === 0) {
    return;
  }

  const id = `drilldown-r${baseRow}-c${column}`;
  drilldowns[id] = {
    id,
    baseRow,
    column,
    filters,
    sourceRowIndexes: [...sourceRowIndexes]
  };
}

function rowFilters(rowFields: string[], rowKey: string[]): Record<string, string> {
  return Object.fromEntries(rowFields.map((field, index) => [field, rowKey[index] ?? ""]));
}

function pushRecord(records: Map<string, number[]>, key: string, rowIndex: number) {
  const rows = records.get(key);
  if (rows) {
    rows.push(rowIndex);
    return;
  }
  records.set(key, [rowIndex]);
}

function expandedEntriesForBaseRow(
  metadata: PivotSheetMetadata,
  baseRow: number,
  index: PivotDrilldownIndex
): PivotDrilldownEntry[] {
  return (index.entriesByBaseRow.get(baseRow) ?? []).filter((entry) => metadata.expanded[entry.id]);
}

function sourceRowsForEntry(metadata: PivotSheetMetadata, entry: PivotDrilldownEntry): string[][] {
  const legacyRows = (entry as PivotDrilldownEntry & { sourceRows?: string[][] }).sourceRows;
  if (Array.isArray(legacyRows)) {
    return legacyRows.map((row) => [...row]);
  }

  return (entry.sourceRowIndexes ?? [])
    .map((sourceRowIndex) => metadata.sourceRows?.[sourceRowIndex])
    .filter((row): row is string[] => Array.isArray(row))
    .map((row) => [...row]);
}

function sourceRowCount(metadata: PivotSheetMetadata, entry: PivotDrilldownEntry): number {
  const legacyRows = (entry as PivotDrilldownEntry & { sourceRows?: string[][] }).sourceRows;
  return Array.isArray(legacyRows) ? legacyRows.length : (entry.sourceRowIndexes ?? []).length;
}

function matrixWidth(rows: string[][]): number {
  return rows.reduce((width, row) => Math.max(width, row.length), 0);
}

function padRow(row: string[], width: number): string[] {
  return [...row, ...Array(Math.max(width - row.length, 0)).fill("")];
}

function formatNumber(value: number): string {
  return Number.isInteger(value) ? String(value) : String(Number(value.toFixed(8)));
}
