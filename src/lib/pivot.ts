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
  sourceRows: number[];
};

/**
 * Per-cell drill-down map parallel to the pivot table: each entry lists the
 * indexes (into the input `rows` array) of the source data rows that were
 * aggregated into that cell, or null for cells with nothing to drill into
 * (header row, empty intersections).
 */
export type PivotDrillDownGrid = Array<Array<number[] | null>>;

export type PivotTableWithDetails = {
  table: string[][];
  drillDown: PivotDrillDownGrid;
};

type PivotGroup = {
  key: string[];
  total: AggregateState;
  columns: Map<string, AggregateState>;
};

const PIVOT_LABEL_COLLATOR = new Intl.Collator("en", { numeric: true, sensitivity: "base" });

export function createPivotTable(rows: string[][], config: PivotConfig): string[][] {
  return createPivotTableWithDetails(rows, config).table;
}

export function createPivotTableWithDetails(rows: string[][], config: PivotConfig): PivotTableWithDetails {
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
  const columnValues: string[] = [];
  const grandTotal = createAggregateState();

  for (let sourceIndex = 1; sourceIndex < rows.length; sourceIndex += 1) {
    const row = rows[sourceIndex];
    const rowKey = rowFieldIndexes.map((index) => normalizeDimensionValue(row[index]));
    const rowKeyId = serializeKey(rowKey);
    const group = getOrCreateGroup(groups, rowKeyId, rowKey);
    const rawValue = row[valueIndex] ?? "";

    addAggregateValue(group.total, rawValue, sourceIndex);
    addAggregateValue(grandTotal, rawValue, sourceIndex);

    if (columnIndex !== undefined && config.columnField) {
      const columnValue = normalizeDimensionValue(row[columnIndex]);
      let columnState = group.columns.get(columnValue);
      if (!columnState) {
        columnState = createAggregateState();
        group.columns.set(columnValue, columnState);
      }
      addAggregateValue(columnState, rawValue, sourceIndex);

      let columnTotal = columnTotals.get(columnValue);
      if (!columnTotal) {
        columnTotal = createAggregateState();
        columnTotals.set(columnValue, columnTotal);
        columnValues.push(columnValue);
      }
      addAggregateValue(columnTotal, rawValue, sourceIndex);
    }
  }

  const sortedGroups = Array.from(groups.values()).sort(comparePivotGroups);
  const sortedColumnValues = [...columnValues].sort(comparePivotLabels);

  if (!config.columnField) {
    const header = [...config.rowFields, `${config.aggregator} of ${config.valueField}`];
    const body = sortedGroups.map((group) => [...group.key, formatAggregate(group.total, config.aggregator)]);
    const table = [
      header,
      ...body,
      [...grandTotalCells(config.rowFields.length), formatAggregate(grandTotal, config.aggregator)]
    ];
    const drillDown: PivotDrillDownGrid = [
      header.map(() => null),
      ...sortedGroups.map((group) => header.map(() => stateDrillDown(group.total))),
      header.map(() => stateDrillDown(grandTotal))
    ];
    return { table, drillDown };
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

  const drillDown: PivotDrillDownGrid = [
    header.map(() => null),
    ...sortedGroups.map((group) => [
      ...group.key.map(() => stateDrillDown(group.total)),
      ...sortedColumnValues.map((columnValue) => stateDrillDown(group.columns.get(columnValue))),
      stateDrillDown(group.total)
    ]),
    [
      ...grandTotalCells(config.rowFields.length).map(() => stateDrillDown(grandTotal)),
      ...sortedColumnValues.map((columnValue) => stateDrillDown(columnTotals.get(columnValue))),
      stateDrillDown(grandTotal)
    ]
  ];

  return { table: [header, ...body, grandTotalRow], drillDown };
}

function stateDrillDown(state: AggregateState | undefined): number[] | null {
  return state && state.sourceRows.length > 0 ? [...state.sourceRows] : null;
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
      columns: new Map()
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
    max: null,
    sourceRows: []
  };
}

function addAggregateValue(state: AggregateState, rawValue: string, sourceIndex: number) {
  state.sourceRows.push(sourceIndex);
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

function formatNumber(value: number): string {
  return Number.isInteger(value) ? String(value) : String(Number(value.toFixed(8)));
}
