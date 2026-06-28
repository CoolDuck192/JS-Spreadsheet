export type ChartDatum = {
  label: string;
  value: number;
};

export type ChartData = {
  title: string;
  labelsTitle: string;
  valuesTitle: string;
  data: ChartDatum[];
};

export function createChartData(rows: string[][], fallbackTitle = "Chart"): ChartData {
  const firstRow = rows[0] ?? [];
  const hasHeader = rows.length > 1 && firstRow[0]?.trim() !== "" && firstRow[1]?.trim() !== "" && parseNumericValue(firstRow[1]) === null;
  const labelsTitle = hasHeader ? firstRow[0].trim() : "Label";
  const valuesTitle = hasHeader ? firstRow[1].trim() : "Value";
  const sourceRows = hasHeader ? rows.slice(1) : rows;
  const data = sourceRows.flatMap((row) => {
    const label = row[0]?.trim() ?? "";
    const value = parseNumericValue(row[1]);
    if (!label || value === null) {
      return [];
    }
    return [{ label, value }];
  });

  return {
    title: hasHeader ? `${valuesTitle} by ${labelsTitle}` : fallbackTitle,
    labelsTitle,
    valuesTitle,
    data
  };
}

function parseNumericValue(value: string | undefined): number | null {
  if (value === undefined) {
    return null;
  }

  const normalized = value.trim().replace(/[,$%]/g, "");
  if (!normalized) {
    return null;
  }

  const parsed = Number(normalized);
  return Number.isFinite(parsed) ? parsed : null;
}
