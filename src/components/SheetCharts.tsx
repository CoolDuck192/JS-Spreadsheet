import { Trash2 } from "lucide-react";
import type { SheetChart, SheetModel } from "../types";
import { formatCellAddress, normalizeRange } from "../lib/addressing";
import { createChartData, type ChartDatum } from "../lib/charts";
import type { FormulaEngine } from "../lib/formulaEngine";

type SheetChartsProps = {
  charts: SheetChart[];
  sheet: SheetModel;
  formulaEngine: FormulaEngine;
  onDelete: (chartId: string) => void;
};

const SVG_WIDTH = 280;
const SVG_HEIGHT = 138;
const PLOT_TOP = 14;
const PLOT_RIGHT = 12;
const PLOT_BOTTOM = 28;
const PLOT_LEFT = 32;
const PLOT_WIDTH = SVG_WIDTH - PLOT_LEFT - PLOT_RIGHT;
const PLOT_HEIGHT = SVG_HEIGHT - PLOT_TOP - PLOT_BOTTOM;
const PIE_COLORS = ["#2f7d9f", "#398d71", "#b26b3f", "#7a6fb2", "#b04f6f", "#5f7f38"];

export function SheetCharts({ charts, sheet, formulaEngine, onDelete }: SheetChartsProps) {
  if (charts.length === 0) {
    return null;
  }

  return (
    <div className="sheet-chart-layer" aria-label="Embedded charts">
      {charts.map((chart) => {
        const chartData = createChartData(chartRows(sheet, chart, formulaEngine), chart.title || "Chart");
        const title = chart.title || chartData.title;
        const pointLabel = `${chartData.data.length} ${chartData.data.length === 1 ? "point" : "points"}`;

        return (
          <figure key={chart.id} role="figure" aria-label={`Chart ${title}`} className="sheet-chart" data-testid="sheet-chart">
            <figcaption className="sheet-chart-header">
              <span>
                <strong>{title}</strong>
                <small>{pointLabel}</small>
              </span>
              <button type="button" aria-label={`Delete chart ${title}`} title={`Delete chart ${title}`} onClick={() => onDelete(chart.id)}>
                <Trash2 />
              </button>
            </figcaption>
            {chartData.data.length > 0 ? (
              renderChartSvg(chart.type, chartData.data)
            ) : (
              <div className="sheet-chart-empty">No numeric data</div>
            )}
          </figure>
        );
      })}
    </div>
  );
}

function renderChartSvg(type: SheetChart["type"], data: ChartDatum[]) {
  if (type === "line") {
    return <LineChartSvg data={data} />;
  }

  if (type === "pie") {
    return <PieChartSvg data={data} />;
  }

  return <BarChartSvg data={data} />;
}

function chartRows(sheet: SheetModel, chart: SheetChart, formulaEngine: FormulaEngine): string[][] {
  const range = normalizeRange(chart.range);
  const rows: string[][] = [];

  for (let row = range.start.row; row <= range.end.row; row += 1) {
    const values: string[] = [];
    for (let column = range.start.column; column <= range.end.column; column += 1) {
      values.push(formulaEngine.getDisplayValue(sheet.id, formatCellAddress({ row, column })));
    }
    rows.push(values);
  }

  return rows;
}

function BarChartSvg({ data }: { data: ChartDatum[] }) {
  const maxValue = maxMagnitude(data);
  const barGap = 8;
  const barWidth = Math.max(12, (PLOT_WIDTH - barGap * Math.max(data.length - 1, 0)) / data.length);

  return (
    <svg className="sheet-chart-svg" viewBox={`0 0 ${SVG_WIDTH} ${SVG_HEIGHT}`} role="img" aria-label="Bar chart">
      <ChartAxis />
      {data.map((datum, index) => {
        const height = Math.max(2, (Math.abs(datum.value) / maxValue) * PLOT_HEIGHT);
        const x = PLOT_LEFT + index * (barWidth + barGap);
        const y = PLOT_TOP + PLOT_HEIGHT - height;
        return (
          <g key={`${datum.label}-${index}`}>
            <rect className="sheet-chart-bar" x={x} y={y} width={barWidth} height={height} rx="3" />
            <title>{`${datum.label}: ${formatChartValue(datum.value)}`}</title>
            {index < 5 ? (
              <text className="sheet-chart-label" x={x + barWidth / 2} y={SVG_HEIGHT - 8} textAnchor="middle">
                {shortLabel(datum.label)}
              </text>
            ) : null}
          </g>
        );
      })}
    </svg>
  );
}

function LineChartSvg({ data }: { data: ChartDatum[] }) {
  const maxValue = maxMagnitude(data);
  const xStep = data.length > 1 ? PLOT_WIDTH / (data.length - 1) : 0;
  const points = data.map((datum, index) => ({
    x: PLOT_LEFT + (data.length > 1 ? index * xStep : PLOT_WIDTH / 2),
    y: PLOT_TOP + PLOT_HEIGHT - (Math.abs(datum.value) / maxValue) * PLOT_HEIGHT,
    datum
  }));

  return (
    <svg className="sheet-chart-svg" viewBox={`0 0 ${SVG_WIDTH} ${SVG_HEIGHT}`} role="img" aria-label="Line chart">
      <ChartAxis />
      <polyline className="sheet-chart-line" points={points.map((point) => `${point.x},${point.y}`).join(" ")} />
      {points.map((point, index) => (
        <g key={`${point.datum.label}-${index}`}>
          <circle className="sheet-chart-dot" cx={point.x} cy={point.y} r="4" />
          <title>{`${point.datum.label}: ${formatChartValue(point.datum.value)}`}</title>
          {index < 5 ? (
            <text className="sheet-chart-label" x={point.x} y={SVG_HEIGHT - 8} textAnchor="middle">
              {shortLabel(point.datum.label)}
            </text>
          ) : null}
        </g>
      ))}
    </svg>
  );
}

function PieChartSvg({ data }: { data: ChartDatum[] }) {
  const slices = createPieSlices(data);
  if (slices.length === 0) {
    return <div className="sheet-chart-empty">No positive values</div>;
  }

  const centerX = 78;
  const centerY = 66;
  const radius = 43;

  return (
    <svg className="sheet-chart-svg" viewBox={`0 0 ${SVG_WIDTH} ${SVG_HEIGHT}`} role="img" aria-label="Pie chart">
      {slices.map((slice, index) =>
        slice.endAngle - slice.startAngle >= Math.PI * 2 - 0.0001 ? (
          <circle
            key={`${slice.datum.label}-${index}`}
            className="sheet-chart-pie-slice"
            cx={centerX}
            cy={centerY}
            r={radius}
            fill={PIE_COLORS[index % PIE_COLORS.length]}
          >
            <title>{`${slice.datum.label}: ${formatChartValue(slice.datum.value)}`}</title>
          </circle>
        ) : (
          <path
            key={`${slice.datum.label}-${index}`}
            className="sheet-chart-pie-slice"
            d={describePieSlice(centerX, centerY, radius, slice.startAngle, slice.endAngle)}
            fill={PIE_COLORS[index % PIE_COLORS.length]}
          >
            <title>{`${slice.datum.label}: ${formatChartValue(slice.datum.value)}`}</title>
          </path>
        )
      )}
      {slices.slice(0, 5).map((slice, index) => {
        const y = 28 + index * 18;
        return (
          <g key={`${slice.datum.label}-${index}-legend`}>
            <rect className="sheet-chart-pie-key" x={146} y={y - 9} width={9} height={9} fill={PIE_COLORS[index % PIE_COLORS.length]} />
            <text className="sheet-chart-label" x={162} y={y} textAnchor="start">
              {`${shortLabel(slice.datum.label)} ${Math.round(slice.share * 100)}%`}
            </text>
          </g>
        );
      })}
    </svg>
  );
}

function ChartAxis() {
  return (
    <>
      <line className="sheet-chart-axis" x1={PLOT_LEFT} y1={PLOT_TOP} x2={PLOT_LEFT} y2={PLOT_TOP + PLOT_HEIGHT} />
      <line
        className="sheet-chart-axis"
        x1={PLOT_LEFT}
        y1={PLOT_TOP + PLOT_HEIGHT}
        x2={SVG_WIDTH - PLOT_RIGHT}
        y2={PLOT_TOP + PLOT_HEIGHT}
      />
      <line
        className="sheet-chart-gridline"
        x1={PLOT_LEFT}
        y1={PLOT_TOP + PLOT_HEIGHT / 2}
        x2={SVG_WIDTH - PLOT_RIGHT}
        y2={PLOT_TOP + PLOT_HEIGHT / 2}
      />
    </>
  );
}

function createPieSlices(data: ChartDatum[]) {
  const positiveData = data.filter((datum) => datum.value > 0);
  const total = positiveData.reduce((sum, datum) => sum + datum.value, 0);
  if (total <= 0) {
    return [];
  }

  let startAngle = -Math.PI / 2;
  return positiveData.map((datum) => {
    const share = datum.value / total;
    const endAngle = startAngle + share * Math.PI * 2;
    const slice = { datum, startAngle, endAngle, share };
    startAngle = endAngle;
    return slice;
  });
}

function describePieSlice(centerX: number, centerY: number, radius: number, startAngle: number, endAngle: number): string {
  const start = polarPoint(centerX, centerY, radius, startAngle);
  const end = polarPoint(centerX, centerY, radius, endAngle);
  const largeArcFlag = endAngle - startAngle > Math.PI ? 1 : 0;
  return [`M ${centerX} ${centerY}`, `L ${start.x} ${start.y}`, `A ${radius} ${radius} 0 ${largeArcFlag} 1 ${end.x} ${end.y}`, "Z"].join(" ");
}

function polarPoint(centerX: number, centerY: number, radius: number, angle: number) {
  return {
    x: centerX + radius * Math.cos(angle),
    y: centerY + radius * Math.sin(angle)
  };
}

function maxMagnitude(data: ChartDatum[]): number {
  return Math.max(1, ...data.map((datum) => Math.abs(datum.value)));
}

function formatChartValue(value: number): string {
  return Number.isInteger(value) ? value.toLocaleString() : value.toLocaleString(undefined, { maximumFractionDigits: 2 });
}

function shortLabel(value: string): string {
  return value.length > 9 ? `${value.slice(0, 8)}...` : value;
}
