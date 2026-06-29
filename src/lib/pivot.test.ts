import { describe, expect, it } from "vitest";
import {
  createPivotDrilldownIndex,
  createPivotTable,
  createPivotTableWithDrilldowns,
  getPivotDrilldownCell,
  getPivotMaterializedBaseRow,
  getPivotMaterializedRowKind,
  materializePivotRows
} from "./pivot";

describe("pivot", () => {
  const rows = [
    ["Region", "Product", "Sales"],
    ["West", "Hardware", "10"],
    ["West", "Software", "20"],
    ["East", "Hardware", "8"],
    ["East", "Software", "7"]
  ];

  it("creates a row and column pivot with grand totals", () => {
    expect(
      createPivotTable(rows, {
        rowFields: ["Region"],
        columnField: "Product",
        valueField: "Sales",
        aggregator: "SUM"
      })
    ).toEqual([
      ["Region", "Hardware", "Software", "Grand Total"],
      ["East", "8", "7", "15"],
      ["West", "10", "20", "30"],
      ["Grand Total", "18", "27", "45"]
    ]);
  });

  it("creates a row-only pivot with averages", () => {
    expect(
      createPivotTable(rows, {
        rowFields: ["Region"],
        valueField: "Sales",
        aggregator: "AVERAGE"
      })
    ).toEqual([
      ["Region", "AVERAGE of Sales"],
      ["East", "7.5"],
      ["West", "15"],
      ["Grand Total", "11.25"]
    ]);
  });

  it("sorts pivot row groups and column labels naturally", () => {
    expect(
      createPivotTable(
        [
          ["Account", "Quarter", "Sales"],
          ["Team 10", "Q2", "20"],
          ["Team 2", "Q1", "5"],
          ["Team 1", "Q2", "7"],
          ["Team 2", "Q2", "8"],
          ["Team 1", "Q1", "3"]
        ],
        {
          rowFields: ["Account"],
          columnField: "Quarter",
          valueField: "Sales",
          aggregator: "SUM"
        }
      )
    ).toEqual([
      ["Account", "Q1", "Q2", "Grand Total"],
      ["Team 1", "3", "7", "10"],
      ["Team 2", "5", "8", "13"],
      ["Team 10", "", "20", "20"],
      ["Grand Total", "8", "35", "43"]
    ]);
  });

  it("counts non-empty values", () => {
    expect(
      createPivotTable(
        [
          ["Team", "Assignee"],
          ["Ops", "Asha"],
          ["Ops", ""],
          ["Design", "Mina"]
        ],
        {
          rowFields: ["Team"],
          valueField: "Assignee",
          aggregator: "COUNT"
        }
      )
    ).toEqual([
      ["Team", "COUNT of Assignee"],
      ["Design", "1"],
      ["Ops", "1"],
      ["Grand Total", "2"]
    ]);
  });

  it("aggregates formatted numeric values in real-world data", () => {
    expect(
      createPivotTable(
        [
          ["Region", "Product", "Revenue"],
          ["West", "Hardware", "$1,200.50"],
          ["West", "Hardware", "(200.25)"],
          ["East", "Hardware", "3,000"],
          ["East", "Hardware", ""]
        ],
        {
          rowFields: ["Region"],
          columnField: "Product",
          valueField: "Revenue",
          aggregator: "SUM"
        }
      )
    ).toEqual([
      ["Region", "Hardware", "Grand Total"],
      ["East", "3000", "3000"],
      ["West", "1000.25", "1000.25"],
      ["Grand Total", "4000.25", "4000.25"]
    ]);
  });

  it("rejects blank and duplicate headers before building a pivot", () => {
    expect(() =>
      createPivotTable(
        [
          ["Region", "Region", "Sales"],
          ["West", "Hardware", "10"]
        ],
        {
          rowFields: ["Region"],
          valueField: "Sales",
          aggregator: "SUM"
        }
      )
    ).toThrow("unique");

    expect(() =>
      createPivotTable(
        [
          ["Region", "", "Sales"],
          ["West", "Hardware", "10"]
        ],
        {
          rowFields: ["Region"],
          valueField: "Sales",
          aggregator: "SUM"
        }
      )
    ).toThrow("header");
  });

  it("aggregates high-cardinality datasets quickly", () => {
    const largeRows = [["Account", "Month", "Amount"]];
    for (let index = 0; index < 3000; index += 1) {
      largeRows.push([`Account ${index}`, `Month ${index % 6}`, "1"]);
    }

    const startedAt = performance.now();
    const result = createPivotTable(largeRows, {
      rowFields: ["Account"],
      columnField: "Month",
      valueField: "Amount",
      aggregator: "SUM"
    });
    const duration = performance.now() - startedAt;

    expect(result).toHaveLength(3002);
    expect(result[0]).toEqual(["Account", "Month 0", "Month 1", "Month 2", "Month 3", "Month 4", "Month 5", "Grand Total"]);
    expect(result.at(-1)).toEqual(["Grand Total", "500", "500", "500", "500", "500", "500", "3000"]);
    expect(duration).toBeLessThan(250);
  });

  it("tracks source rows for pivot drilldown cells", () => {
    const pivot = createPivotTableWithDrilldowns(rows, {
      rowFields: ["Region", "Product"],
      valueField: "Sales",
      aggregator: "SUM"
    });

    expect(pivot.rows).toEqual([
      ["Region", "Product", "SUM of Sales"],
      ["East", "Hardware", "8"],
      ["East", "Software", "7"],
      ["West", "Hardware", "10"],
      ["West", "Software", "20"],
      ["Grand Total", "", "45"]
    ]);

    const drilldown = getPivotDrilldownCell(pivot.metadata, 1, 2);
    expect(pivot.metadata.sourceRows).toEqual(rows.slice(1));
    expect(drilldown).toMatchObject({
      expanded: false,
      sourceRowCount: 1,
      sourceRows: [["East", "Hardware", "8"]],
      entry: {
        filters: { Region: "East", Product: "Hardware" },
        sourceRowIndexes: [2]
      }
    });
    expect("sourceRows" in drilldown!.entry).toBe(false);
  });

  it("materializes expanded pivot drilldown rows below the summary row", () => {
    const pivot = createPivotTableWithDrilldowns(rows, {
      rowFields: ["Region", "Product"],
      valueField: "Sales",
      aggregator: "SUM"
    });
    const drilldown = getPivotDrilldownCell(pivot.metadata, 1, 2);

    const expandedRows = materializePivotRows({
      ...pivot.metadata,
      expanded: { [drilldown!.entry.id]: true }
    });

    expect(expandedRows.slice(0, 5)).toEqual([
      ["Region", "Product", "SUM of Sales"],
      ["East", "Hardware", "8"],
      ["Region", "Product", "Sales"],
      ["East", "Hardware", "8"],
      ["East", "Software", "7"]
    ]);
  });

  it("indexes pivot drilldowns for repeated render lookups", () => {
    const pivot = createPivotTableWithDrilldowns(rows, {
      rowFields: ["Region", "Product"],
      valueField: "Sales",
      aggregator: "SUM"
    });
    const drilldown = getPivotDrilldownCell(pivot.metadata, 1, 2);
    const expandedMetadata = {
      ...pivot.metadata,
      expanded: { [drilldown!.entry.id]: true }
    };

    const index = createPivotDrilldownIndex(expandedMetadata);

    expect(getPivotDrilldownCell(expandedMetadata, 1, 2, index)?.entry.id).toBe(drilldown!.entry.id);
    expect(getPivotMaterializedRowKind(expandedMetadata, 2, index)).toEqual({
      kind: "detail-header",
      entryId: drilldown!.entry.id
    });
    expect(getPivotMaterializedBaseRow(index, 2)).toBe(4);
  });
});
