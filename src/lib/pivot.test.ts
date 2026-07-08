import { describe, expect, it } from "vitest";
import { createPivotTable, createPivotTableWithDetails } from "./pivot";

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

  it("maps every value cell to its contributing source rows for drill-down", () => {
    const { table, drillDown } = createPivotTableWithDetails(rows, {
      rowFields: ["Region"],
      columnField: "Product",
      valueField: "Sales",
      aggregator: "SUM"
    });

    expect(table[1]).toEqual(["East", "8", "7", "15"]);
    // Header row has nothing to drill into.
    expect(drillDown[0]).toEqual([null, null, null, null]);
    // East/Hardware = source row 3, East/Software = row 4, East total = both.
    expect(drillDown[1]).toEqual([[3, 4], [3], [4], [3, 4]]);
    // West/Hardware = row 1, West/Software = row 2.
    expect(drillDown[2]).toEqual([[1, 2], [1], [2], [1, 2]]);
    // Grand-total row drills into column totals and the full data set.
    expect(drillDown[3]).toEqual([[1, 2, 3, 4], [1, 3], [2, 4], [1, 2, 3, 4]]);
  });

  it("returns null drill-down for empty row/column intersections", () => {
    const { drillDown, table } = createPivotTableWithDetails(
      [
        ["Region", "Product", "Sales"],
        ["West", "Hardware", "10"],
        ["East", "Software", "7"]
      ],
      { rowFields: ["Region"], columnField: "Product", valueField: "Sales", aggregator: "SUM" }
    );

    // East has no Hardware entry: the empty intersection is not drillable.
    expect(table[1]).toEqual(["East", "", "7", "7"]);
    expect(drillDown[1][1]).toBeNull();
    expect(drillDown[1][2]).toEqual([2]);
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

  it("labels empty dimension values as (blank) like Excel", () => {
    const table = createPivotTable(
      [
        ["Region", "Product", "Sales"],
        ["West", "Hardware", "10"],
        ["", "Software", "5"]
      ],
      { rowFields: ["Region"], columnField: "Product", valueField: "Sales", aggregator: "SUM" }
    );

    expect(table[0]).toEqual(["Region", "Hardware", "Software", "Grand Total"]);
    // Blank labels sort last and render as "(blank)".
    expect(table[1]).toEqual(["West", "10", "", "10"]);
    expect(table[2]).toEqual(["(blank)", "", "5", "5"]);
  });

  it("supports count-numbers and product aggregators", () => {
    const mixedRows = [
      ["Region", "Sales"],
      ["West", "2"],
      ["West", "pending"],
      ["West", "3"],
      ["East", "4"]
    ];

    expect(
      createPivotTable(mixedRows, { rowFields: ["Region"], valueField: "Sales", aggregator: "COUNTNUMS" })
    ).toEqual([
      ["Region", "COUNTNUMS of Sales"],
      ["East", "1"],
      ["West", "2"],
      ["Grand Total", "3"]
    ]);

    expect(
      createPivotTable(mixedRows, { rowFields: ["Region"], valueField: "Sales", aggregator: "PRODUCT" })
    ).toEqual([
      ["Region", "PRODUCT of Sales"],
      ["East", "4"],
      ["West", "6"],
      ["Grand Total", "24"]
    ]);
  });

  it("renders empty intersections as 0 for count aggregators and blank otherwise", () => {
    const sparseRows = [
      ["Region", "Product", "Sales"],
      ["West", "Hardware", "10"],
      ["East", "Software", "7"]
    ];

    expect(
      createPivotTable(sparseRows, { rowFields: ["Region"], columnField: "Product", valueField: "Sales", aggregator: "COUNT" })
    ).toEqual([
      ["Region", "Hardware", "Software", "Grand Total"],
      ["East", "0", "1", "1"],
      ["West", "1", "0", "1"],
      ["Grand Total", "1", "1", "2"]
    ]);

    expect(
      createPivotTable(sparseRows, { rowFields: ["Region"], columnField: "Product", valueField: "Sales", aggregator: "COUNTNUMS" })
    ).toEqual([
      ["Region", "Hardware", "Software", "Grand Total"],
      ["East", "0", "1", "1"],
      ["West", "1", "0", "1"],
      ["Grand Total", "1", "1", "2"]
    ]);

    expect(
      createPivotTable(sparseRows, { rowFields: ["Region"], columnField: "Product", valueField: "Sales", aggregator: "PRODUCT" })
    ).toEqual([
      ["Region", "Hardware", "Software", "Grand Total"],
      ["East", "", "7", "7"],
      ["West", "10", "", "10"],
      ["Grand Total", "10", "7", "70"]
    ]);
  });

  it("parses scientific-notation values into aggregates", () => {
    const table = createPivotTable(
      [
        ["Region", "Sales"],
        ["West", "1e2"],
        ["West", "2.5E+2"]
      ],
      { rowFields: ["Region"], valueField: "Sales", aggregator: "SUM" }
    );

    expect(table[1]).toEqual(["West", "350"]);
  });
});
