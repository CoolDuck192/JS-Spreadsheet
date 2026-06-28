import { describe, expect, it } from "vitest";
import { createChartData } from "./charts";

describe("charts", () => {
  it("extracts labeled numeric data with inferred headers and title", () => {
    const chartData = createChartData([
      ["Region", "Sales"],
      ["West", "$1,200"],
      ["East", "800"],
      ["South", "25%"]
    ]);

    expect(chartData).toEqual({
      title: "Sales by Region",
      labelsTitle: "Region",
      valuesTitle: "Sales",
      data: [
        { label: "West", value: 1200 },
        { label: "East", value: 800 },
        { label: "South", value: 25 }
      ]
    });
  });

  it("skips blank labels and nonnumeric values", () => {
    const chartData = createChartData(
      [
        ["Name", "Value"],
        ["One", "10"],
        ["", "20"],
        ["Two", "n/a"],
        ["Three", "30"]
      ],
      "Fallback"
    );

    expect(chartData.data).toEqual([
      { label: "One", value: 10 },
      { label: "Three", value: 30 }
    ]);
    expect(chartData.title).toBe("Value by Name");
  });
});
