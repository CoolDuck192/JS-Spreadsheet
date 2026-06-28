import { describe, expect, it } from "vitest";
import { getFormulaSuggestions, insertFormulaSuggestion, searchFormulaCatalog } from "./formulaSuggestions";

describe("formulaSuggestions", () => {
  it("shows common functions as soon as the formula starts", () => {
    expect(getFormulaSuggestions("=").map((suggestion) => suggestion.name)).toEqual([
      "SUM",
      "AVERAGE",
      "COUNT",
      "COUNTA",
      "MIN",
      "MAX"
    ]);
  });

  it("shows common functions after a function prefix starts", () => {
    expect(getFormulaSuggestions("=s").map((suggestion) => suggestion.name)).toContain("SUM");
  });

  it("keeps the suggestion set compact", () => {
    expect(getFormulaSuggestions("=c")).toHaveLength(6);
  });

  it("filters suggestions by typed function prefix", () => {
    expect(getFormulaSuggestions("=av").map((suggestion) => suggestion.name)).toEqual([
      "AVERAGE",
      "AVEDEV",
      "AVERAGEA",
      "AVERAGEIF"
    ]);
  });

  it("suggests HyperFormula functions beyond the starter set", () => {
    expect(getFormulaSuggestions("=xl").map((suggestion) => suggestion.name)).toEqual(["XLOOKUP"]);
    expect(getFormulaSuggestions("=networkdays.").map((suggestion) => suggestion.name)).toEqual(["NETWORKDAYS.INTL"]);
  });

  it("searches the broader function catalog by name and description", () => {
    expect(searchFormulaCatalog("lookup").map((suggestion) => suggestion.name)).toEqual(expect.arrayContaining(["VLOOKUP", "XLOOKUP"]));
    expect(searchFormulaCatalog("workdays").map((suggestion) => suggestion.name)).toEqual(
      expect.arrayContaining(["NETWORKDAYS", "NETWORKDAYS.INTL"])
    );
  });

  it("hides function suggestions after a cell reference is typed", () => {
    expect(getFormulaSuggestions("=A1")).toEqual([]);
    expect(getFormulaSuggestions("=$A$1")).toEqual([]);
  });

  it("inserts the selected function call after the equals sign", () => {
    expect(insertFormulaSuggestion("=su", "SUM")).toBe("=SUM(");
    expect(insertFormulaSuggestion("=IF(A1>0,", "IF")).toBe("=IF(");
    expect(insertFormulaSuggestion("=networkdays.", "NETWORKDAYS.INTL")).toBe("=NETWORKDAYS.INTL(");
  });
});
