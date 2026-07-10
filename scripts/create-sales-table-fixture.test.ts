import { execFile } from "node:child_process";
import { mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { resolve } from "node:path";
import { promisify } from "node:util";
import { DOMParser } from "@xmldom/xmldom";
import { strFromU8, unzipSync } from "fflate";
import { expect, it } from "vitest";

const execFileAsync = promisify(execFile);

it("targets the Region filter by its table-column name when columns move", async () => {
  const temporaryDirectory = await mkdtemp(resolve("scripts", ".sales-fixture-test-"));
  const temporaryScript = resolve(temporaryDirectory, "create-sales-table-fixture.mjs");
  const temporaryOutput = resolve(temporaryDirectory, "sales.xlsx");

  try {
    const source = await readFile(resolve("scripts/create-sales-table-fixture.mjs"), "utf8");
    const reordered = source
      .replace(
        'const OUTPUT = resolve("src/test/fixtures/xlsx/generated-sales-structured-table.xlsx");',
        `const OUTPUT = ${JSON.stringify(temporaryOutput)};`
      )
      .replace(
        '    { name: "Order Date", filterButton: true, totalsRowFunction: "none" },\n' +
          '    { name: "Region", filterButton: true, totalsRowFunction: "none" },',
        '    { name: "Region", filterButton: true, totalsRowFunction: "none" },\n' +
          '    { name: "Order Date", filterButton: true, totalsRowFunction: "none" },'
      );
    expect(reordered).not.toBe(source);

    await writeFile(temporaryScript, reordered);
    await execFileAsync(process.execPath, [temporaryScript], { cwd: resolve(".") });

    const archive = unzipSync(await readFile(temporaryOutput));
    const xml = archive["xl/tables/table1.xml"];
    expect(xml).toBeDefined();
    const document = new DOMParser().parseFromString(strFromU8(xml), "application/xml");
    const columns = Array.from(document.getElementsByTagNameNS("*", "tableColumn"));
    const regionIndex = columns.findIndex((column) => column.getAttribute("name") === "Region");
    const filterColumn = Array.from(document.getElementsByTagNameNS("*", "filterColumn"))
      .find((column) => Array.from(column.getElementsByTagNameNS("*", "filter"))
        .map((filter) => filter.getAttribute("val"))
        .join(",") === "East,West");

    expect(regionIndex).toBe(1);
    expect(filterColumn?.getAttribute("colId")).toBe(String(regionIndex));
  } finally {
    await rm(temporaryDirectory, { recursive: true, force: true });
  }
});
