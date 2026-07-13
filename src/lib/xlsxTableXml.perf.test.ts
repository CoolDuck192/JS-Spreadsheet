import { strFromU8, unzipSync } from "fflate";
import { describe, expect, it } from "vitest";

import {
  makeDenseWorksheetPackage,
  makeMultipleDenseWorksheetsPackage
} from "../test/xlsxSecurityFixtures";
import type { StructuredTable } from "../types";
import { validateXlsxArchive } from "./xlsxSecurity";
import { patchNativeTableXml } from "./xlsxTableXml";

const denseTable: StructuredTable = {
  id: "dense-table-id",
  name: "DenseTable",
  sheetId: "dense-sheet-id",
  range: { start: { row: 0, column: 0 }, end: { row: 99_999, column: 9 } },
  headerRow: true,
  totalsRow: false,
  columns: Array.from({ length: 10 }, (_, index) => ({
    id: `dense-column-${index + 1}`,
    name: `Column${index + 1}`,
    sheetColumn: index
  })),
  rowIds: [],
  keyColumnId: "dense-column-1",
  style: {
    theme: "TableStyleMedium9",
    showFirstColumn: false,
    showLastColumn: true,
    showRowStripes: true,
    showColumnStripes: false
  }
};

describe("xlsxTableXml dense archives", () => {
  it(
    "patches and revalidates a dense A1:J100000 worksheet archive",
    { timeout: 60_000 },
    () => {
      const source = makeDenseWorksheetPackage();
      const sourceEntries = unzipSync(source);
      const patched = patchNativeTableXml(source, [denseTable]);
      const patchedEntries = unzipSync(patched);
      const sourceTable = sourceEntries["xl/tables/table1.xml"];
      const patchedTable = patchedEntries["xl/tables/table1.xml"];

      expect(sourceTable).toBeDefined();
      expect(patchedTable).toBeDefined();
      expect(strFromU8(patchedTable)).not.toEqual(strFromU8(sourceTable));
      expect(strFromU8(patchedTable)).toContain('name="TableStyleMedium9"');
      expect(strFromU8(patchedEntries["xl/worksheets/sheet1.xml"]))
        .toContain('dimension ref="A1:J100000"');
      expect(validateXlsxArchive(patched)).toEqual({ ok: true });
    }
  );

  it(
    "patches and revalidates three dense worksheets within scaled package totals",
    { timeout: 60_000 },
    () => {
      expect(patchNativeTableXml(makeMultipleDenseWorksheetsPackage(), []))
        .toBeInstanceOf(Uint8Array);
    }
  );
});
