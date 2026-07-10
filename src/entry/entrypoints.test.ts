// @vitest-environment node
import { readFile } from "node:fs/promises";
import { describe, expect, expectTypeOf, it } from "vitest";
import type {
  WorkbookTableRow as CoreWorkbookTableRow,
  WorkbookTableSession as CoreWorkbookTableSession,
  XlsxImportOptions
} from "./core";
import type {
  WorkbookTableRow as ReactWorkbookTableRow,
  WorkbookTableSession as ReactWorkbookTableSession
} from "./react";

describe("public entrypoints", () => {
  it("keeps the core entrypoint free of React view exports", async () => {
    const core = await import("./core");
    expect(core.createWorkbookSession).toBeTypeOf("function");
    expect(core.createLocalRecordTableSession).toBeTypeOf("function");
    expect(core.createRemoteTableSource).toBeTypeOf("function");
    expect(core.createRemoteTableSession).toBeTypeOf("function");
    expect(core.createWorkbookTableSession).toBeTypeOf("function");
    expect(core.createBlankWorkbook).toBeTypeOf("function");
    expect("Spreadsheet" in core).toBe(false);
    expect("DataTable" in core).toBe(false);
    expect("WorkbookTableView" in core).toBe(false);
  });

  it("publishes the workbook adapter from core and React surfaces", async () => {
    const [core, react] = await Promise.all([import("./core"), import("./react")]);
    expect(react.createWorkbookTableSession).toBe(core.createWorkbookTableSession);
    expect(react.WorkbookTableView).toBeTypeOf("function");
    expectTypeOf<ReactWorkbookTableRow>().toEqualTypeOf<CoreWorkbookTableRow>();
    expectTypeOf<ReactWorkbookTableSession>().toEqualTypeOf<CoreWorkbookTableSession>();
    expectTypeOf<XlsxImportOptions>().toMatchTypeOf<{
      tableKeys?: Readonly<Record<string, { columnName: string }>>;
    }>();
  });

  it("maps typed subpaths for legacy TypeScript resolution", async () => {
    const manifest = JSON.parse(await readFile(
      new URL("../../package.json", import.meta.url),
      "utf8"
    ));

    expect(manifest.typesVersions).toEqual({
      "*": {
        core: ["dist/types/entry/core.d.ts"],
        react: ["dist/types/entry/react.d.ts"],
        "connectors/google": ["dist/types/entry/google.d.ts"]
      }
    });
  });
});
