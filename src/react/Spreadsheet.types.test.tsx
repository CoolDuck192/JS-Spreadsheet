import type { CSSProperties } from "react";
import { describe, expectTypeOf, it } from "vitest";
import { createWorkbookSession } from "../core/workbook/WorkbookSession";
import { createBlankWorkbook } from "../lib/workbook";
import { Spreadsheet, type SpreadsheetProps } from "./Spreadsheet";

describe("Spreadsheet public ownership types", () => {
  it("accepts common props in every ownership mode", () => {
    const workbook = createBlankWorkbook();
    const session = createWorkbookSession({ workbook });
    const common = {
      className: "embedded",
      style: { minHeight: 320 } satisfies CSSProperties,
      features: { toolbar: false },
      theme: { accent: "rebeccapurple" },
      onDiagnostic: () => {},
      onCommandResult: () => {},
      onWorkbookChangeEvent: () => {},
      onError: () => {}
    } satisfies Omit<SpreadsheetProps, "session" | "workbook" | "defaultWorkbook" | "onWorkbookChange" | "storage">;

    const sessionBacked: SpreadsheetProps = { ...common, session };
    const controlled: SpreadsheetProps = {
      ...common,
      workbook,
      onWorkbookChange: () => {},
      storage: false
    };
    const uncontrolled: SpreadsheetProps = { ...common, defaultWorkbook: workbook, storage: false };
    const standalone: SpreadsheetProps = {};

    expectTypeOf(sessionBacked).toMatchTypeOf<SpreadsheetProps>();
    expectTypeOf(controlled).toMatchTypeOf<SpreadsheetProps>();
    expectTypeOf(uncontrolled).toMatchTypeOf<SpreadsheetProps>();
    expectTypeOf(standalone).toMatchTypeOf<SpreadsheetProps>();
    void <Spreadsheet {...sessionBacked} />;
    void <Spreadsheet {...controlled} />;
    void <Spreadsheet {...uncontrolled} />;
    session.destroy();
  });

  it("rejects conflicting ownership props", () => {
    const workbook = createBlankWorkbook();
    const session = createWorkbookSession({ workbook });

    // @ts-expect-error a supplied session is the sole workbook authority
    void <Spreadsheet session={session} workbook={workbook} onWorkbookChange={() => {}} />;
    // @ts-expect-error a supplied session cannot be combined with an uncontrolled default
    void <Spreadsheet session={session} defaultWorkbook={workbook} />;
    // @ts-expect-error controlled workbook mode requires acknowledgement
    void <Spreadsheet workbook={workbook} />;
    // @ts-expect-error controlled and uncontrolled workbook values are mutually exclusive
    void <Spreadsheet workbook={workbook} onWorkbookChange={() => {}} defaultWorkbook={workbook} />;
    // @ts-expect-error uncontrolled mode cannot accept the controlled acknowledgement callback
    void <Spreadsheet defaultWorkbook={workbook} onWorkbookChange={() => {}} />;
    // @ts-expect-error session storage is owned by the supplied session
    void <Spreadsheet session={session} storage={false} />;
    // @ts-expect-error unknown root theme tokens are rejected
    void <Spreadsheet theme={{ neon: "lime" }} />;

    session.destroy();
  });
});
