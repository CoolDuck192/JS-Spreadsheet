import type { CSSProperties } from "react";
import { StrictMode } from "react";
import { act, fireEvent, render, screen, waitFor, within } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import { describe, expect, it, vi } from "vitest";
import { createWorkbookSession } from "./core/workbook/WorkbookSession";
import { createFormulaEngine } from "./lib/formulaEngine";
import { createBlankWorkbook, getCellContent, setCellContent } from "./lib/workbook";
import { Spreadsheet } from "./react/Spreadsheet";

function workbookWith(value: string) {
  const workbook = createBlankWorkbook();
  return setCellContent(workbook, workbook.activeSheetId, "A1", value);
}

describe("controlled Spreadsheet", () => {
  it("renders host rerenders and acknowledges a user workbook edit", async () => {
    const user = userEvent.setup();
    const first = workbookWith("first");
    const second = workbookWith("second");
    const onWorkbookChange = vi.fn();
    const { rerender } = render(
      <Spreadsheet workbook={first} onWorkbookChange={onWorkbookChange} storage={false} />
    );
    expect(screen.getByRole("gridcell", { name: "A1 first" })).toBeInTheDocument();

    rerender(<Spreadsheet workbook={second} onWorkbookChange={onWorkbookChange} storage={false} />);
    expect(await screen.findByRole("gridcell", { name: "A1 second" })).toBeInTheDocument();

    await user.dblClick(screen.getByRole("gridcell", { name: "A1 second" }));
    const editor = screen.getByLabelText("Cell editor A1");
    await user.clear(editor);
    await user.type(editor, "edited{Enter}");
    expect(onWorkbookChange).toHaveBeenCalledTimes(1);
    expect(getCellContent(onWorkbookChange.mock.calls[0][0], second.activeSheetId, "A1")).toBe("edited");
  });

  it("subscribes to a supplied session without owning or overriding it", async () => {
    const user = userEvent.setup();
    const workbook = workbookWith("session");
    const session = createWorkbookSession({ workbook });
    const destroy = vi.spyOn(session, "destroy");
    const externalFactory = vi.fn(createFormulaEngine);
    const onCommandResult = vi.fn();
    const onWorkbookChangeEvent = vi.fn();
    const { unmount } = render(
      <StrictMode>
        <Spreadsheet
          session={session}
          services={{ formulaEngineFactory: externalFactory }}
          onCommandResult={onCommandResult}
          onWorkbookChangeEvent={onWorkbookChangeEvent}
        />
      </StrictMode>
    );
    expect(screen.getByRole("gridcell", { name: "A1 session" })).toBeInTheDocument();

    act(() => {
      session.dispatch({
        type: "cell.set",
        sheetId: workbook.activeSheetId,
        address: "A1",
        input: "shared"
      });
    });
    expect(await screen.findByRole("gridcell", { name: "A1 shared" })).toBeInTheDocument();
    expect(externalFactory).not.toHaveBeenCalled();
    expect(onCommandResult).not.toHaveBeenCalled();
    expect(onWorkbookChangeEvent).toHaveBeenCalledTimes(1);

    await user.dblClick(screen.getByRole("gridcell", { name: "A1 shared" }));
    const editor = screen.getByLabelText("Cell editor A1");
    await user.clear(editor);
    await user.type(editor, "view-command{Enter}");
    expect(onCommandResult).toHaveBeenCalledTimes(1);

    unmount();
    await act(async () => Promise.resolve());
    expect(destroy).not.toHaveBeenCalled();
    session.destroy();
  });

  it("preserves no-prop browser storage behavior", () => {
    localStorage.setItem("javascript-spreadsheet-workbook", JSON.stringify(workbookWith("stored")));
    render(<Spreadsheet />);
    expect(screen.getByRole("gridcell", { name: "A1 stored" })).toBeInTheDocument();
  });

  it("applies root-scoped identity, feature visibility, and theme precedence", () => {
    const style = {
      minHeight: 200,
      "--js-spreadsheet-accent": "orange"
    } as CSSProperties;
    const { container } = render(
      <Spreadsheet
        storage={false}
        className="host-sheet"
        style={style}
        features={{ toolbar: false, formulaBar: false, sheetTabs: false }}
        theme={{ accent: "purple", surface: "ivory" }}
      />
    );
    const root = container.querySelector<HTMLElement>("[data-js-spreadsheet-root='workbook']")!;
    expect(root).toHaveClass("js-spreadsheet-root", "js-spreadsheet-workbook", "host-sheet");
    expect(root.style.getPropertyValue("--js-spreadsheet-accent")).toBe("orange");
    expect(root.style.getPropertyValue("--js-spreadsheet-surface")).toBe("ivory");
    expect(screen.queryByRole("toolbar")).not.toBeInTheDocument();
    expect(screen.queryByLabelText("Formula input")).not.toBeInTheDocument();
    expect(screen.queryByRole("tablist", { name: "Sheet tabs" })).not.toBeInTheDocument();
    expect(screen.getByRole("grid", { name: "Spreadsheet grid" })).toBeInTheDocument();
  });

  it("injects owned-session services and delivers isolated host events once", async () => {
    const user = userEvent.setup();
    const formulaEngineFactory = vi.fn(createFormulaEngine);
    const onWorkbookChange = vi.fn(() => { throw new Error("host callback"); });
    const onWorkbookChangeEvent = vi.fn();
    const onCommandResult = vi.fn();
    const onDiagnostic = vi.fn();
    const workbook = createBlankWorkbook();
    render(
      <Spreadsheet
        workbook={workbook}
        onWorkbookChange={onWorkbookChange}
        storage={false}
        services={{
          formulaEngineFactory,
          now: () => 7,
          createCommandId: () => "host-command"
        }}
        onWorkbookChangeEvent={onWorkbookChangeEvent}
        onCommandResult={onCommandResult}
        onDiagnostic={onDiagnostic}
      />
    );

    await user.dblClick(screen.getByRole("gridcell", { name: "A1" }));
    await user.type(screen.getByLabelText("Cell editor A1"), "value{Enter}");

    expect(formulaEngineFactory).toHaveBeenCalledTimes(1);
    expect(onWorkbookChange).toHaveBeenCalledTimes(1);
    expect(onCommandResult).toHaveBeenCalledTimes(1);
    expect(onCommandResult.mock.calls[0][0].result.status).toBe("committed");
    expect(onWorkbookChangeEvent).toHaveBeenCalledTimes(1);
    expect(onWorkbookChangeEvent.mock.calls[0][0]).toMatchObject({
      previousRevision: "0",
      revision: "1",
      origin: "command",
      commandId: "host-command"
    });
    expect(onDiagnostic).toHaveBeenCalled();
  });

  it("sanitizes invalid service registries and validates them once", async () => {
    const onError = vi.fn();
    const invalidImporters = Object.create(null) as Record<string, { import: () => ReturnType<typeof createBlankWorkbook> }>;
    invalidImporters[" "] = { import: createBlankWorkbook };
    invalidImporters["constructor"] = { import: () => { throw new Error("secret importer value"); } };
    const { rerender } = render(
      <Spreadsheet storage={false} services={{ importers: invalidImporters }} onError={onError} />
    );
    rerender(<Spreadsheet storage={false} services={{ importers: invalidImporters }} onError={onError} />);

    await waitFor(() => expect(onError).toHaveBeenCalledTimes(1));
    expect(onError.mock.calls[0][0]).toMatchObject({ recoverable: true });
    expect(onError.mock.calls[0][0].message).not.toContain("secret");
  });

  it("adapts browser files into the platform-neutral importer contract", async () => {
    const imported = workbookWith("service-import");
    const importer = { import: vi.fn().mockResolvedValue(imported) };
    render(<Spreadsheet storage={false} services={{ importers: { csv: importer } }} />);

    const file = new File(["ignored by host importer"], "host.csv", { type: "text/csv" });
    fireEvent.change(screen.getByLabelText("CSV file"), { target: { files: [file] } });

    expect(await screen.findByRole("gridcell", { name: "A1 service-import" })).toBeInTheDocument();
    expect(importer.import).toHaveBeenCalledWith({
      kind: "text",
      text: "ignored by host importer",
      fileName: "host.csv"
    });
  });

  it("applies individual capability visibility without deleting workbook state", async () => {
    const user = userEvent.setup();
    const workbook = createBlankWorkbook();
    workbook.sheets[0].charts = [{
      id: "chart-1",
      title: "Kept",
      type: "bar",
      range: { start: { row: 0, column: 0 }, end: { row: 1, column: 1 } },
      anchor: { row: 2, column: 2 }
    }];
    render(<Spreadsheet
      defaultWorkbook={workbook}
      storage={false}
      features={{ import: false, export: false, googleSheets: false, charts: false }}
    />);
    await user.click(screen.getByRole("tab", { name: "File" }));
    expect(screen.queryByRole("button", { name: "Import CSV" })).not.toBeInTheDocument();
    expect(screen.queryByRole("button", { name: "Import XLSX" })).not.toBeInTheDocument();
    expect(screen.queryByRole("button", { name: "Export CSV" })).not.toBeInTheDocument();
    expect(screen.queryByRole("button", { name: "Export XLSX" })).not.toBeInTheDocument();
    expect(screen.queryByRole("button", { name: "Link Google Sheet" })).not.toBeInTheDocument();
    expect(screen.queryByLabelText("CSV file")).not.toBeInTheDocument();
    expect(screen.queryByLabelText("XLSX file")).not.toBeInTheDocument();
    await user.click(screen.getByRole("tab", { name: "Insert" }));
    expect(screen.queryByRole("button", { name: "Chart" })).not.toBeInTheDocument();
    expect(workbook.sheets[0].charts).toHaveLength(1);
    expect(within(screen.getByRole("grid", { name: "Spreadsheet grid" })).getAllByRole("gridcell").length).toBeGreaterThan(0);
  });
});
