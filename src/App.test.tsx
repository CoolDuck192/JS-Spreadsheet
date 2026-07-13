import { readFile } from "node:fs/promises";
import { resolve } from "node:path";
import { act, fireEvent, render, screen, waitFor, within } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import { HyperFormula } from "hyperformula";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import App, { Spreadsheet } from "./App";
import { createWorkbookSession } from "./core/workbook/WorkbookSession";
import { WORKBOOK_STORAGE_KEY } from "./lib/persistence";
import { exportWorkbookToXlsx, importWorkbookFromXlsx } from "./lib/xlsx";
import { createBlankWorkbook, getCellContent, setCellContent } from "./lib/workbook";
import { GOOGLE_CLIENT_ID_STORAGE_KEY } from "./react/browserGoogleClientIdStorage";
import {
  addXlsxZipConsistencyMismatch,
  addMalformedWorkbookXml,
  addSyntheticVbaProject,
  addUnsafeWorkbookDoctype,
  removeXlsxContentTypes,
  replaceNativeTableDocumentRoot
} from "./test/xlsxImportFixtures";

describe("App", () => {
  beforeEach(() => {
    localStorage.clear();
  });

  afterEach(() => {
    vi.restoreAllMocks();
    vi.unstubAllGlobals();
    vi.unstubAllEnvs();
  });

  it("owns the viewport only for the standalone workbook", () => {
    const { container } = render(<App />);
    const root = container.querySelector<HTMLElement>("[data-js-spreadsheet-root='workbook']")!;

    expect(root).toHaveClass("js-spreadsheet-standalone");
    expect(root).toHaveStyle({ height: "100dvh", minHeight: 0 });
  });

  it("shows a recovery warning when the stored workbook cannot be opened", async () => {
    localStorage.setItem(WORKBOOK_STORAGE_KEY, "{invalid workbook");

    render(<App />);

    const alert = await screen.findByRole("alert", { name: /Workbook storage status/ });
    expect(alert).toHaveTextContent(
      "Stored workbook could not be opened. A recovery copy was preserved, and autosave is paused."
    );
    expect(screen.getByLabelText("Status").firstElementChild).toHaveTextContent("Ready");
  });

  it("keeps command feedback visible beside a sanitized autosave alert and clears it after recovery", async () => {
    const user = userEvent.setup();
    let resolveRecovery!: () => void;
    const recovery = new Promise<void>((resolve) => {
      resolveRecovery = resolve;
    });
    const save = vi.fn()
      .mockRejectedValueOnce(new Error("secret quota detail"))
      .mockReturnValueOnce(recovery);
    render(<Spreadsheet storage={{ load: () => null, save }} />);

    await editCell(user, "A1", "unsaved");

    const alert = await screen.findByRole("alert");
    expect(alert).toHaveTextContent("Workbook could not be saved");
    expect(alert).toHaveAccessibleName("Workbook storage status: Workbook could not be saved");
    expect(screen.getByLabelText("Status").firstElementChild).toHaveTextContent("Saved");
    expect(screen.getByLabelText("Status").firstElementChild).not.toBe(alert);
    expect(screen.getByLabelText("Status")).not.toHaveTextContent("secret quota detail");

    await user.click(screen.getByRole("button", { name: "Zoom in" }));
    expect(screen.getByLabelText("Status").firstElementChild).toHaveTextContent("Zoom 125%");
    expect(screen.getByRole("alert")).toBe(alert);

    await editCell(user, "A2", "saved");
    await waitFor(() => expect(save).toHaveBeenCalledTimes(2));
    expect(screen.getByRole("alert")).toBe(alert);
    await act(async () => resolveRecovery());
    await waitFor(() => expect(screen.queryByRole("alert")).not.toBeInTheDocument());
    expect(screen.getByLabelText("Status").firstElementChild).toHaveTextContent("Saved");
  });

  it("always opens standalone Google setup and stores only the normalized public client ID", async () => {
    vi.stubEnv("VITE_GOOGLE_CLIENT_ID", "");
    const prompt = vi.spyOn(window, "prompt");
    const user = userEvent.setup();
    render(<App />);

    await openRibbonTab(user, "File");
    await user.click(screen.getByRole("button", { name: "Import Google Sheet" }));

    expect(screen.getByRole("dialog", { name: "Import Google Sheet" })).toBeInTheDocument();
    const clientId = screen.getByLabelText("Google OAuth client ID");
    await user.type(clientId, " 123-abc.apps.googleusercontent.com ");
    await user.click(screen.getByRole("button", { name: "Save and continue" }));

    await waitFor(() => {
      expect(localStorage.getItem(GOOGLE_CLIENT_ID_STORAGE_KEY)).toBe(
        "123-abc.apps.googleusercontent.com"
      );
    });
    expect(prompt).not.toHaveBeenCalled();
    expect(localStorage.length).toBeGreaterThanOrEqual(1);
  });

  it("imports a Google Sheet through session history and reports Imported", async () => {
    const user = userEvent.setup();
    const tokenProvider = {
      prepare: vi.fn().mockResolvedValue(undefined),
      getAccessToken: vi.fn().mockResolvedValue("test-access-token")
    };
    const fetchMock = vi.fn()
      .mockResolvedValueOnce(new Response(JSON.stringify({
        properties: { title: "Budget" },
        sheets: [{ properties: { title: "Imported" } }]
      }), { status: 200, headers: { "Content-Type": "application/json" } }))
      .mockResolvedValueOnce(new Response(JSON.stringify({
        valueRanges: [{ values: [["from Google"]] }]
      }), { status: 200, headers: { "Content-Type": "application/json" } }));
    vi.stubGlobal("fetch", fetchMock);
    render(
      <Spreadsheet
        defaultWorkbook={setCellContent(createBlankWorkbook(), "sheet-1", "A1", "before")}
        storage={false}
        services={{ googleSheets: { tokenProvider } }}
      />
    );

    await openRibbonTab(user, "File");
    await user.click(screen.getByRole("button", { name: "Import Google Sheet" }));
    const sheetInput = await screen.findByLabelText("Google Sheet URL or spreadsheet ID");
    await user.type(sheetInput, "12345678901234567890");
    await user.click(screen.getByRole("button", { name: "Import and replace workbook" }));

    expect(await screen.findByRole("gridcell", { name: "A1 from Google" })).toBeInTheDocument();
    expect(screen.getByLabelText("Status")).toHaveTextContent("Imported Budget");
    expect(document.body).not.toHaveTextContent("test-access-token");

    await openRibbonTab(user, "Home");
    await user.click(screen.getByRole("button", { name: "Undo" }));
    expect(await screen.findByRole("gridcell", { name: "A1 before" })).toBeInTheDocument();
  });

  it("keeps Google import state open when a destroyed session rejects workbook replacement", async () => {
    const user = userEvent.setup();
    const workbook = setCellContent(createBlankWorkbook(), "sheet-1", "A1", "before");
    const session = createWorkbookSession({ workbook });
    const tokenProvider = {
      prepare: vi.fn().mockResolvedValue(undefined),
      getAccessToken: vi.fn().mockResolvedValue("test-access-token")
    };
    vi.stubGlobal("fetch", vi.fn()
      .mockResolvedValueOnce(new Response(JSON.stringify({
        properties: { title: "Rejected import" },
        sheets: [{ properties: { title: "Imported" } }]
      }), { status: 200, headers: { "Content-Type": "application/json" } }))
      .mockResolvedValueOnce(new Response(JSON.stringify({
        valueRanges: [{ values: [["replacement"]] }]
      }), { status: 200, headers: { "Content-Type": "application/json" } })));
    render(
      <Spreadsheet
        session={session}
        services={{ googleSheets: { tokenProvider } }}
      />
    );
    await openRibbonTab(user, "File");
    await user.click(screen.getByRole("button", { name: "Import Google Sheet" }));
    const sheetInput = await screen.findByLabelText("Google Sheet URL or spreadsheet ID");
    await user.type(sheetInput, "12345678901234567890");
    session.destroy();

    await user.click(screen.getByRole("button", { name: "Import and replace workbook" }));

    await waitFor(() => expect(screen.getByRole("dialog", { name: "Import Google Sheet" })).toBeInTheDocument());
    expect(sheetInput).toHaveValue("12345678901234567890");
    expect(screen.getByLabelText("Status")).not.toHaveTextContent("Imported Rejected import");
    expect(session.getSnapshot().workbook).toBe(workbook);
    expect(getCellContent(session.getSnapshot().workbook, "sheet-1", "A1")).toBe("before");
    expect(session.getSnapshot()).toMatchObject({ canUndo: false, canRedo: false });
  });

  it("edits cells and recalculates formulas", async () => {
    const user = userEvent.setup();
    render(<App />);

    await editCell(user, "A1", "10");
    await editCell(user, "A2", "20");
    await editCell(user, "A3", "=SUM(A1:A2)");

    expect(screen.getByRole("gridcell", { name: "A1 10" })).toHaveTextContent("10");
    expect(screen.getByRole("gridcell", { name: "A3 30" })).toHaveTextContent("30");

    await user.click(screen.getByRole("gridcell", { name: "A3 30" }));
    expect(screen.getByLabelText("Formula input")).toHaveValue("=SUM(A1:A2)");
  });

  it("stores typed grid input as number, boolean, date, and date-time values", async () => {
    const user = userEvent.setup();
    render(<App />);

    await editCell(user, "A1", "1.25e3");
    await editCell(user, "A2", "FALSE");
    await editCell(user, "A3", "2026-01-15");
    await editCell(user, "A4", "2026-01-15T12:00:00Z");

    expect(screen.getByRole("gridcell", { name: "A1 1250" })).toHaveTextContent("1250");
    expect(screen.getByRole("gridcell", { name: "A2 FALSE" })).toHaveTextContent("FALSE");
    expect(screen.getByRole("gridcell", { name: "A3 Jan 15, 2026" })).toHaveTextContent("Jan 15, 2026");
    expect(screen.getByRole("gridcell", { name: "A4 Jan 15, 2026, 12:00 PM" })).toHaveTextContent(
      "Jan 15, 2026, 12:00 PM"
    );

    await user.click(screen.getByRole("gridcell", { name: "A4 Jan 15, 2026, 12:00 PM" }));
    expect(screen.getByLabelText("Number format")).toHaveValue("dateTime");
    expect(screen.getByRole("option", { name: "Date and time" })).toHaveValue("dateTime");

    await user.click(screen.getByRole("gridcell", { name: "A1 1250" }));
    expect(screen.getByLabelText("Formula input")).toHaveValue("1250");

    await waitFor(() => {
      const saved = JSON.parse(localStorage.getItem("javascript-spreadsheet-workbook") ?? "null") as ReturnType<
        typeof createBlankWorkbook
      > | null;
      expect(saved).not.toBeNull();
      const sheetId = saved?.activeSheetId ?? "";
      expect(["A1", "A2", "A3", "A4"].map((address) => getCellContent(saved!, sheetId, address))).toEqual([
        1250,
        false,
        46037,
        46037.5
      ]);
      expect(saved?.sheets[0].formats).toMatchObject({
        A3: { numberFormat: "date" },
        A4: { numberFormat: "dateTime" }
      });
    });
  });

  it("parses typed formula-bar input through the same commit path", async () => {
    const user = userEvent.setup();
    render(<App />);

    const formulaInput = screen.getByLabelText("Formula input");
    await user.click(formulaInput);
    await user.type(formulaInput, "1.25e3{Enter}");

    expect(screen.getByRole("gridcell", { name: "A1 1250" })).toHaveTextContent("1250");
    expect(formulaInput).toHaveValue("1250");
  });

  it("retains both synchronous workbook commits in one React batch", async () => {
    const user = userEvent.setup();
    render(<App />);
    const bold = screen.getByRole("button", { name: "Bold" });
    const italic = screen.getByRole("button", { name: "Italic" });

    act(() => {
      fireEvent.click(bold);
      fireEvent.click(italic);
    });

    expect(bold).toHaveAttribute("aria-pressed", "true");
    expect(italic).toHaveAttribute("aria-pressed", "true");

    await user.click(screen.getByRole("button", { name: "Undo" }));
    expect(bold).toHaveAttribute("aria-pressed", "true");
    expect(italic).toHaveAttribute("aria-pressed", "false");

    await user.click(screen.getByRole("button", { name: "Undo" }));
    expect(bold).toHaveAttribute("aria-pressed", "false");
    expect(italic).toHaveAttribute("aria-pressed", "false");
  });

  it("adds and renames sheets", async () => {
    const user = userEvent.setup();
    vi.spyOn(window, "prompt").mockReturnValue("Budget");

    render(<App />);

    await user.click(screen.getByRole("button", { name: "Add sheet" }));
    expect(screen.getByRole("tab", { name: "Sheet2" })).toHaveAttribute("aria-selected", "true");
    expect(screen.getByRole("gridcell", { name: "A1" })).toHaveFocus();
    expect(screen.getByLabelText("Status").firstElementChild).toHaveTextContent(/^Added Sheet2$/);

    await user.click(screen.getByRole("button", { name: "Rename active sheet" }));
    expect(screen.getByRole("tab", { name: "Budget" })).toBeInTheDocument();
  });

  it("hides and restores sheet tabs from the toolbar", async () => {
    const user = userEvent.setup();
    render(<App />);

    await user.click(screen.getByRole("button", { name: "Add sheet" }));
    expect(screen.getByRole("tab", { name: "Sheet2" })).toHaveAttribute("aria-selected", "true");

    await user.click(screen.getByRole("button", { name: "Hide sheet" }));

    expect(screen.queryByRole("tab", { name: "Sheet2" })).not.toBeInTheDocument();
    expect(screen.getByRole("tab", { name: "Sheet1" })).toHaveAttribute("aria-selected", "true");
    expect(screen.getByLabelText("Status")).toHaveTextContent("Hid Sheet2");

    await user.click(screen.getByRole("button", { name: "Hide sheet" }));

    expect(screen.getByRole("tab", { name: "Sheet1" })).toBeInTheDocument();
    expect(screen.getByLabelText("Status")).toHaveTextContent("Cannot hide the only visible sheet");

    await user.click(screen.getByRole("button", { name: "Unhide sheets" }));

    expect(screen.getByRole("tab", { name: "Sheet2" })).toBeInTheDocument();
    expect(screen.getByLabelText("Status")).toHaveTextContent("Unhid sheets");
  });

  it("moves the active sheet left and right from the toolbar", async () => {
    const user = userEvent.setup();
    render(<App />);

    await user.click(screen.getByRole("button", { name: "Add sheet" }));
    await user.click(screen.getByRole("button", { name: "Add sheet" }));

    expect(sheetTabNames()).toEqual(["Sheet1", "Sheet2", "Sheet3"]);
    expect(screen.getByRole("tab", { name: "Sheet3" })).toHaveAttribute("aria-selected", "true");

    await user.click(screen.getByRole("button", { name: "Move sheet left" }));

    expect(sheetTabNames()).toEqual(["Sheet1", "Sheet3", "Sheet2"]);
    expect(screen.getByRole("tab", { name: "Sheet3" })).toHaveAttribute("aria-selected", "true");
    expect(screen.getByLabelText("Status")).toHaveTextContent("Moved Sheet3 left");

    await user.click(screen.getByRole("button", { name: "Move sheet left" }));
    await user.click(screen.getByRole("button", { name: "Move sheet left" }));

    expect(sheetTabNames()).toEqual(["Sheet3", "Sheet1", "Sheet2"]);
    expect(screen.getByLabelText("Status")).toHaveTextContent("Sheet3 is already first");

    await user.click(screen.getByRole("button", { name: "Move sheet right" }));

    expect(sheetTabNames()).toEqual(["Sheet1", "Sheet3", "Sheet2"]);
    expect(screen.getByLabelText("Status")).toHaveTextContent("Moved Sheet3 right");
  });

  it("toggles worksheet gridlines from the toolbar", async () => {
    const user = userEvent.setup();
    render(<App />);

    const grid = screen.getByRole("grid", { name: "Spreadsheet grid" });

    expect(grid).toHaveAttribute("data-gridlines", "visible");

    await openRibbonTab(user, "View");
    const gridlinesToggle = screen.getByRole("button", { name: "Gridlines" });
    expect(gridlinesToggle).toHaveAttribute("aria-pressed", "true");
    expect(gridlinesToggle).toHaveClass("active-toolbar-button");

    await user.click(gridlinesToggle);

    expect(grid).toHaveAttribute("data-gridlines", "hidden");
    expect(gridlinesToggle).toHaveAttribute("aria-pressed", "false");
    expect(gridlinesToggle).not.toHaveClass("active-toolbar-button");
    expect(screen.getByLabelText("Status")).toHaveTextContent("Gridlines hidden");

    await user.click(gridlinesToggle);

    expect(grid).toHaveAttribute("data-gridlines", "visible");
    expect(gridlinesToggle).toHaveAttribute("aria-pressed", "true");
    expect(screen.getByLabelText("Status")).toHaveTextContent("Gridlines shown");
  });

  it("toggles worksheet headers from the toolbar", async () => {
    const user = userEvent.setup();
    render(<App />);

    const grid = screen.getByRole("grid", { name: "Spreadsheet grid" });

    expect(grid).toHaveAttribute("data-headers", "visible");
    expect(screen.getByRole("columnheader", { name: "Column A" })).toBeInTheDocument();
    expect(screen.getByRole("rowheader", { name: "Row 1" })).toBeInTheDocument();

    await openRibbonTab(user, "View");
    const headersToggle = screen.getByRole("button", { name: "Headers" });
    expect(headersToggle).toHaveAttribute("aria-pressed", "true");

    await user.click(headersToggle);

    expect(grid).toHaveAttribute("data-headers", "hidden");
    expect(screen.queryByRole("columnheader", { name: "Column A" })).not.toBeInTheDocument();
    expect(screen.queryByRole("rowheader", { name: "Row 1" })).not.toBeInTheDocument();
    expect(headersToggle).toHaveAttribute("aria-pressed", "false");
    expect(headersToggle).not.toHaveClass("active-toolbar-button");
    expect(screen.getByLabelText("Status")).toHaveTextContent("Headers hidden");
    expect(screen.getByRole("gridcell", { name: "A1" })).toBeInTheDocument();

    await user.click(headersToggle);

    expect(grid).toHaveAttribute("data-headers", "visible");
    expect(screen.getByRole("columnheader", { name: "Column A" })).toBeInTheDocument();
    expect(screen.getByRole("rowheader", { name: "Row 1" })).toBeInTheDocument();
    expect(screen.getByLabelText("Status")).toHaveTextContent("Headers shown");
  });

  it("toggles the formula bar while preserving direct cell formula editing", async () => {
    const user = userEvent.setup();
    render(<App />);

    expect(screen.getByRole("group", { name: "Formula bar" })).toBeInTheDocument();

    await editCell(user, "A1", "10");
    await openRibbonTab(user, "View");
    const formulaBarToggle = screen.getByRole("button", { name: "Formula bar" });
    expect(formulaBarToggle).toHaveAttribute("aria-pressed", "true");

    await user.click(formulaBarToggle);

    expect(screen.queryByRole("group", { name: "Formula bar" })).not.toBeInTheDocument();
    expect(screen.queryByLabelText("Formula input")).not.toBeInTheDocument();
    expect(formulaBarToggle).toHaveAttribute("aria-pressed", "false");
    expect(formulaBarToggle).not.toHaveClass("active-toolbar-button");
    expect(screen.getByLabelText("Status")).toHaveTextContent("Formula bar hidden");

    await editCell(user, "A2", "=A1*2");

    expect(screen.getByRole("gridcell", { name: "A2 20" })).toHaveTextContent("20");

    await user.click(formulaBarToggle);
    await user.click(screen.getByRole("gridcell", { name: "A2 20" }));

    expect(screen.getByRole("group", { name: "Formula bar" })).toBeInTheDocument();
    expect(screen.getByLabelText("Formula input")).toHaveValue("=A1*2");
    expect(screen.getByLabelText("Status")).toHaveTextContent("Formula bar shown");
  });

  it("shows formula text in worksheet cells without changing formula editing", async () => {
    const user = userEvent.setup();
    render(<App />);

    await editCell(user, "A1", "10");
    await editCell(user, "A2", "20");
    await editCell(user, "A3", "=SUM(A1:A2)");
    await user.click(screen.getByRole("gridcell", { name: "A3 30" }));

    expect(screen.getByRole("gridcell", { name: "A3 30" })).toHaveTextContent("30");
    expect(screen.getByLabelText("Formula input")).toHaveValue("=SUM(A1:A2)");

    await openRibbonTab(user, "View");
    const showFormulasToggle = screen.getByRole("button", { name: "Show formulas" });
    expect(showFormulasToggle).toHaveAttribute("aria-pressed", "false");

    await user.click(showFormulasToggle);

    expect(showFormulasToggle).toHaveAttribute("aria-pressed", "true");
    expect(showFormulasToggle).toHaveClass("active-toolbar-button");
    expect(screen.getByRole("gridcell", { name: "A3 =SUM(A1:A2)" })).toHaveTextContent("=SUM(A1:A2)");
    expect(screen.getByRole("gridcell", { name: "A1 10" })).toHaveTextContent("10");
    expect(screen.getByLabelText("Formula input")).toHaveValue("=SUM(A1:A2)");
    expect(screen.getByLabelText("Status")).toHaveTextContent("Formulas shown");

    await user.click(showFormulasToggle);

    expect(screen.getByRole("gridcell", { name: "A3 30" })).toHaveTextContent("30");
    expect(screen.getByLabelText("Formula input")).toHaveValue("=SUM(A1:A2)");
    expect(screen.getByLabelText("Status")).toHaveTextContent("Formula results shown");
  });

  it("audits formula precedents and dependents from the toolbar", async () => {
    const user = userEvent.setup();
    render(<App />);

    await editCell(user, "A1", "10");
    await editCell(user, "A2", "20");
    await editCell(user, "C1", "5");
    await editCell(user, "D1", "=SUM(A1:A2)+C1");

    await user.click(screen.getByRole("gridcell", { name: "D1 35" }));
    await openRibbonTab(user, "Formulas");
    await user.click(screen.getByRole("button", { name: "Audit formulas" }));

    const panel = screen.getByRole("complementary", { name: "Formula audit" });
    expect(panel).toHaveTextContent("D1");
    expect(panel).toHaveTextContent("=SUM(A1:A2)+C1");
    expect(within(panel).getByRole("button", { name: "Go to A1:A2" })).toBeInTheDocument();
    expect(within(panel).getByRole("button", { name: "Go to C1" })).toBeInTheDocument();
    expect(within(panel).getByText("No dependents")).toBeInTheDocument();
    expect(screen.getByLabelText("Status")).toHaveTextContent("Formula audit opened");

    await user.click(within(panel).getByRole("button", { name: "Go to A1:A2" }));

    expect(screen.getByRole("gridcell", { name: "A1 10" })).toHaveAttribute("aria-selected", "true");
    expect(screen.getByRole("gridcell", { name: "A2 20" })).toHaveAttribute("aria-selected", "true");
    expect(screen.getByLabelText("Status")).toHaveTextContent("Selected A1:A2");
    expect(within(panel).getByRole("button", { name: "Go to D1" })).toBeInTheDocument();
  });

  it("toggles sheet tabs while preserving sheet creation", async () => {
    const user = userEvent.setup();
    render(<App />);

    expect(screen.getByRole("tablist", { name: "Sheet tabs" })).toBeInTheDocument();

    await openRibbonTab(user, "View");
    const sheetTabsToggle = screen.getByRole("button", { name: "Sheet tabs" });
    expect(sheetTabsToggle).toHaveAttribute("aria-pressed", "true");

    await user.click(sheetTabsToggle);

    expect(screen.queryByRole("tablist", { name: "Sheet tabs" })).not.toBeInTheDocument();
    expect(sheetTabsToggle).toHaveAttribute("aria-pressed", "false");
    expect(screen.getByLabelText("Status")).toHaveTextContent("Sheet tabs hidden");

    await openRibbonTab(user, "Home");
    await user.click(screen.getByRole("button", { name: "Add sheet" }));

    expect(screen.queryByRole("tablist", { name: "Sheet tabs" })).not.toBeInTheDocument();
    expect(screen.getByLabelText("Status")).toHaveTextContent("Added Sheet2");

    // Switching ribbon tabs remounts the View controls, so re-query the toggle.
    await openRibbonTab(user, "View");
    await user.click(screen.getByRole("button", { name: "Sheet tabs" }));

    expect(sheetTabNames()).toEqual(["Sheet1", "Sheet2"]);
    expect(screen.getByRole("tab", { name: "Sheet2" })).toHaveAttribute("aria-selected", "true");
    expect(screen.getByLabelText("Status")).toHaveTextContent("Sheet tabs shown");
  });

  it("colors active sheet tabs from the toolbar", async () => {
    const user = userEvent.setup();
    render(<App />);

    fireEvent.input(screen.getByLabelText("Sheet tab color"), { target: { value: "#0f766e" } });

    expect(screen.getByRole("tab", { name: "Sheet1" })).toHaveStyle({ borderTopColor: "#0f766e" });
    expect(screen.getByLabelText("Status")).toHaveTextContent("Changed Sheet1 tab color");

    await user.click(screen.getByRole("button", { name: "Add sheet" }));
    fireEvent.input(screen.getByLabelText("Sheet tab color"), { target: { value: "#7c3aed" } });

    expect(screen.getByRole("tab", { name: "Sheet1" })).toHaveStyle({ borderTopColor: "#0f766e" });
    expect(screen.getByRole("tab", { name: "Sheet2" })).toHaveStyle({ borderTopColor: "#7c3aed" });
    expect(screen.getByLabelText("Status")).toHaveTextContent("Changed Sheet2 tab color");
  });

  it("adds and removes cell comments from the toolbar", async () => {
    const user = userEvent.setup();
    const prompt = vi.spyOn(window, "prompt");
    prompt.mockReturnValueOnce("Follow up with finance").mockReturnValueOnce("");
    render(<App />);

    await user.click(screen.getByRole("gridcell", { name: "A1" }));
    await openRibbonTab(user, "Review");
    await user.click(screen.getByRole("button", { name: "Comment" }));

    expect(screen.getByRole("gridcell", { name: "A1" })).toHaveClass("commented-cell");
    expect(screen.getByRole("gridcell", { name: "A1" })).toHaveAttribute("title", "Comment: Follow up with finance");
    expect(screen.getByLabelText("Status")).toHaveTextContent("Added comment to A1");

    await user.click(screen.getByRole("button", { name: "Comment" }));

    expect(screen.getByRole("gridcell", { name: "A1" })).not.toHaveClass("commented-cell");
    expect(screen.getByRole("gridcell", { name: "A1" })).not.toHaveAttribute("title");
    expect(screen.getByLabelText("Status")).toHaveTextContent("Removed comment from A1");
  });

  it("adds and removes cell links from the toolbar", async () => {
    const user = userEvent.setup();
    vi.spyOn(window, "prompt").mockReturnValueOnce("example.com/report");
    render(<App />);

    await editCell(user, "A1", "Report");
    await user.click(screen.getByRole("gridcell", { name: "A1 Report" }));
    await openRibbonTab(user, "Review");
    await user.click(screen.getByRole("button", { name: "Link" }));

    const link = screen.getByRole("link", { name: "Report" });
    expect(link).toHaveAttribute("href", "https://example.com/report");
    expect(screen.getByRole("gridcell", { name: "A1 Report" })).toHaveClass("hyperlink-cell");
    expect(screen.getByLabelText("Status")).toHaveTextContent("Added link to A1");

    await user.click(screen.getByRole("button", { name: "Unlink" }));

    expect(screen.queryByRole("link", { name: "Report" })).not.toBeInTheDocument();
    expect(screen.getByRole("gridcell", { name: "A1 Report" })).not.toHaveClass("hyperlink-cell");
    expect(screen.getByLabelText("Status")).toHaveTextContent("Removed link from A1");
  });

  it("rejects unsafe cell links", async () => {
    const user = userEvent.setup();
    vi.spyOn(window, "prompt").mockReturnValueOnce("javascript:alert(1)");
    render(<App />);

    await editCell(user, "A1", "Report");
    await openRibbonTab(user, "Review");
    await user.click(screen.getByRole("button", { name: "Link" }));

    expect(screen.queryByRole("link", { name: "Report" })).not.toBeInTheDocument();
    expect(screen.getByLabelText("Status")).toHaveTextContent("Links must use http, https, or mailto");
  });

  it("locks and unlocks selected cells from the toolbar", async () => {
    const user = userEvent.setup();
    render(<App />);

    await editCell(user, "A1", "Locked");
    await user.click(screen.getByRole("gridcell", { name: "A1 Locked" }));
    await openRibbonTab(user, "Review");
    await user.click(screen.getByRole("button", { name: "Lock cells" }));

    expect(screen.getByRole("gridcell", { name: "A1 Locked" })).toHaveClass("read-only-cell");
    expect(screen.getByLabelText("Status")).toHaveTextContent("Locked A1");

    await user.dblClick(screen.getByRole("gridcell", { name: "A1 Locked" }));

    expect(screen.queryByLabelText("Cell editor A1")).not.toBeInTheDocument();
    expect(screen.getByLabelText("Status")).toHaveTextContent("A1 is read-only");

    await user.keyboard("X");

    expect(screen.queryByLabelText("Cell editor A1")).not.toBeInTheDocument();
    expect(screen.getByRole("gridcell", { name: "A1 Locked" })).toHaveTextContent("Locked");

    await user.click(screen.getByRole("button", { name: "Unlock cells" }));
    await editCell(user, "A1", "Editable");

    expect(screen.getByRole("gridcell", { name: "A1 Editable" })).not.toHaveClass("read-only-cell");
    expect(screen.getByRole("gridcell", { name: "A1 Editable" })).toHaveTextContent("Editable");
  });

  it("protects the active sheet while allowing explicitly unlocked cells", async () => {
    const user = userEvent.setup();
    render(<App />);

    await editCell(user, "A1", "Formula");
    await editCell(user, "B1", "Input");
    await user.click(screen.getByRole("gridcell", { name: "B1 Input" }));
    await openRibbonTab(user, "Review");
    await user.click(screen.getByRole("button", { name: "Unlock cells" }));
    await user.click(screen.getByRole("button", { name: "Protect sheet" }));

    expect(screen.getByRole("button", { name: "Protect sheet" })).toHaveClass("active-toolbar-button");
    expect(screen.getByRole("gridcell", { name: "A1 Formula" })).toHaveClass("read-only-cell");
    expect(screen.getByRole("gridcell", { name: "B1 Input" })).not.toHaveClass("read-only-cell");

    await user.dblClick(screen.getByRole("gridcell", { name: "A1 Formula" }));
    expect(screen.queryByLabelText("Cell editor A1")).not.toBeInTheDocument();
    expect(screen.getByLabelText("Status")).toHaveTextContent("A1 is read-only");

    await editCell(user, "B1", "Editable input");
    expect(screen.getByRole("gridcell", { name: "B1 Editable input" })).toHaveTextContent("Editable input");
  });

  it("defines a named range from the name box and uses it in a formula", async () => {
    const user = userEvent.setup();
    render(<App />);

    await editCell(user, "A1", "10");
    await editCell(user, "A2", "20");
    await editCell(user, "A3", "30");
    selectRange("A1 10", "A3 30");

    const nameBox = screen.getByLabelText("Name box");
    expect(nameBox).toHaveValue("A1:A3");
    await user.clear(nameBox);
    await user.type(nameBox, "Sales{Enter}");

    expect(nameBox).toHaveValue("Sales");
    expect(screen.getByLabelText("Status")).toHaveTextContent("Named A1:A3 as Sales");

    await editCell(user, "B1", "=SUM(Sales)");

    expect(screen.getByRole("gridcell", { name: "B1 60" })).toHaveTextContent("60");
  });

  it("manages named ranges from the toolbar", async () => {
    const user = userEvent.setup();
    render(<App />);

    await editCell(user, "A1", "10");
    await editCell(user, "A2", "20");
    await editCell(user, "A3", "30");
    selectRange("A1 10", "A3 30");

    const nameBox = screen.getByLabelText("Name box");
    await user.clear(nameBox);
    await user.type(nameBox, "Sales{Enter}");

    await openRibbonTab(user, "Formulas");
    await user.click(screen.getByRole("button", { name: "Named ranges" }));

    const panel = screen.getByRole("complementary", { name: "Named ranges" });
    expect(within(panel).getByText("Sales")).toBeInTheDocument();
    expect(within(panel).getByText("Sheet1!A1:A3")).toBeInTheDocument();

    await user.click(within(panel).getByRole("button", { name: "Select Sales" }));

    expect(nameBox).toHaveValue("Sales");
    expect(screen.getByLabelText("Status")).toHaveTextContent("Selected Sales");

    await user.click(within(panel).getByRole("button", { name: "Delete Sales" }));

    expect(within(panel).queryByText("Sales")).not.toBeInTheDocument();
    expect(screen.getByLabelText("Status")).toHaveTextContent("Deleted named range Sales");
  });

  it("selects typed cell addresses and ranges from the name box", async () => {
    const user = userEvent.setup();
    render(<App />);

    const nameBox = screen.getByLabelText("Name box");

    await user.clear(nameBox);
    await user.type(nameBox, "C5{Enter}");

    expect(nameBox).toHaveValue("C5");
    expect(screen.getByRole("gridcell", { name: "C5" })).toHaveClass("selected-cell");
    expect(screen.getByLabelText("Status")).toHaveTextContent("Selected C5");

    await user.clear(nameBox);
    await user.type(nameBox, "B2:C3{Enter}");

    expect(nameBox).toHaveValue("B2:C3");
    expect(screen.getByRole("gridcell", { name: "B2" })).toHaveClass("selected-cell");
    expect(screen.getByRole("gridcell", { name: "C3" })).toHaveClass("selected-cell");
    expect(screen.getByLabelText("Status")).toHaveTextContent("Selected B2:C3");
  });

  it("jumps to typed references from the Go To panel", async () => {
    const user = userEvent.setup();
    render(<App />);

    await user.click(screen.getByRole("button", { name: "Go to" }));
    const panel = screen.getByRole("complementary", { name: "Go to" });
    await user.type(within(panel).getByLabelText("Go to reference"), "C5{Enter}");

    expect(screen.getByLabelText("Name box")).toHaveValue("C5");
    expect(screen.getByRole("gridcell", { name: "C5" })).toHaveClass("selected-cell");
    expect(screen.queryByRole("complementary", { name: "Go to" })).not.toBeInTheDocument();
    expect(screen.getByLabelText("Status")).toHaveTextContent("Selected C5");

    await user.click(screen.getByRole("button", { name: "Go to" }));
    await user.type(screen.getByLabelText("Go to reference"), "B2:C3");
    await user.click(screen.getByRole("button", { name: "Go" }));

    expect(screen.getByLabelText("Name box")).toHaveValue("B2:C3");
    expect(screen.getByRole("gridcell", { name: "B2" })).toHaveClass("selected-cell");
    expect(screen.getByRole("gridcell", { name: "C3" })).toHaveClass("selected-cell");
    expect(screen.getByLabelText("Status")).toHaveTextContent("Selected B2:C3");
  });

  it("jumps to named ranges from the Go To panel", async () => {
    const user = userEvent.setup();
    render(<App />);

    await editCell(user, "A1", "10");
    await editCell(user, "A2", "20");
    await editCell(user, "A3", "30");
    selectRange("A1 10", "A3 30");
    await user.clear(screen.getByLabelText("Name box"));
    await user.type(screen.getByLabelText("Name box"), "Sales{Enter}");

    await user.click(screen.getByRole("button", { name: "Go to" }));
    const panel = screen.getByRole("complementary", { name: "Go to" });

    expect(within(panel).getByText("Sales")).toBeInTheDocument();

    await user.click(within(panel).getByRole("button", { name: "Go to Sales" }));

    expect(screen.getByLabelText("Name box")).toHaveValue("Sales");
    expect(screen.getByRole("gridcell", { name: "A1 10" })).toHaveClass("selected-cell");
    expect(screen.getByRole("gridcell", { name: "A3 30" })).toHaveClass("selected-cell");
    expect(screen.getByLabelText("Status")).toHaveTextContent("Selected Sales");
  });

  it("selects full columns, rows, and the sheet from grid headers", async () => {
    const user = userEvent.setup();
    render(<App />);

    await user.click(screen.getByRole("columnheader", { name: "Column B" }));

    expect(screen.getByLabelText("Name box")).toHaveValue("B1:B100");
    expect(screen.getByLabelText("Status")).toHaveTextContent("100 selected");
    expect(screen.getByRole("gridcell", { name: "B1" })).toHaveClass("selected-cell");

    await user.click(screen.getByRole("rowheader", { name: "Row 3" }));

    expect(screen.getByLabelText("Name box")).toHaveValue("A3:Z3");
    expect(screen.getByLabelText("Status")).toHaveTextContent("26 selected");
    expect(screen.getByRole("gridcell", { name: "A3" })).toHaveClass("selected-cell");
    expect(screen.getByRole("gridcell", { name: "Z3" })).toHaveClass("selected-cell");

    await user.click(screen.getByRole("button", { name: "Select sheet" }));

    expect(screen.getByLabelText("Name box")).toHaveValue("A1:Z100");
    expect(screen.getByLabelText("Status")).toHaveTextContent("2600 selected");
  });

  it("hides and unhides selected rows and columns from the toolbar", async () => {
    const user = userEvent.setup();
    render(<App />);

    await editCell(user, "A2", "Hidden row");
    await editCell(user, "B1", "Hidden column");

    await user.click(screen.getByRole("rowheader", { name: "Row 2" }));
    await user.click(screen.getByRole("button", { name: "Hide rows" }));

    expect(screen.queryByRole("rowheader", { name: "Row 2" })).not.toBeInTheDocument();
    expect(screen.queryByRole("gridcell", { name: "A2 Hidden row" })).not.toBeInTheDocument();
    expect(screen.getByLabelText("Status")).toHaveTextContent("Hid 1 row");

    await user.click(screen.getByRole("columnheader", { name: "Column B" }));
    await user.click(screen.getByRole("button", { name: "Hide columns" }));

    expect(screen.queryByRole("columnheader", { name: "Column B" })).not.toBeInTheDocument();
    expect(screen.queryByRole("gridcell", { name: "B1 Hidden column" })).not.toBeInTheDocument();
    expect(screen.getByLabelText("Status")).toHaveTextContent("Hid 1 column");

    await user.click(screen.getByRole("button", { name: "Unhide all" }));

    expect(screen.getByRole("rowheader", { name: "Row 2" })).toBeInTheDocument();
    expect(screen.getByRole("columnheader", { name: "Column B" })).toBeInTheDocument();
    expect(screen.getByRole("gridcell", { name: "A2 Hidden row" })).toBeInTheDocument();
    expect(screen.getByRole("gridcell", { name: "B1 Hidden column" })).toBeInTheDocument();
    expect(screen.getByLabelText("Status")).toHaveTextContent("Unhid rows and columns");
  });

  it("merges and unmerges the selected cells from the toolbar", async () => {
    const user = userEvent.setup();
    render(<App />);

    await editCell(user, "A1", "Quarterly report");
    await editCell(user, "B1", "hidden");
    selectRange("A1 Quarterly report", "B2");

    await user.click(screen.getByRole("button", { name: "Merge cells" }));

    const mergedCell = screen.getByRole("gridcell", { name: "A1 Quarterly report" });
    expect(mergedCell).toHaveClass("merged-cell");
    expect(mergedCell).toHaveAttribute("aria-colspan", "2");
    expect(mergedCell).toHaveAttribute("aria-rowspan", "2");
    expect(screen.getByRole("gridcell", { name: "B1" })).toHaveClass("merge-covered-cell");
    expect(screen.getByLabelText("Status")).toHaveTextContent("Merged A1:B2");

    await user.click(screen.getByRole("button", { name: "Unmerge cells" }));

    expect(screen.getByRole("gridcell", { name: "A1 Quarterly report" })).not.toHaveClass("merged-cell");
    expect(screen.getByRole("gridcell", { name: "B1" })).not.toHaveClass("merge-covered-cell");
    expect(screen.getByLabelText("Status")).toHaveTextContent("Unmerged A1:B2");
  });

  it("pastes tabular data into the active cell", async () => {
    const user = userEvent.setup();
    render(<App />);

    await user.click(screen.getByRole("gridcell", { name: "A1" }));
    fireEvent.paste(screen.getByRole("grid", { name: "Spreadsheet grid" }), {
      clipboardData: {
        getData: () => "left\tright\nbottom-left\tbottom-right"
      }
    });

    expect(screen.getByRole("gridcell", { name: "A1 left" })).toHaveTextContent("left");
    expect(screen.getByRole("gridcell", { name: "B1 right" })).toHaveTextContent("right");
    expect(screen.getByRole("gridcell", { name: "B2 bottom-right" })).toHaveTextContent("bottom-right");
  });

  it("rejects a plain multi-cell paste whose formula is invalid in the final candidate workbook", async () => {
    const user = userEvent.setup();
    render(<App />);

    await editCell(user, "A1", "5");
    await applyNumberValidation(user, "B1", { max: 10 });
    await user.click(screen.getByRole("gridcell", { name: "A1 5" }));

    fireEvent.paste(screen.getByRole("grid", { name: "Spreadsheet grid" }), {
      clipboardData: { getData: () => "20\t=A1" }
    });

    expect(screen.getByRole("gridcell", { name: "A1 5" })).toBeInTheDocument();
    expect(screen.getByRole("gridcell", { name: "B1" })).toHaveTextContent("");
    expect(screen.getByLabelText("Status")).toHaveTextContent("Enter a number less than or equal to 10");
  });

  it("accepts a plain multi-cell paste whose formula becomes valid in the final candidate workbook", async () => {
    const user = userEvent.setup();
    render(<App />);

    await editCell(user, "A1", "20");
    await applyNumberValidation(user, "B1", { max: 10 });
    await user.click(screen.getByRole("gridcell", { name: "A1 20" }));

    fireEvent.paste(screen.getByRole("grid", { name: "Spreadsheet grid" }), {
      clipboardData: { getData: () => "5\t=A1" }
    });

    expect(screen.getByRole("gridcell", { name: "A1 5" })).toBeInTheDocument();
    expect(screen.getByRole("gridcell", { name: "B1 5" })).toBeInTheDocument();
  });

  it("evaluates every validated formula in one engine for a plain paste transaction", async () => {
    const user = userEvent.setup();
    render(<App />);

    await applyNumberValidation(user, "B1", { max: 100 });
    await applyNumberValidation(user, "C1", { max: 100 });
    const buildSpy = vi.spyOn(HyperFormula, "buildFromSheets");
    await user.click(screen.getByRole("gridcell", { name: "A1" }));

    fireEvent.paste(screen.getByRole("grid", { name: "Spreadsheet grid" }), {
      clipboardData: { getData: () => "5\t=A1\t=A1+1" }
    });

    expect(screen.getByRole("gridcell", { name: "B1 5" })).toBeInTheDocument();
    expect(screen.getByRole("gridcell", { name: "C1 6" })).toBeInTheDocument();
    expect(buildSpy).toHaveBeenCalledTimes(1);
  });

  it("pastes an internally copied range with formatting and adjusted formulas", async () => {
    const user = userEvent.setup();
    render(<App />);

    await editCell(user, "A1", "10");
    await editCell(user, "B1", "=A1");
    await user.click(screen.getByRole("gridcell", { name: "B1 10" }));
    await user.click(screen.getByRole("button", { name: "Bold" }));
    await user.keyboard("{Control>}c{/Control}");

    await user.click(screen.getByRole("gridcell", { name: "C1" }));
    fireEvent.paste(screen.getByRole("grid", { name: "Spreadsheet grid" }), {
      clipboardData: {
        getData: () => "=A1"
      }
    });

    const pastedCell = screen.getByRole("gridcell", { name: "C1 10" });
    expect(pastedCell).toHaveStyle({ fontWeight: "700" });
    await user.click(pastedCell);
    expect(screen.getByLabelText("Formula input")).toHaveValue("=B1");
    expect(screen.getByLabelText("Status")).toHaveTextContent("Pasted cells with formatting");
  });

  it("rejects a rich paste whose translated formula is invalid in the final candidate workbook", async () => {
    const user = userEvent.setup();
    render(<App />);

    await editCell(user, "A1", "5");
    await editCell(user, "C1", "20");
    await editCell(user, "D1", "=C1");
    await applyNumberValidation(user, "D1", { max: 10 });
    selectRange("C1 20", "D1 20");
    fireEvent.keyDown(screen.getByRole("grid", { name: "Spreadsheet grid" }), { key: "c", ctrlKey: true });
    await user.click(screen.getByRole("gridcell", { name: "A1 5" }));

    await user.click(screen.getByRole("button", { name: "Paste" }));

    expect(screen.getByRole("gridcell", { name: "A1 5" })).toBeInTheDocument();
    expect(screen.getByRole("gridcell", { name: "B1" })).toHaveTextContent("");
    expect(screen.getByLabelText("Status")).toHaveTextContent("Enter a number less than or equal to 10");
  });

  it("accepts a rich paste whose translated formula becomes valid in the final candidate workbook", async () => {
    const user = userEvent.setup();
    render(<App />);

    await editCell(user, "A1", "20");
    await editCell(user, "C1", "5");
    await editCell(user, "D1", "=C1");
    await applyNumberValidation(user, "D1", { max: 10 });
    selectRange("C1 5", "D1 5");
    fireEvent.keyDown(screen.getByRole("grid", { name: "Spreadsheet grid" }), { key: "c", ctrlKey: true });
    await user.click(screen.getByRole("gridcell", { name: "A1 20" }));

    await user.click(screen.getByRole("button", { name: "Paste" }));

    expect(screen.getByRole("gridcell", { name: "A1 5" })).toBeInTheDocument();
    expect(screen.getByRole("gridcell", { name: "B1 5" })).toBeInTheDocument();
    await user.click(screen.getByRole("gridcell", { name: "B1 5" }));
    expect(screen.getByLabelText("Formula input")).toHaveValue("=A1");
  });

  it("does not build a candidate engine for a rich formula paste without validation rules", async () => {
    const user = userEvent.setup();
    render(<App />);

    await editCell(user, "C1", "=1+1");
    await user.click(screen.getByRole("gridcell", { name: "C1 2" }));
    fireEvent.keyDown(screen.getByRole("grid", { name: "Spreadsheet grid" }), { key: "c", ctrlKey: true });
    await user.click(screen.getByRole("gridcell", { name: "A1" }));
    const buildSpy = vi.spyOn(HyperFormula, "buildFromSheets");

    await user.click(screen.getByRole("button", { name: "Paste" }));

    expect(screen.getByRole("gridcell", { name: "A1 2" })).toBeInTheDocument();
    expect(buildSpy).not.toHaveBeenCalled();
  });

  it("cuts and pastes a rich selection as a move", async () => {
    const user = userEvent.setup();
    render(<App />);

    await editCell(user, "A1", "10");
    await editCell(user, "B1", "=A1");
    await user.click(screen.getByRole("gridcell", { name: "B1 10" }));
    await user.click(screen.getByRole("button", { name: "Bold" }));
    await user.keyboard("{Control>}x{/Control}");

    expect(screen.getByLabelText("Status")).toHaveTextContent("Cut selection");

    await user.click(screen.getByRole("gridcell", { name: "C1" }));
    fireEvent.paste(screen.getByRole("grid", { name: "Spreadsheet grid" }), {
      clipboardData: {
        getData: () => "=A1"
      }
    });

    expect(screen.getByRole("gridcell", { name: "B1" })).toHaveTextContent("");
    const movedCell = screen.getByRole("gridcell", { name: "C1 10" });
    expect(movedCell).toHaveStyle({ fontWeight: "700" });
    await user.click(movedCell);
    expect(screen.getByLabelText("Formula input")).toHaveValue("=A1");
    expect(screen.getByLabelText("Status")).toHaveTextContent("Moved selection");
  });

  it("blocks cut-paste moves when the source becomes read-only before paste", async () => {
    const user = userEvent.setup();
    render(<App />);

    await editCell(user, "A1", "Move me");
    await user.click(screen.getByRole("gridcell", { name: "B1" }));
    await openRibbonTab(user, "Review");
    await user.click(screen.getByRole("button", { name: "Unlock cells" }));
    await user.click(screen.getByRole("gridcell", { name: "A1 Move me" }));
    await user.keyboard("{Control>}x{/Control}");
    await user.click(screen.getByRole("button", { name: "Protect sheet" }));
    await user.click(screen.getByRole("gridcell", { name: "B1" }));

    fireEvent.paste(screen.getByRole("grid", { name: "Spreadsheet grid" }), {
      clipboardData: {
        getData: () => "Move me"
      }
    });

    expect(screen.getByRole("gridcell", { name: "A1 Move me" })).toHaveTextContent("Move me");
    expect(screen.getByRole("gridcell", { name: "B1" })).toHaveTextContent("");
    expect(screen.getByLabelText("Status")).toHaveTextContent("A1 is read-only");

    await user.click(screen.getByRole("button", { name: "Protect sheet" }));
    fireEvent.paste(screen.getByRole("grid", { name: "Spreadsheet grid" }), {
      clipboardData: {
        getData: () => "Move me"
      }
    });

    expect(screen.getByRole("gridcell", { name: "A1" })).toHaveTextContent("");
    expect(screen.getByRole("gridcell", { name: "B1 Move me" })).toHaveTextContent("Move me");
  });

  it("pastes copied values without formulas or source formatting", async () => {
    const user = userEvent.setup();
    render(<App />);

    await editCell(user, "A1", "10");
    await editCell(user, "B1", "=A1");
    await user.click(screen.getByRole("gridcell", { name: "B1 10" }));
    await user.click(screen.getByRole("button", { name: "Bold" }));
    await user.keyboard("{Control>}c{/Control}");

    await user.click(screen.getByRole("gridcell", { name: "C1" }));
    await user.click(screen.getByRole("button", { name: "Paste options" }));
    await user.click(screen.getByRole("menuitem", { name: "Paste values" }));

    const pastedCell = screen.getByRole("gridcell", { name: "C1 10" });
    expect(pastedCell).not.toHaveStyle({ fontWeight: "700" });
    await user.click(pastedCell);
    expect(screen.getByLabelText("Formula input")).toHaveValue("10");
    expect(screen.getByLabelText("Status")).toHaveTextContent("Pasted values");
  });

  it("pastes copied formats without replacing target content", async () => {
    const user = userEvent.setup();
    render(<App />);

    await editCell(user, "A1", "Styled");
    await user.click(screen.getByRole("gridcell", { name: "A1 Styled" }));
    await user.click(screen.getByRole("button", { name: "Bold" }));
    await user.keyboard("{Control>}c{/Control}");

    await editCell(user, "C1", "Keep content");
    await user.click(screen.getByRole("gridcell", { name: "C1 Keep content" }));
    await user.click(screen.getByRole("button", { name: "Paste options" }));
    await user.click(screen.getByRole("menuitem", { name: "Paste formats" }));

    const formattedCell = screen.getByRole("gridcell", { name: "C1 Keep content" });
    expect(formattedCell).toHaveStyle({ fontWeight: "700" });
    expect(screen.getByLabelText("Status")).toHaveTextContent("Pasted formats");
  });

  it("transposes a copied range from the toolbar", async () => {
    const user = userEvent.setup();
    render(<App />);

    await user.click(screen.getByRole("gridcell", { name: "A1" }));
    fireEvent.paste(screen.getByRole("grid", { name: "Spreadsheet grid" }), {
      clipboardData: {
        getData: () => "A\tB\nC\tD"
      }
    });

    selectRange("A1 A", "B2 D");
    await user.keyboard("{Control>}c{/Control}");
    await user.click(screen.getByRole("gridcell", { name: "D1" }));
    await user.click(screen.getByRole("button", { name: "Paste options" }));
    await user.click(screen.getByRole("menuitem", { name: "Transpose paste" }));

    expect(screen.getByRole("gridcell", { name: "D1 A" })).toHaveTextContent("A");
    expect(screen.getByRole("gridcell", { name: "E1 C" })).toHaveTextContent("C");
    expect(screen.getByRole("gridcell", { name: "D2 B" })).toHaveTextContent("B");
    expect(screen.getByRole("gridcell", { name: "E2 D" })).toHaveTextContent("D");
    expect(screen.getByLabelText("Status")).toHaveTextContent("Transposed paste");
  });

  it("organizes commands into Excel-style ribbon tabs", async () => {
    const user = userEvent.setup();
    render(<App />);

    const toolbar = screen.getByRole("toolbar", { name: "Toolbar" });
    const tabs = within(toolbar).getByRole("tablist", { name: "Ribbon tabs" });
    expect(within(tabs).getAllByRole("tab").map((tab) => tab.textContent)).toEqual([
      "File",
      "Home",
      "Insert",
      "Formulas",
      "Data",
      "Review",
      "View"
    ]);
    expect(within(tabs).getByRole("tab", { name: "Home" })).toHaveAttribute("aria-selected", "true");
    expect(within(toolbar).getByRole("tabpanel", { name: "Home" })).toBeInTheDocument();
    expect(within(toolbar).getByRole("group", { name: "Clipboard" })).toBeInTheDocument();
    expect(within(toolbar).getByRole("group", { name: "Font" })).toBeInTheDocument();
    expect(within(toolbar).queryByRole("button", { name: "Import CSV" })).not.toBeInTheDocument();

    await user.click(within(tabs).getByRole("tab", { name: "File" }));

    expect(within(tabs).getByRole("tab", { name: "File" })).toHaveAttribute("aria-selected", "true");
    expect(within(toolbar).getByRole("tabpanel", { name: "File" })).toBeInTheDocument();
    expect(within(toolbar).getByRole("group", { name: "Workbook" })).toBeInTheDocument();
    expect(within(toolbar).getByRole("button", { name: "Import CSV" })).toBeInTheDocument();
    expect(within(toolbar).getByRole("button", { name: "Export CSV" })).toBeInTheDocument();
    expect(within(toolbar).getByRole("button", { name: "Import XLSX" })).toBeInTheDocument();
    expect(within(toolbar).getByRole("button", { name: "Export XLSX" })).toBeInTheDocument();
    expect(within(toolbar).queryByRole("button", { name: "Bold" })).not.toBeInTheDocument();

    await user.click(within(tabs).getByRole("tab", { name: "Insert" }));

    expect(within(toolbar).getByRole("tabpanel", { name: "Insert" })).toBeInTheDocument();
    expect(within(toolbar).getByRole("button", { name: "Chart" })).toBeInTheDocument();
    expect(within(toolbar).getByRole("button", { name: "Pivot table" })).toBeInTheDocument();
  });

  it("creates a structured table from the Insert ribbon and converts it back without changing values", async () => {
    const user = userEvent.setup();
    const onCommandResult = vi.fn();
    render(<Spreadsheet storage={false} onCommandResult={onCommandResult} />);

    await user.click(screen.getByRole("gridcell", { name: "A1" }));
    fireEvent.paste(screen.getByRole("grid", { name: "Spreadsheet grid" }), {
      clipboardData: { getData: () => "Region\tSales\nWest\t10\nEast\t8" }
    });
    selectRange("A1 Region", "B3 8");

    expect(screen.queryByRole("tab", { name: "Table" })).not.toBeInTheDocument();
    await openRibbonTab(user, "Insert");
    await user.click(screen.getByRole("button", { name: "Table" }));

    expect(screen.getByRole("tab", { name: "Table" })).toBeInTheDocument();
    expect(onCommandResult).toHaveBeenCalledWith(expect.objectContaining({
      command: expect.objectContaining({
        type: "transaction",
        commands: expect.arrayContaining([expect.objectContaining({ type: "table.create" })])
      }),
      result: expect.objectContaining({ status: "committed" })
    }));

    await openRibbonTab(user, "Table");
    await user.click(screen.getByRole("button", { name: "Convert to range" }));

    expect(screen.queryByRole("tab", { name: "Table" })).not.toBeInTheDocument();
    expect(screen.getByRole("gridcell", { name: "A1 Region" })).toHaveTextContent("Region");
    expect(screen.getByRole("gridcell", { name: "B2 10" })).toHaveTextContent("10");
    expect(screen.getByRole("gridcell", { name: "B3 8" })).toHaveTextContent("8");
    expect(onCommandResult).toHaveBeenCalledWith(expect.objectContaining({
      command: expect.objectContaining({ type: "table.convertToRange" }),
      result: expect.objectContaining({ status: "committed" })
    }));
  });

  it("uses the active cell's table when a large selection intersects two tables", async () => {
    const user = userEvent.setup();
    render(<Spreadsheet defaultWorkbook={structuredTableWorkbook()} storage={false} />);

    selectRange("D1 Department", "A1 Region");

    await openRibbonTab(user, "Table");
    expect(screen.getByLabelText("Table name")).toHaveValue("CostsTable");
  });

  it("dispatches table metadata edits through the workbook session and announces rejected drafts", async () => {
    const user = userEvent.setup();
    const onCommandResult = vi.fn();
    render(<Spreadsheet defaultWorkbook={structuredTableWorkbook()} storage={false} onCommandResult={onCommandResult} />);

    await openRibbonTab(user, "Table");
    const name = screen.getByLabelText("Table name");
    await user.clear(name);
    await user.type(name, "RevenueTable{Enter}");

    expect(onCommandResult).toHaveBeenCalledWith(expect.objectContaining({
      command: { type: "table.rename", tableId: "table-sales", name: "RevenueTable" },
      result: expect.objectContaining({ status: "committed" })
    }));

    await user.clear(screen.getByLabelText("Table name"));
    await user.type(screen.getByLabelText("Table name"), "CostsTable{Enter}");

    expect(screen.getByRole("alert")).toHaveTextContent("Table names must be unique");
    expect(screen.getByLabelText("Table name")).toHaveFocus();
    expect(screen.getByLabelText("Status")).toHaveTextContent("Table names must be unique");
  });

  it("downloads the contextual table as a native XLSX artifact and revokes its URL", async () => {
    const user = userEvent.setup();
    const createObjectURL = vi.spyOn(URL, "createObjectURL").mockReturnValue("blob:table-xlsx");
    const revokeObjectURL = vi.spyOn(URL, "revokeObjectURL").mockImplementation(() => undefined);
    const originalCreateElement = document.createElement.bind(document);
    const anchors: HTMLAnchorElement[] = [];
    vi.spyOn(document, "createElement").mockImplementation((tagName) => {
      const element = originalCreateElement(tagName);
      if (tagName.toLowerCase() === "a") {
        vi.spyOn(element, "click").mockImplementation(() => undefined);
        anchors.push(element as HTMLAnchorElement);
      }
      return element;
    });
    render(<Spreadsheet defaultWorkbook={structuredTableWorkbook()} storage={false} />);

    await openRibbonTab(user, "Table");
    const exportButton = screen.getByRole("button", { name: "Export table" });
    expect(exportButton).toBeEnabled();
    await user.click(exportButton);

    await waitFor(() => expect(createObjectURL).toHaveBeenCalledOnce());
    const blob = createObjectURL.mock.calls[0][0] as Blob;
    expect(blob.type).toBe("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    const imported = await importWorkbookFromXlsx(await blob.arrayBuffer());
    expect(imported.tables).toHaveLength(1);
    expect(imported.tables[0].name).toBe("SalesTable");
    expect(anchors[0].download).toBe("SalesTable.xlsx");
    expect(revokeObjectURL).toHaveBeenCalledWith("blob:table-xlsx");
    expect(screen.getByLabelText("Status")).toHaveTextContent("Exported SalesTable");
  });

  it("revokes the contextual table URL and reports a download failure when the click throws", async () => {
    const user = userEvent.setup();
    const onError = vi.fn();
    vi.spyOn(URL, "createObjectURL").mockReturnValue("blob:table-failure");
    const revokeObjectURL = vi.spyOn(URL, "revokeObjectURL").mockImplementation(() => undefined);
    const originalCreateElement = document.createElement.bind(document);
    vi.spyOn(document, "createElement").mockImplementation((tagName) => {
      const element = originalCreateElement(tagName);
      if (tagName.toLowerCase() === "a") {
        vi.spyOn(element, "click").mockImplementation(() => {
          throw new Error("blocked download");
        });
      }
      return element;
    });
    render(<Spreadsheet defaultWorkbook={structuredTableWorkbook()} storage={false} onError={onError} />);

    await openRibbonTab(user, "Table");
    await user.click(screen.getByRole("button", { name: "Export table" }));

    await waitFor(() => expect(screen.getByLabelText("Status")).toHaveTextContent("Table export failed"));
    expect(revokeObjectURL).toHaveBeenCalledWith("blob:table-failure");
    expect(onError).toHaveBeenCalledWith(expect.objectContaining({
      code: "service.export.table.failed",
      message: "Table export failed",
      recoverable: true
    }));
  });

  it("keeps grid focus when the contextual Table tab disappears after a cell selection", async () => {
    const user = userEvent.setup();
    render(<Spreadsheet defaultWorkbook={structuredTableWorkbook()} storage={false} />);

    const homeTab = screen.getByRole("tab", { name: "Home" });
    homeTab.focus();
    await user.keyboard("{End}");
    expect(screen.getByRole("tab", { name: "Table" })).toHaveAttribute("aria-selected", "true");
    expect(screen.getByRole("tab", { name: "Table" })).toHaveFocus();

    await user.click(screen.getByRole("gridcell", { name: "H10" }));
    expect(screen.queryByRole("tab", { name: "Table" })).not.toBeInTheDocument();
    expect(screen.getByRole("tab", { name: "Home" })).toHaveAttribute("aria-selected", "true");
    expect(screen.getByRole("gridcell", { name: "H10" })).toHaveFocus();
  });

  it("restores Home focus when a focused contextual control tears down the Table tab", async () => {
    const user = userEvent.setup();
    render(<Spreadsheet defaultWorkbook={structuredTableWorkbook()} storage={false} />);

    await openRibbonTab(user, "Table");
    await user.click(screen.getByRole("button", { name: "Convert to range" }));

    expect(screen.queryByRole("tab", { name: "Table" })).not.toBeInTheDocument();
    expect(screen.getByRole("tab", { name: "Home" })).toHaveFocus();
  });

  it("opens a live adaptive table and returns the stable cell selection to the spreadsheet", async () => {
    const user = userEvent.setup();
    render(<Spreadsheet defaultWorkbook={structuredTableWorkbook()} storage={false} />);

    await openRibbonTab(user, "Table");
    await user.click(screen.getByRole("button", { name: "Open table view" }));
    expect(screen.getByRole("dialog", { name: "Workbook table table-sales" })).toBeInTheDocument();

    const tableCell = screen.getByRole("gridcell", { name: "sales-west Sales" });
    expect(tableCell).toHaveTextContent("10");
    await user.click(tableCell);
    await user.click(screen.getByRole("button", { name: "Open in Spreadsheet" }));

    expect(screen.queryByRole("dialog", { name: "Workbook table table-sales" })).not.toBeInTheDocument();
    expect(screen.getByRole("gridcell", { name: "B2 10" })).toHaveAttribute("aria-selected", "true");
    await waitFor(() => {
      expect(screen.getByRole("grid", { name: "Spreadsheet grid" })).toHaveFocus();
    });
  });

  it("opens saved workbooks with stale active sheet ids", () => {
    localStorage.setItem(
      "javascript-spreadsheet-workbook",
      JSON.stringify({
        version: 1,
        activeSheetId: "missing-sheet",
        namedRanges: [],
        sheets: [
          {
            id: "sheet-1",
            name: "Recovered",
            rowCount: 100,
            columnCount: 26,
            cells: { A1: "still here" }
          }
        ]
      })
    );

    render(<App />);

    expect(screen.getByRole("tab", { name: "Recovered" })).toHaveAttribute("aria-selected", "true");
    expect(screen.getByRole("gridcell", { name: "A1 still here" })).toBeInTheDocument();
  });

  it("resets workbook history and stale cut state when creating a new workbook", async () => {
    const user = userEvent.setup();
    render(<App />);

    await editCell(user, "A1", "discard me");
    await user.click(screen.getByRole("gridcell", { name: "A1 discard me" }));
    await user.keyboard("{Control>}x{/Control}");
    await openRibbonTab(user, "File");
    await user.click(screen.getByRole("button", { name: "New workbook" }));

    expect(screen.getByRole("gridcell", { name: "A1" })).toHaveTextContent("");
    await openRibbonTab(user, "Home");
    expect(screen.getByRole("button", { name: "Undo" })).toBeDisabled();
    expect(screen.getByRole("button", { name: "Paste" })).toBeDisabled();
    expect(screen.getByLabelText("Status")).toHaveTextContent("New workbook");
  });

  it("imports CSV data and resets the file picker", async () => {
    render(<App />);

    const fileInput = screen.getByLabelText("CSV file") as HTMLInputElement;
    const file = new File(["Name,Amount\nRent,1200"], "budget.csv", { type: "text/csv" });
    Object.defineProperty(fileInput, "value", {
      configurable: true,
      value: "C:\\fakepath\\budget.csv",
      writable: true
    });

    fireEvent.change(fileInput, { target: { files: [file] } });

    expect(fileInput.value).toBe("");
    expect(await screen.findByRole("gridcell", { name: "A1 Name" })).toHaveTextContent("Name");
    expect(await screen.findByRole("gridcell", { name: "B2 1200" })).toHaveTextContent("1200");
  });

  it("imports XLSX workbook data and resets the file picker", async () => {
    render(<App />);

    let workbook = createBlankWorkbook();
    workbook = setCellContent(workbook, workbook.activeSheetId, "A1", "Project");
    workbook = setCellContent(workbook, workbook.activeSheetId, "B1", "=LEN(A1)");
    const bytes = await exportWorkbookToXlsx(workbook);
    const fileInput = screen.getByLabelText("XLSX file") as HTMLInputElement;
    const file = new File([bytes], "budget.xlsx", {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    });
    Object.defineProperty(fileInput, "value", {
      configurable: true,
      value: "C:\\fakepath\\budget.xlsx",
      writable: true
    });

    fireEvent.change(fileInput, { target: { files: [file] } });

    expect(fileInput.value).toBe("");
    expect(await screen.findByRole("gridcell", { name: "A1 Project" })).toHaveTextContent("Project");
    await userEvent.setup().click(screen.getByRole("gridcell", { name: "B1 7" }));
    expect(screen.getByLabelText("Formula input")).toHaveValue("=LEN(A1)");
    expect(screen.getByLabelText("Status")).toHaveTextContent("Imported budget.xlsx");
  });

  it("keeps drop guidance for unsupported TXT files without invoking the XLSX importer", () => {
    const importer = { import: vi.fn().mockResolvedValue(createBlankWorkbook()) };
    const { container } = render(
      <Spreadsheet storage={false} services={{ importers: { xlsx: importer } }} />
    );
    const shell = container.querySelector("main.app-shell")!;

    fireEvent.drop(shell, {
      dataTransfer: { files: [new File(["unsupported"], "notes.txt")] }
    });

    expect(screen.getByLabelText("Status").firstElementChild).toHaveTextContent(
      /^Drop an \.xlsx, \.xlsm, or \.csv file to import it$/
    );
    expect(importer.import).not.toHaveBeenCalled();
  });

  it("uses picker guidance for unsupported TXT files, resets the picker, and skips the XLSX importer", () => {
    const importer = { import: vi.fn().mockResolvedValue(createBlankWorkbook()) };
    render(<Spreadsheet storage={false} services={{ importers: { xlsx: importer } }} />);
    const fileInput = screen.getByLabelText("XLSX file") as HTMLInputElement;
    Object.defineProperty(fileInput, "value", {
      configurable: true,
      value: "C:\\fakepath\\notes.txt",
      writable: true
    });

    fireEvent.change(fileInput, {
      target: { files: [new File(["unsupported"], "notes.txt")] }
    });

    expect(screen.getByLabelText("Status").firstElementChild).toHaveTextContent(
      /^Select an \.xlsx, \.xlsm, or \.csv file to import it$/
    );
    expect(fileInput.value).toBe("");
    expect(importer.import).not.toHaveBeenCalled();
  });

  it("explains that dropped legacy XLS files must be re-saved", () => {
    const { container } = render(<App />);
    const shell = container.querySelector("main.app-shell")!;

    fireEvent.drop(shell, {
      dataTransfer: { files: [new File(["legacy"], "budget.xls")] }
    });

    expect(screen.getByLabelText("Status")).toHaveTextContent(
      "Legacy .xls files aren't supported. Re-save the workbook as .xlsx and try again."
    );
  });

  it("explains that legacy XLS files selected from the picker must be re-saved", async () => {
    const importer = { import: vi.fn().mockResolvedValue(createBlankWorkbook()) };
    render(<Spreadsheet storage={false} services={{ importers: { xlsx: importer } }} />);

    fireEvent.change(screen.getByLabelText("XLSX file"), {
      target: { files: [new File(["legacy"], "budget.xls")] }
    });

    await waitFor(() => expect(screen.getByLabelText("Status")).toHaveTextContent(
      "Legacy .xls files aren't supported. Re-save the workbook as .xlsx and try again."
    ));
    expect(importer.import).not.toHaveBeenCalled();
  });

  it("imports dropped XLSM workbooks through the sanitizer and reports ignored macros", async () => {
    const workbook = setCellContent(createBlankWorkbook(), "sheet-1", "A1", "macro-safe");
    const bytes = addSyntheticVbaProject(await exportWorkbookToXlsx(workbook));
    const file = new File([bytes], "macro-budget.xlsm", {
      type: "application/vnd.ms-excel.sheet.macroEnabled.12"
    });
    const { container } = render(<App />);
    const shell = container.querySelector("main.app-shell")!;

    fireEvent.drop(shell, { dataTransfer: { files: [file] } });

    expect(await screen.findByRole("gridcell", { name: "A1 macro-safe" })).toBeInTheDocument();
    expect(screen.getByLabelText("Status")).toHaveTextContent(
      "Imported macro-budget.xlsm (macros ignored)"
    );
    expect(screen.getByLabelText("XLSX file")).toHaveAttribute(
      "accept",
      expect.stringContaining(".xlsm")
    );
  });

  it("leaves XLSM macro policy and status to a host-provided importer", async () => {
    const importer = { import: vi.fn().mockResolvedValue(createBlankWorkbook()) };
    render(<Spreadsheet storage={false} services={{ importers: { xlsx: importer } }} />);

    fireEvent.change(screen.getByLabelText("XLSX file"), {
      target: { files: [new File(["host-owned"], "host-budget.xlsm")] }
    });

    await waitFor(() => expect(screen.getByLabelText("Status")).toHaveTextContent(
      "Imported host-budget.xlsm"
    ));
    expect(screen.getByLabelText("Status")).not.toHaveTextContent("macros ignored");
    expect(importer.import).toHaveBeenCalledWith(expect.objectContaining({
      fileName: "host-budget.xlsm"
    }));
  });

  it("surfaces typed XLSX security rejections and logs the underlying error", async () => {
    const consoleError = vi.spyOn(console, "error").mockImplementation(() => undefined);
    const unsafe = addUnsafeWorkbookDoctype(await exportWorkbookToXlsx(createBlankWorkbook()));
    render(<App />);

    fireEvent.change(screen.getByLabelText("XLSX file"), {
      target: { files: [new File([unsafe], "unsafe.xlsx")] }
    });

    await waitFor(() => expect(screen.getByLabelText("Status")).toHaveTextContent(
      "XLSX import rejected by security checks"
    ));
    expect(consoleError).toHaveBeenCalledWith("XLSX import failed", expect.objectContaining({
      issue: expect.objectContaining({ code: expect.stringMatching(/^XLSX_/) })
    }));
  });

  it("surfaces native-table XML rejections as typed security failures", async () => {
    const consoleError = vi.spyOn(console, "error").mockImplementation(() => undefined);
    const source = await readFile(resolve("src/test/fixtures/xlsx/variant-table.xlsx"));
    const unsafe = replaceNativeTableDocumentRoot(
      new Uint8Array(source.buffer, source.byteOffset, source.byteLength)
    );
    render(<App />);

    fireEvent.change(screen.getByLabelText("XLSX file"), {
      target: { files: [new File([unsafe], "unsafe-table.xlsx")] }
    });

    await waitFor(() => expect(screen.getByLabelText("Status")).toHaveTextContent(
      "XLSX import rejected by security checks"
    ));
    expect(consoleError).toHaveBeenCalledWith("XLSX import failed", expect.objectContaining({
      code: "XLSX_XML_UNSAFE",
      issue: {
        code: "XLSX_XML_UNSAFE",
        message: "Expected a native table document"
      }
    }));
  });

  it.each([
    ["invalid local header", "invalidLocalHeader", /^Invalid local ZIP header for /],
    ["inconsistent local header", "inconsistentLocalHeader", /^Inconsistent local ZIP header for /],
    ["inconsistent entry name", "inconsistentEntryName", /^Inconsistent ZIP entry name for /],
    [
      "multiline inconsistent entry name",
      "inconsistentEntryNameWithNewline",
      /^Inconsistent ZIP entry name for \n/
    ],
    ["inconsistent sizes", "inconsistentSizes", /^Inconsistent ZIP sizes for /],
    ["invalid stored sizes", "invalidStoredSizes", /^Invalid stored ZIP sizes for /]
  ] as const)("routes a ZIP %s mismatch to security rejection", async (_label, mismatch, message) => {
    const consoleError = vi.spyOn(console, "error").mockImplementation(() => undefined);
    const unsafe = addXlsxZipConsistencyMismatch(
      await exportWorkbookToXlsx(createBlankWorkbook()),
      mismatch
    );
    render(<App />);

    fireEvent.change(screen.getByLabelText("XLSX file"), {
      target: { files: [new File([unsafe], "inconsistent.xlsx")] }
    });

    await waitFor(() => expect(screen.getByLabelText("Status").children[0]).toHaveTextContent(
      /^XLSX import rejected by security checks$/
    ));
    expect(consoleError).toHaveBeenCalledWith("XLSX import failed", expect.objectContaining({
      code: "XLSX_ARCHIVE_LIMIT",
      issue: expect.objectContaining({
        code: "XLSX_ARCHIVE_LIMIT",
        message: expect.stringMatching(message)
      })
    }));
  });

  it("identifies corrupt XLSX archives separately from security rejections", async () => {
    const consoleError = vi.spyOn(console, "error").mockImplementation(() => undefined);
    render(<App />);

    fireEvent.change(screen.getByLabelText("XLSX file"), {
      target: { files: [new File([new Uint8Array([1, 2, 3])], "corrupt.xlsx")] }
    });

    await waitFor(() => expect(screen.getByLabelText("Status")).toHaveTextContent(
      "Could not read XLSX file; it may be corrupt or unreadable"
    ));
    expect(consoleError).toHaveBeenCalledWith("XLSX import failed", expect.objectContaining({
      issue: expect.objectContaining({ code: "XLSX_ARCHIVE_LIMIT" })
    }));
  });

  it.each([
    ["malformed workbook XML", addMalformedWorkbookXml, "XLSX_XML_UNSAFE"],
    ["missing package content types", removeXlsxContentTypes, "XLSX_RELATIONSHIP_INVALID"]
  ] as const)("identifies %s as a corrupt workbook", async (_label, mutate, issueCode) => {
    const consoleError = vi.spyOn(console, "error").mockImplementation(() => undefined);
    const corrupt = mutate(await exportWorkbookToXlsx(createBlankWorkbook()));
    render(<App />);

    fireEvent.change(screen.getByLabelText("XLSX file"), {
      target: { files: [new File([corrupt], "corrupt-package.xlsx")] }
    });

    await waitFor(() => expect(screen.getByLabelText("Status")).toHaveTextContent(
      "Could not read XLSX file; it may be corrupt or unreadable"
    ));
    expect(consoleError).toHaveBeenCalledWith("XLSX import failed", expect.objectContaining({
      issue: expect.objectContaining({ code: issueCode })
    }));
  });

  it("distinguishes unreadable XLSX files from parser failures", async () => {
    const consoleError = vi.spyOn(console, "error").mockImplementation(() => undefined);
    const file = new File(["unreadable"], "unreadable.xlsx");
    Object.defineProperty(file, "arrayBuffer", {
      value: vi.fn().mockRejectedValue(new DOMException("read failed", "NotReadableError"))
    });
    render(<App />);

    fireEvent.change(screen.getByLabelText("XLSX file"), { target: { files: [file] } });

    await waitFor(() => expect(screen.getByLabelText("Status")).toHaveTextContent(
      "Could not read XLSX file; it may be corrupt or unreadable"
    ));
    expect(consoleError).toHaveBeenCalledWith("XLSX import failed", expect.any(DOMException));
  });

  it("distinguishes unexpected XLSX parser failures", async () => {
    const failure = new Error("exceljs internal detail");
    const consoleError = vi.spyOn(console, "error").mockImplementation(() => undefined);
    const importer = { import: vi.fn().mockRejectedValue(failure) };
    render(<Spreadsheet storage={false} services={{ importers: { xlsx: importer } }} />);

    fireEvent.change(screen.getByLabelText("XLSX file"), {
      target: { files: [new File(["parser input"], "parser.xlsx")] }
    });

    await waitFor(() => expect(screen.getByLabelText("Status")).toHaveTextContent(
      "XLSX import failed because the workbook parser encountered an unexpected error"
    ));
    expect(consoleError).toHaveBeenCalledWith("XLSX import failed", failure);
  });

  it("exports the active workbook as an XLSX file", async () => {
    const user = userEvent.setup();
    const createObjectURL = vi.spyOn(URL, "createObjectURL").mockReturnValue("blob:xlsx");
    vi.spyOn(URL, "revokeObjectURL").mockImplementation(() => undefined);
    const originalCreateElement = document.createElement.bind(document);
    const anchorClicks: HTMLAnchorElement[] = [];
    vi.spyOn(document, "createElement").mockImplementation((tagName) => {
      const element = originalCreateElement(tagName);
      if (tagName.toLowerCase() === "a") {
        vi.spyOn(element, "click").mockImplementation(() => undefined);
        anchorClicks.push(element as HTMLAnchorElement);
      }
      return element;
    });
    render(<App />);

    await editCell(user, "A1", "Exported");
    await openRibbonTab(user, "File");
    await user.click(screen.getByRole("button", { name: "Export XLSX" }));

    await waitFor(() => expect(createObjectURL).toHaveBeenCalledTimes(1));
    const blob = createObjectURL.mock.calls[0][0] as Blob;
    expect(blob.type).toBe("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    const imported = await importWorkbookFromXlsx(await blob.arrayBuffer());
    expect(getCellContent(imported, imported.activeSheetId, "A1")).toBe("Exported");
    expect(anchorClicks[0].download).toBe("spreadsheet.xlsx");
    expect(screen.getByLabelText("Status")).toHaveTextContent("Exported XLSX");
  });

  it("prints the active workbook from the toolbar", async () => {
    const user = userEvent.setup();
    const print = vi.spyOn(window, "print").mockImplementation(() => undefined);
    render(<App />);

    await editCell(user, "A1", "Printable");
    await openRibbonTab(user, "File");
    await user.click(screen.getByRole("button", { name: "Print workbook" }));

    expect(print).toHaveBeenCalledTimes(1);
    expect(screen.getByLabelText("Status")).toHaveTextContent("Opened print dialog");
  });

  it("keeps edits made while a CSV import is reading in undo history", async () => {
    const user = userEvent.setup();
    const readers: Array<{ onload: (() => void) | null; readAsText: ReturnType<typeof vi.fn>; result: string | null }> = [];
    function MockFileReader(this: { onload: (() => void) | null; readAsText: ReturnType<typeof vi.fn>; result: string | null }) {
      this.onload = null;
      this.readAsText = vi.fn();
      this.result = null;
      readers.push(this);
    }
    vi.stubGlobal("FileReader", MockFileReader);
    render(<App />);

    const file = new File(["Name\nImported"], "budget.csv", { type: "text/csv" });
    fireEvent.change(screen.getByLabelText("CSV file"), { target: { files: [file] } });
    await editCell(user, "A1", "interim");

    await act(async () => {
      readers[0].result = "Name\nImported";
      readers[0].onload?.();
    });

    expect(await screen.findByRole("gridcell", { name: "A1 Name" })).toHaveTextContent("Name");

    await user.click(screen.getByRole("button", { name: "Undo" }));

    expect(screen.getByRole("gridcell", { name: "A1 interim" })).toHaveTextContent("interim");
  });

  it("does not add undo history when the formula bar is blurred without a change", () => {
    render(<App />);

    const undoButton = screen.getByRole("button", { name: "Undo" });
    expect(undoButton).toBeDisabled();

    fireEvent.blur(screen.getByLabelText("Formula input"));

    expect(undoButton).toBeDisabled();
  });

  it("does not add undo history when clearing an empty selection", async () => {
    const user = userEvent.setup();
    render(<App />);

    await user.click(screen.getByRole("button", { name: "Clear selection" }));

    expect(screen.getByRole("button", { name: "Undo" })).toBeDisabled();
  });

  it("does not announce no-op edit commands as completed", async () => {
    const user = userEvent.setup();
    render(<App />);

    const status = screen.getByLabelText("Status");

    await user.click(screen.getByRole("button", { name: "Clear selection" }));
    await user.click(screen.getByRole("button", { name: "Fill down" }));
    await user.click(screen.getByRole("button", { name: "Fill right" }));

    expect(status).toHaveTextContent("Ready");
    expect(status).not.toHaveTextContent("Cleared selection");
    expect(status).not.toHaveTextContent("Filled down");
    expect(status).not.toHaveTextContent("Filled right");
  });

  it("does not announce unavailable undo or redo shortcuts", () => {
    render(<App />);

    const grid = screen.getByRole("grid", { name: "Spreadsheet grid" });
    const status = screen.getByLabelText("Status");

    fireEvent.keyDown(grid, { key: "z", ctrlKey: true });
    fireEvent.keyDown(grid, { key: "y", ctrlKey: true });

    expect(status).toHaveTextContent("Ready");
    expect(status).not.toHaveTextContent("Undone");
    expect(status).not.toHaveTextContent("Redone");
  });

  it("does not add undo history when renaming a sheet to its current name", async () => {
    const user = userEvent.setup();
    vi.spyOn(window, "prompt").mockReturnValue("Sheet1");
    render(<App />);

    await user.click(screen.getByRole("button", { name: "Rename active sheet" }));

    expect(screen.getByRole("button", { name: "Undo" })).toBeDisabled();
    expect(screen.getByLabelText("Status")).toHaveTextContent("Ready");
  });

  it("commits formula bar edits as a single undoable action", async () => {
    const user = userEvent.setup();
    render(<App />);

    await user.click(screen.getByLabelText("Formula input"));
    await user.type(screen.getByLabelText("Formula input"), "10{Enter}");
    expect(screen.getByRole("gridcell", { name: "A1 10" })).toHaveTextContent("10");

    await user.click(screen.getByRole("button", { name: "Undo" }));

    expect(screen.getByRole("gridcell", { name: "A1" })).toHaveTextContent("");
    expect(screen.getByRole("button", { name: "Undo" })).toBeDisabled();
  });

  it("rejects an invalid proposed formula from the formula bar without publishing it", async () => {
    const user = userEvent.setup();
    render(<App />);

    await applyNumberValidation(user, "A1", { max: 10 });
    const formulaInput = screen.getByLabelText("Formula input");
    await user.click(formulaInput);
    await user.type(formulaInput, "=5+5{Enter}");
    expect(screen.getByRole("gridcell", { name: "A1 10" })).toBeInTheDocument();

    await user.clear(formulaInput);
    await user.type(formulaInput, "=5+6{Enter}");

    expect(formulaInput).toHaveValue("=5+6");
    expect(screen.getByRole("gridcell", { name: "A1 10" })).toBeInTheDocument();
    expect(screen.getByLabelText("Status")).toHaveTextContent("Enter a number less than or equal to 10");
  });

  it("commits cell editor edits as a single undoable action", async () => {
    const user = userEvent.setup();
    render(<App />);

    await editCell(user, "A1", "10");
    expect(screen.getByRole("gridcell", { name: "A1 10" })).toHaveTextContent("10");

    await user.click(screen.getByRole("button", { name: "Undo" }));

    expect(screen.getByRole("gridcell", { name: "A1" })).toHaveTextContent("");
    expect(screen.getByRole("button", { name: "Undo" })).toBeDisabled();
  });

  it("fills formulas from the toolbar with adjusted relative references", async () => {
    const user = userEvent.setup();
    render(<App />);

    await editCell(user, "A1", "10");
    await editCell(user, "A2", "20");
    await editCell(user, "A3", "30");
    await editCell(user, "B1", "=A1");

    selectRange("B1 10", "B3");
    await user.click(screen.getByRole("button", { name: "Fill down" }));

    expect(screen.getByRole("gridcell", { name: "B2 20" })).toHaveTextContent("20");
    expect(screen.getByRole("gridcell", { name: "B3 30" })).toHaveTextContent("30");

    await user.click(screen.getByRole("gridcell", { name: "B2 20" }));
    expect(screen.getByLabelText("Formula input")).toHaveValue("=A2");
  });

  it.each(["Fill Down", "Fill Right", "AutoFill"] as const)(
    "rejects a translated formula from %s when the final destination violates validation",
    async (operation) => {
      const user = userEvent.setup();
      render(<App />);
      const destination = await setupFormulaFill(user, operation, 20);
      const buildSpy = vi.spyOn(HyperFormula, "buildFromSheets");

      await runFormulaFill(user, operation);

      expect(screen.getByRole("gridcell", { name: destination.address })).toHaveTextContent("");
      expect(screen.getByLabelText("Status")).toHaveTextContent("Enter a number less than or equal to 10");
      expect(buildSpy).toHaveBeenCalledTimes(1);
    }
  );

  it.each(["Fill Down", "Fill Right", "AutoFill"] as const)(
    "commits a translated formula from %s when the complete candidate is valid",
    async (operation) => {
      const user = userEvent.setup();
      render(<App />);
      const destination = await setupFormulaFill(user, operation, 8);
      const buildSpy = vi.spyOn(HyperFormula, "buildFromSheets");

      await runFormulaFill(user, operation);

      const filledCell = screen.getByRole("gridcell", { name: `${destination.address} 8` });
      expect(filledCell).toBeInTheDocument();
      await user.click(filledCell);
      expect(screen.getByLabelText("Formula input")).toHaveValue(destination.formula);
      expect(buildSpy).toHaveBeenCalledTimes(1);
    }
  );

  it("resets the scroll position of the grid view", async () => {
    const user = userEvent.setup();
    render(<App />);

    const grid = screen.getByRole("grid", { name: "Spreadsheet grid" });
    grid.scrollLeft = 320;
    grid.scrollTop = 240;

    await openRibbonTab(user, "View");
    await user.click(screen.getByRole("button", { name: "Reset view" }));

    expect(grid.scrollLeft).toBe(0);
    expect(grid.scrollTop).toBe(0);
  });

  it("resizes a column and row from their headers", () => {
    render(<App />);

    fireEvent.mouseDown(screen.getByLabelText("Resize column A"), { clientX: 100 });
    fireEvent.mouseMove(window, { clientX: 140 });
    fireEvent.mouseUp(window);

    expect(screen.getByRole("gridcell", { name: "A1" })).toHaveStyle({ width: "136px" });
    expect(screen.getByLabelText("Status")).toHaveTextContent("Set column A width");

    fireEvent.mouseDown(screen.getByLabelText("Resize row 1"), { clientY: 20 });
    fireEvent.mouseMove(window, { clientY: 34 });
    fireEvent.mouseUp(window);

    expect(screen.getByRole("gridcell", { name: "A1" })).toHaveStyle({ height: "42px" });
    expect(screen.getByLabelText("Status")).toHaveTextContent("Set row 1 height");
  });

  it("auto-fits selected columns and rows from the toolbar", async () => {
    const user = userEvent.setup();
    render(<App />);

    await editCell(user, "B1", "A much longer customer segment name");
    await user.click(screen.getByRole("columnheader", { name: "Column B" }));
    await user.click(screen.getByRole("button", { name: "Auto-fit columns" }));

    expect(screen.getByRole("gridcell", { name: "B1 A much longer customer segment name" })).toHaveStyle({
      width: "272px"
    });
    expect(screen.getByLabelText("Status")).toHaveTextContent("Auto-fit 1 column");

    fireEvent.mouseDown(screen.getByLabelText("Resize row 1"), { clientY: 20 });
    fireEvent.mouseMove(window, { clientY: 54 });
    fireEvent.mouseUp(window);
    expect(screen.getByRole("gridcell", { name: "B1 A much longer customer segment name" })).toHaveStyle({
      height: "62px"
    });

    await user.click(screen.getByRole("rowheader", { name: "Row 1" }));
    await user.click(screen.getByRole("button", { name: "Auto-fit rows" }));

    expect(screen.getByRole("gridcell", { name: "B1 A much longer customer segment name" })).toHaveStyle({
      height: "28px"
    });
    expect(screen.getByLabelText("Status")).toHaveTextContent("Auto-fit 1 row");
  });

  it("toggles wrapped text formatting from the toolbar", async () => {
    const user = userEvent.setup();
    render(<App />);

    await editCell(user, "A1", "Line one Line two Line three");
    await user.click(screen.getByRole("gridcell", { name: "A1 Line one Line two Line three" }));
    await user.click(screen.getByRole("button", { name: "Wrap text" }));

    const wrappedCell = screen.getByRole("gridcell", { name: "A1 Line one Line two Line three" });
    expect(wrappedCell).toHaveClass("wrapped-cell");
    expect(wrappedCell).toHaveStyle({ whiteSpace: "normal" });
    expect(screen.getByRole("button", { name: "Wrap text" })).toHaveClass("active-toolbar-button");
    expect(screen.getByLabelText("Status")).toHaveTextContent("Wrapped text");

    await user.click(screen.getByRole("button", { name: "Wrap text" }));

    expect(screen.getByRole("gridcell", { name: "A1 Line one Line two Line three" })).not.toHaveClass("wrapped-cell");
    expect(screen.getByRole("button", { name: "Wrap text" })).not.toHaveClass("active-toolbar-button");
    expect(screen.getByLabelText("Status")).toHaveTextContent("Unwrapped text");
  });

  it("uses keyboard shortcuts for common spreadsheet commands", async () => {
    const user = userEvent.setup();
    vi.spyOn(window, "prompt").mockReturnValueOnce("example.com/report");
    render(<App />);

    await editCell(user, "A1", "Report");
    const grid = screen.getByRole("grid", { name: "Spreadsheet grid" });
    await user.click(screen.getByRole("gridcell", { name: "A1 Report" }));

    fireEvent.keyDown(grid, { key: "b", ctrlKey: true });
    fireEvent.keyDown(grid, { key: "i", ctrlKey: true });
    const formattedCell = screen.getByRole("gridcell", { name: "A1 Report" });
    expect(formattedCell).toHaveStyle({ fontWeight: "700", fontStyle: "italic" });
    expect(screen.getByRole("button", { name: "Bold" })).toHaveClass("active-toolbar-button");
    expect(screen.getByRole("button", { name: "Italic" })).toHaveClass("active-toolbar-button");

    fireEvent.keyDown(grid, { key: "k", ctrlKey: true });
    expect(screen.getByRole("link", { name: "Report" })).toHaveAttribute("href", "https://example.com/report");

    fireEvent.keyDown(grid, { key: "l", ctrlKey: true, shiftKey: true });
    expect(screen.getByRole("complementary", { name: "Filter" })).toBeInTheDocument();
  });

  it("moves DOM focus with the active cell through repeated keyboard navigation", async () => {
    const user = userEvent.setup();
    render(<App />);

    await user.click(screen.getByRole("gridcell", { name: "A1" }));
    expect(screen.getByRole("gridcell", { name: "A1" })).toHaveFocus();

    await user.keyboard("{ArrowDown}");
    expect(screen.getByRole("gridcell", { name: "A2" })).toHaveFocus();

    await user.keyboard("{ArrowDown}");
    expect(screen.getByRole("gridcell", { name: "A3" })).toHaveFocus();

    await user.keyboard("{ArrowRight}");
    expect(screen.getByRole("gridcell", { name: "B3" })).toHaveFocus();

    await user.keyboard("{ArrowDown}");

    expect(screen.getByLabelText("Formula input")).toHaveValue("");
    expect(screen.getByLabelText("Name box")).toHaveValue("B4");
    expect(screen.getByRole("gridcell", { name: "B4" })).toHaveFocus();
  });

  it("extends the selection with Shift+Arrow keys", async () => {
    const user = userEvent.setup();
    render(<App />);

    await user.click(screen.getByRole("gridcell", { name: "A1" }));
    await user.keyboard("{Shift>}{ArrowDown}{ArrowDown}{ArrowRight}{/Shift}");

    expect(screen.getByLabelText("Name box")).toHaveValue("A1:B3");
    expect(screen.getByLabelText("Status")).toHaveTextContent("6 selected");

    await user.keyboard("{ArrowDown}");

    expect(screen.getByLabelText("Name box")).toHaveValue("A2");
  });

  it("jumps to data-region edges with Ctrl+Arrow keys", async () => {
    const user = userEvent.setup();
    render(<App />);

    await editCell(user, "A1", "10");
    await editCell(user, "A2", "20");
    await editCell(user, "A3", "30");
    await user.click(screen.getByRole("gridcell", { name: "A1 10" }));

    const grid = screen.getByRole("grid", { name: "Spreadsheet grid" });
    fireEvent.keyDown(grid, { key: "ArrowDown", ctrlKey: true });
    expect(screen.getByLabelText("Name box")).toHaveValue("A3");

    fireEvent.keyDown(grid, { key: "ArrowDown", ctrlKey: true });
    expect(screen.getByLabelText("Name box")).toHaveValue("A100");

    fireEvent.keyDown(grid, { key: "ArrowUp", ctrlKey: true });
    expect(screen.getByLabelText("Name box")).toHaveValue("A3");

    fireEvent.keyDown(grid, { key: "ArrowUp", ctrlKey: true, shiftKey: true });
    expect(screen.getByLabelText("Name box")).toHaveValue("A1:A3");
    expect(screen.getByLabelText("Status")).toHaveTextContent("3 selected");
  });

  it("moves the active cell with Tab and Shift+Tab", async () => {
    const user = userEvent.setup();
    render(<App />);

    await user.click(screen.getByRole("gridcell", { name: "A1" }));
    const grid = screen.getByRole("grid", { name: "Spreadsheet grid" });

    fireEvent.keyDown(grid, { key: "Tab" });
    expect(screen.getByLabelText("Name box")).toHaveValue("B1");

    fireEvent.keyDown(grid, { key: "Tab" });
    expect(screen.getByLabelText("Name box")).toHaveValue("C1");

    fireEvent.keyDown(grid, { key: "Tab", shiftKey: true });
    expect(screen.getByLabelText("Name box")).toHaveValue("B1");
  });

  it("moves down after committing an edit with Enter and right with Tab", async () => {
    const user = userEvent.setup();
    render(<App />);

    await editCell(user, "A1", "10");
    expect(screen.getByLabelText("Name box")).toHaveValue("A2");

    await user.dblClick(screen.getByRole("gridcell", { name: "B2" }));
    await user.type(screen.getByLabelText("Cell editor B2"), "hello");
    await user.keyboard("{Tab}");

    expect(screen.getByRole("gridcell", { name: "B2 hello" })).toBeInTheDocument();
    expect(screen.getByLabelText("Name box")).toHaveValue("C2");

    await user.dblClick(screen.getByRole("gridcell", { name: "C2" }));
    await user.type(screen.getByLabelText("Cell editor C2"), "up");
    await user.keyboard("{Shift>}{Enter}{/Shift}");

    expect(screen.getByRole("gridcell", { name: "C2 up" })).toBeInTheDocument();
    expect(screen.getByLabelText("Name box")).toHaveValue("C1");
  });

  it("returns focus to the spreadsheet grid after commit and cancel", async () => {
    const user = userEvent.setup();
    render(<App />);

    await user.dblClick(screen.getByRole("gridcell", { name: "A1" }));
    await user.type(screen.getByLabelText("Cell editor A1"), "saved{Enter}");
    const grid = screen.getByRole("grid", { name: "Spreadsheet grid" });
    expect(grid).toContainElement(document.activeElement as HTMLElement);

    await user.dblClick(screen.getByRole("gridcell", { name: "B1" }));
    await user.type(screen.getByLabelText("Cell editor B1"), "discarded");
    await user.keyboard("{Escape}");

    expect(screen.queryByLabelText("Cell editor B1")).not.toBeInTheDocument();
    expect(grid).toContainElement(document.activeElement as HTMLElement);
  });

  it("selects the data region, then the whole sheet, with Ctrl+A", async () => {
    const user = userEvent.setup();
    render(<App />);

    await editCell(user, "A1", "Name");
    await editCell(user, "B1", "Amount");
    await editCell(user, "A2", "West");
    await editCell(user, "B2", "5");
    await user.click(screen.getByRole("gridcell", { name: "A1 Name" }));

    const grid = screen.getByRole("grid", { name: "Spreadsheet grid" });
    fireEvent.keyDown(grid, { key: "a", ctrlKey: true });
    expect(screen.getByLabelText("Name box")).toHaveValue("A1:B2");
    expect(screen.getByLabelText("Status")).toHaveTextContent("4 selected");

    fireEvent.keyDown(grid, { key: "a", ctrlKey: true });
    expect(screen.getByLabelText("Name box")).toHaveValue("A1:Z100");
    expect(screen.getByLabelText("Status")).toHaveTextContent("2600 selected");
  });

  it("extends the selection with shift-click", async () => {
    const user = userEvent.setup();
    render(<App />);

    await user.click(screen.getByRole("gridcell", { name: "A1" }));
    fireEvent.click(screen.getByRole("gridcell", { name: "C3" }), { shiftKey: true });

    expect(screen.getByLabelText("Name box")).toHaveValue("A1:C3");
    expect(screen.getByLabelText("Status")).toHaveTextContent("9 selected");
  });

  it("shows numeric summaries for the selected range", async () => {
    const user = userEvent.setup();
    render(<App />);

    await editCell(user, "A1", "10");
    await editCell(user, "A2", "20");
    await editCell(user, "A3", "30");

    selectRange("A1 10", "A3 30");

    const status = screen.getByLabelText("Status");
    expect(status).toHaveTextContent("Count 3");
    expect(status).toHaveTextContent("Sum 60");
    expect(status).toHaveTextContent("Avg 20");
    expect(status).toHaveTextContent("Min 10");
    expect(status).toHaveTextContent("Max 30");
  });

  it("zooms the worksheet from status bar controls", async () => {
    const user = userEvent.setup();
    render(<App />);

    const grid = screen.getByRole("grid", { name: "Spreadsheet grid" });
    expect(grid).toHaveAttribute("data-zoom-level", "100");
    expect(screen.getByRole("button", { name: "Reset zoom" })).toHaveTextContent("100%");

    await user.click(screen.getByRole("button", { name: "Zoom in" }));

    expect(grid).toHaveAttribute("data-zoom-level", "125");
    expect(screen.getByRole("button", { name: "Reset zoom" })).toHaveTextContent("125%");
    expect(screen.getByLabelText("Status")).toHaveTextContent("Zoom 125%");

    await user.click(screen.getByRole("button", { name: "Zoom out" }));
    await user.click(screen.getByRole("button", { name: "Zoom out" }));

    expect(grid).toHaveAttribute("data-zoom-level", "75");

    await user.click(screen.getByRole("button", { name: "Reset zoom" }));

    expect(grid).toHaveAttribute("data-zoom-level", "100");
    expect(screen.getByLabelText("Status")).toHaveTextContent("Zoom reset");
  });

  it("inserts an AutoSum formula for the selected range", async () => {
    const user = userEvent.setup();
    render(<App />);

    await user.click(screen.getByRole("gridcell", { name: "A1" }));
    fireEvent.paste(screen.getByRole("grid", { name: "Spreadsheet grid" }), {
      clipboardData: {
        getData: () => "10\n20\n30"
      }
    });

    selectRange("A1 10", "A3 30");
    await user.click(screen.getByRole("button", { name: "AutoSum" }));

    expect(screen.getByRole("gridcell", { name: "A4 60" })).toHaveTextContent("60");
    expect(screen.getByLabelText("Formula input")).toHaveValue("=SUM(A1:A3)");
    expect(screen.getByLabelText("Status")).toHaveTextContent("Inserted AutoSum");
  });

  it("inserts an AutoAverage formula from the toolbar picker", async () => {
    const user = userEvent.setup();
    render(<App />);

    await user.click(screen.getByRole("gridcell", { name: "A1" }));
    fireEvent.paste(screen.getByRole("grid", { name: "Spreadsheet grid" }), {
      clipboardData: {
        getData: () => "10\n20\n30"
      }
    });

    selectRange("A1 10", "A3 30");
    await openRibbonTab(user, "Formulas");
    await user.click(screen.getByRole("button", { name: "AutoSum options" }));
    await user.click(screen.getByRole("menuitem", { name: "Average" }));

    expect(screen.getByRole("gridcell", { name: "A4 20" })).toHaveTextContent("20");
    expect(screen.getByLabelText("Formula input")).toHaveValue("=AVERAGE(A1:A3)");
    expect(screen.getByLabelText("Status")).toHaveTextContent("Inserted Average for A1:A3");
  });

  it("infers an AutoSum range above a blank active cell", async () => {
    const user = userEvent.setup();
    render(<App />);

    await editCell(user, "A1", "10");
    await editCell(user, "A2", "20");
    await editCell(user, "A3", "30");

    await user.click(screen.getByRole("gridcell", { name: "A4" }));
    await user.click(screen.getByRole("button", { name: "AutoSum" }));

    expect(screen.getByRole("gridcell", { name: "A4 60" })).toHaveTextContent("60");
    expect(screen.getByLabelText("Formula input")).toHaveValue("=SUM(A1:A3)");
  });

  it("sorts the selected range by the first selected column", async () => {
    const user = userEvent.setup();
    render(<App />);

    await user.click(screen.getByRole("gridcell", { name: "A1" }));
    fireEvent.paste(screen.getByRole("grid", { name: "Spreadsheet grid" }), {
      clipboardData: {
        getData: () => "Delta\t4\nAlpha\t1\nCharlie\t3"
      }
    });
    selectRange("A1 Delta", "B3 3");

    await user.click(screen.getByRole("button", { name: "Sort A to Z" }));

    expect(screen.getByRole("gridcell", { name: "A1 Alpha" })).toHaveTextContent("Alpha");
    expect(screen.getByRole("gridcell", { name: "B1 1" })).toHaveTextContent("1");
    expect(screen.getByRole("gridcell", { name: "A3 Delta" })).toHaveTextContent("Delta");
    expect(screen.getByRole("gridcell", { name: "B3 4" })).toHaveTextContent("4");

    await user.click(screen.getByRole("button", { name: "Sort Z to A" }));

    expect(screen.getByRole("gridcell", { name: "A1 Delta" })).toHaveTextContent("Delta");
    expect(screen.getByRole("gridcell", { name: "B1 4" })).toHaveTextContent("4");
  });

  it("sorts a structured table from the ribbon through the row-id-aware command", async () => {
    const user = userEvent.setup();
    const onCommandResult = vi.fn();
    render(
      <Spreadsheet
        defaultWorkbook={structuredTableWorkbook()}
        storage={false}
        onCommandResult={onCommandResult}
      />
    );

    await user.click(screen.getByRole("gridcell", { name: "A2 West" }));
    await user.click(screen.getByRole("button", { name: "Sort A to Z" }));

    expect(screen.getByRole("gridcell", { name: "A2 East" })).toHaveTextContent("East");
    expect(screen.getByRole("gridcell", { name: "B2 8" })).toHaveTextContent("8");
    expect(screen.getByRole("gridcell", { name: "A3 West" })).toHaveTextContent("West");
    expect(onCommandResult).toHaveBeenCalledWith(expect.objectContaining({
      command: {
        type: "table.sort",
        tableId: "table-sales",
        sorting: [{ columnId: "sales-region", direction: "asc" }]
      },
      result: expect.objectContaining({ status: "committed" })
    }));
  });

  it("sorts a structured AutoFilter column through the row-id-aware command", async () => {
    const user = userEvent.setup();
    const onCommandResult = vi.fn();
    const workbook = structuredTableWorkbook();
    workbook.sheets[0] = {
      ...workbook.sheets[0],
      autoFilterRange: workbook.tables[0].range
    };
    render(<Spreadsheet defaultWorkbook={workbook} storage={false} onCommandResult={onCommandResult} />);

    await user.click(screen.getByRole("button", { name: "Open AutoFilter menu for Sales" }));
    await user.click(within(screen.getByRole("menu", { name: "AutoFilter menu for Sales" }))
      .getByRole("menuitem", { name: "Sort A to Z" }));

    expect(screen.getByRole("gridcell", { name: "B2 8" })).toHaveTextContent("8");
    expect(screen.getByRole("gridcell", { name: "B3 10" })).toHaveTextContent("10");
    expect(onCommandResult).toHaveBeenCalledWith(expect.objectContaining({
      command: {
        type: "table.sort",
        tableId: "table-sales",
        sorting: [{ columnId: "sales-value", direction: "asc" }]
      },
      result: expect.objectContaining({ status: "committed" })
    }));
  });

  it("explains why a partially overlapping structured-table selection cannot be sorted", async () => {
    const user = userEvent.setup();
    const onCommandResult = vi.fn();
    render(
      <Spreadsheet
        defaultWorkbook={structuredTableWorkbook()}
        storage={false}
        onCommandResult={onCommandResult}
      />
    );

    selectRange("A2 West", "C3");
    await user.click(screen.getByRole("button", { name: "Sort A to Z" }));

    expect(screen.getByLabelText("Status")).toHaveTextContent(
      "Sort is not supported for selections that partially overlap a structured table. Select only cells within the table or convert it to a range first."
    );
    expect(onCommandResult).not.toHaveBeenCalledWith(expect.objectContaining({
      command: expect.objectContaining({ type: "range.sort" })
    }));
  });

  it("explains why a partially overlapping AutoFilter range cannot be sorted", async () => {
    const user = userEvent.setup();
    const onCommandResult = vi.fn();
    const workbook = structuredTableWorkbook();
    workbook.sheets[0] = {
      ...workbook.sheets[0],
      autoFilterRange: {
        start: { ...workbook.tables[0].range.start },
        end: { row: workbook.tables[0].range.end.row, column: 2 }
      }
    };
    render(<Spreadsheet defaultWorkbook={workbook} storage={false} onCommandResult={onCommandResult} />);

    await user.click(screen.getByRole("button", { name: "Open AutoFilter menu for Region" }));
    await user.click(within(screen.getByRole("menu", { name: "AutoFilter menu for Region" }))
      .getByRole("menuitem", { name: "Sort A to Z" }));

    expect(screen.getByLabelText("Status")).toHaveTextContent(
      "Sort is not supported for selections that partially overlap a structured table. Select only cells within the table or convert it to a range first."
    );
    expect(onCommandResult).not.toHaveBeenCalledWith(expect.objectContaining({
      command: expect.objectContaining({ type: "range.sort" })
    }));
  });

  it("explains how to remove duplicates when the selection is inside a structured table", async () => {
    const user = userEvent.setup();
    const onCommandResult = vi.fn();
    render(
      <Spreadsheet
        defaultWorkbook={structuredTableWorkbook()}
        storage={false}
        onCommandResult={onCommandResult}
      />
    );

    await user.click(screen.getByRole("gridcell", { name: "A2 West" }));
    await user.click(screen.getByRole("button", { name: "Remove duplicates" }));

    expect(screen.getByLabelText("Status")).toHaveTextContent(
      "Remove duplicates is not supported for structured tables. Convert the table to a range first."
    );
    expect(onCommandResult).not.toHaveBeenCalledWith(expect.objectContaining({
      command: expect.objectContaining({ type: "range.removeDuplicates" })
    }));
  });

  it("sorts formula rows by evaluated values and keeps moved formulas relative", async () => {
    const user = userEvent.setup();
    render(<App />);

    await user.click(screen.getByRole("gridcell", { name: "A1" }));
    fireEvent.paste(screen.getByRole("grid", { name: "Spreadsheet grid" }), {
      clipboardData: {
        getData: () => "=B1\t20\n=B2\t10"
      }
    });
    selectRange("A1 20", "B2 10");

    await user.click(screen.getByRole("button", { name: "Sort A to Z" }));

    expect(screen.getByRole("gridcell", { name: "A1 10" })).toHaveTextContent("10");
    expect(screen.getByRole("gridcell", { name: "B1 10" })).toHaveTextContent("10");
    expect(screen.getByRole("gridcell", { name: "A2 20" })).toHaveTextContent("20");
    expect(screen.getByRole("gridcell", { name: "B2 20" })).toHaveTextContent("20");
    await user.click(screen.getByRole("gridcell", { name: "A1 10" }));
    expect(screen.getByLabelText("Formula input")).toHaveValue("=B1");
  });

  it("filters table data from an AutoFilter header menu", async () => {
    const user = userEvent.setup();
    render(<App />);

    await user.click(screen.getByRole("gridcell", { name: "A1" }));
    fireEvent.paste(screen.getByRole("grid", { name: "Spreadsheet grid" }), {
      clipboardData: {
        getData: () => "Region\tSales\nWest\t10\nEast\t8\nNorth\t11\nWest\t12"
      }
    });
    selectRange("A1 Region", "B5 12");

    await user.click(screen.getByRole("button", { name: "Filter" }));
    await user.selectOptions(screen.getByLabelText("Filter operator"), "equals");
    await user.type(screen.getByLabelText("Filter value"), "West");
    await user.click(screen.getByRole("button", { name: "Apply filter" }));

    await user.click(screen.getByRole("button", { name: "Open AutoFilter menu for Region" }));
    await user.click(within(screen.getByRole("menu", { name: "AutoFilter menu for Region" })).getByRole("menuitem", { name: "Clear filter from Region" }));
    expect(screen.getByRole("gridcell", { name: "A3 East" })).toHaveTextContent("East");

    await user.click(screen.getByRole("button", { name: "Open AutoFilter menu for Region" }));
    const menu = screen.getByRole("menu", { name: "AutoFilter menu for Region" });
    await user.click(within(menu).getByRole("menuitemcheckbox", { name: "East" }));
    await user.click(within(menu).getByRole("menuitemcheckbox", { name: "West" }));
    await user.click(within(menu).getByRole("menuitem", { name: "Apply selected values" }));

    expect(screen.getByRole("gridcell", { name: "A1 Region" })).toHaveTextContent("Region");
    expect(screen.getByRole("gridcell", { name: "A2 West" })).toHaveTextContent("West");
    expect(screen.getByRole("gridcell", { name: "A3 East" })).toHaveTextContent("East");
    expect(screen.queryByRole("gridcell", { name: "A4 North" })).not.toBeInTheDocument();
    expect(screen.getByRole("gridcell", { name: "A5 West" })).toHaveTextContent("West");
    expect(screen.queryByRole("menu", { name: "AutoFilter menu for Region" })).not.toBeInTheDocument();
  });

  it("filters zero without treating blanks or formula-empty results as zero", async () => {
    const user = userEvent.setup();
    render(<App />);

    await editCell(user, "A1", "Value");
    await editCell(user, "A2", "0");
    await editCell(user, "A4", '=""');
    selectRange("A1 Value", "A4");

    await user.click(screen.getByRole("button", { name: "Filter" }));
    await user.selectOptions(screen.getByLabelText("Filter operator"), "equals");
    await user.type(screen.getByLabelText("Filter value"), "0");
    await user.click(screen.getByRole("button", { name: "Apply filter" }));

    expect(screen.getByRole("gridcell", { name: "A1 Value" })).toBeInTheDocument();
    expect(screen.getByRole("gridcell", { name: "A2 0" })).toBeInTheDocument();
    expect(screen.queryByRole("gridcell", { name: "A3" })).not.toBeInTheDocument();
    expect(screen.queryByRole("gridcell", { name: "A4" })).not.toBeInTheDocument();
  });

  it("inserts and deletes selected rows and columns from the toolbar", async () => {
    const user = userEvent.setup();
    render(<App />);

    await user.click(screen.getByRole("gridcell", { name: "A1" }));
    fireEvent.paste(screen.getByRole("grid", { name: "Spreadsheet grid" }), {
      clipboardData: {
        getData: () => "Name\tAmount\nWest\t10"
      }
    });

    await user.click(screen.getByRole("gridcell", { name: "A2 West" }));
    await user.click(screen.getByRole("button", { name: "Insert row above" }));
    expect(screen.getByRole("gridcell", { name: "A2" })).toHaveTextContent("");
    expect(screen.getByRole("gridcell", { name: "A3 West" })).toHaveTextContent("West");

    await user.click(screen.getByRole("button", { name: "Delete row" }));
    expect(screen.getByRole("gridcell", { name: "A2 West" })).toHaveTextContent("West");

    await user.click(screen.getByRole("gridcell", { name: "B1 Amount" }));
    await user.click(screen.getByRole("button", { name: "Insert column left" }));
    expect(screen.getByRole("gridcell", { name: "B1" })).toHaveTextContent("");
    expect(screen.getByRole("gridcell", { name: "C1 Amount" })).toHaveTextContent("Amount");

    await user.click(screen.getByRole("button", { name: "Delete column" }));
    expect(screen.getByRole("gridcell", { name: "B1 Amount" })).toHaveTextContent("Amount");
  });

  it("inserts the selected width to the right from the ribbon split menu", async () => {
    const user = userEvent.setup();
    render(<Spreadsheet storage={false} />);

    selectRange("A1", "B1");
    await user.click(screen.getByRole("button", { name: "Insert columns options" }));
    await user.click(screen.getByRole("menuitem", { name: "Insert column right" }));

    expect(screen.getByLabelText("Status")).toHaveTextContent("Inserted 2 columns right");
    expect(screen.getByLabelText("Name box")).toHaveValue("C1:D1");
  });

  it("expands the selected structured table at its right boundary", async () => {
    const user = userEvent.setup();
    render(<Spreadsheet defaultWorkbook={structuredTableWorkbook()} storage={false} />);

    await user.click(screen.getByRole("gridcell", { name: "B2 10" }));
    await user.click(screen.getByRole("button", { name: "Insert columns options" }));
    await user.click(screen.getByRole("menuitem", { name: "Insert column right" }));

    expect(screen.getByRole("gridcell", { name: "C1 Column3" })).toBeVisible();
    expect(screen.getByLabelText("Status")).toHaveTextContent("Inserted 1 column right");
  });

  it("does not expand same-sheet tables outside the selected row or column boundary context", async () => {
    const user = userEvent.setup();
    render(<Spreadsheet defaultWorkbook={columnBoundaryContextWorkbook()} storage={false} />);

    await user.click(screen.getByRole("gridcell", { name: "B2 10" }));
    await user.click(screen.getByRole("button", { name: "Insert columns options" }));
    await user.click(screen.getByRole("menuitem", { name: "Insert column right" }));

    expect(screen.getByRole("gridcell", { name: "C1 Column3" })).toBeVisible();
    expect(screen.getByRole("gridcell", { name: "C5" })).toHaveTextContent("");
    expect(screen.getByRole("gridcell", { name: "D1 Department" })).toHaveTextContent("Department");
    expect(screen.getByRole("gridcell", { name: "E1 Cost" })).toHaveTextContent("Cost");
    expect(screen.getByLabelText("Status")).toHaveTextContent("Inserted 1 column right");
  });

  it("does not include a matching table boundary from another sheet", async () => {
    const user = userEvent.setup();
    render(<Spreadsheet defaultWorkbook={columnBoundaryContextWorkbook()} storage={false} />);

    await user.click(screen.getByRole("gridcell", { name: "B2 10" }));
    await user.click(screen.getByRole("button", { name: "Insert columns options" }));
    await user.click(screen.getByRole("menuitem", { name: "Insert column right" }));

    expect(screen.getByRole("gridcell", { name: "C1 Column3" })).toBeVisible();
    await user.click(screen.getByRole("tab", { name: "Other" }));
    expect(screen.getByRole("grid", { name: "Spreadsheet grid" })).toHaveAttribute("aria-colcount", "27");
    expect(screen.getByRole("gridcell", { name: "A1 OtherRegion" })).toHaveTextContent("OtherRegion");
    expect(screen.getByRole("gridcell", { name: "B1 OtherSales" })).toHaveTextContent("OtherSales");
    expect(screen.getByRole("gridcell", { name: "C1" })).toHaveTextContent("");
  });

  it("appends a column after the final worksheet column", async () => {
    const user = userEvent.setup();
    render(<Spreadsheet storage={false} />);

    const nameBox = screen.getByLabelText("Name box");
    await user.clear(nameBox);
    await user.type(nameBox, "Z1{Enter}");
    await user.click(screen.getByRole("button", { name: "Insert columns options" }));
    await user.click(screen.getByRole("menuitem", { name: "Insert column right" }));

    expect(screen.getByLabelText("Status")).toHaveTextContent("Inserted 1 column right");
    expect(screen.getByRole("grid", { name: "Spreadsheet grid" })).toHaveAttribute("aria-colcount", "28");
    expect(nameBox).toHaveValue("AA1");
    expect(screen.getByRole("gridcell", { name: "AA1" })).toHaveFocus();

    expect(screen.getByRole("button", { name: "Undo" })).toBeEnabled();
    await user.click(screen.getByRole("button", { name: "Undo" }));
    await waitFor(() => expect(screen.getByRole("grid", { name: "Spreadsheet grid" })).toHaveAttribute("aria-colcount", "27"));
    expect(nameBox).toHaveValue("Z1");
    expect(screen.getByRole("gridcell", { name: "Z1" })).toHaveFocus();

    await user.click(screen.getByRole("button", { name: "Redo" }));
    await waitFor(() => expect(screen.getByRole("grid", { name: "Spreadsheet grid" })).toHaveAttribute("aria-colcount", "28"));
    expect(nameBox).toHaveValue("AA1");
    expect(screen.getByRole("gridcell", { name: "AA1" })).toHaveFocus();
  });

  it("rejects right column insertion on a protected sheet", async () => {
    const user = userEvent.setup();
    const workbook = createBlankWorkbook();
    render(
      <Spreadsheet
        defaultWorkbook={{
          ...workbook,
          sheets: workbook.sheets.map((sheet) => ({
            ...sheet,
            protection: { ...sheet.protection, isProtected: true }
          }))
        }}
        storage={false}
      />
    );

    const grid = screen.getByRole("grid", { name: "Spreadsheet grid" });
    expect(grid).toHaveAttribute("aria-colcount", "27");
    await user.click(screen.getByRole("button", { name: "Insert columns options" }));
    await user.click(screen.getByRole("menuitem", { name: "Insert column right" }));

    expect(screen.getByLabelText("Status")).toHaveTextContent("Sheet is protected");
    expect(screen.getByLabelText("Name box")).toHaveValue("A1");
    expect(grid).toHaveAttribute("aria-colcount", "27");
    expect(screen.getByRole("button", { name: "Undo" })).toBeDisabled();
  });

  it("undoes a multi-column right insertion as one action", async () => {
    const user = userEvent.setup();
    render(<Spreadsheet storage={false} />);

    const grid = screen.getByRole("grid", { name: "Spreadsheet grid" });
    selectRange("A1", "B1");
    await user.click(screen.getByRole("button", { name: "Insert columns options" }));
    await user.click(screen.getByRole("menuitem", { name: "Insert column right" }));

    expect(grid).toHaveAttribute("aria-colcount", "29");
    expect(screen.getByRole("button", { name: "Undo" })).toBeEnabled();
    await user.click(screen.getByRole("button", { name: "Undo" }));

    expect(grid).toHaveAttribute("aria-colcount", "27");
    expect(screen.getByRole("button", { name: "Undo" })).toBeDisabled();
  });

  it("inserts to the right from a column-header context menu", async () => {
    const user = userEvent.setup();
    render(<Spreadsheet storage={false} />);

    fireEvent.contextMenu(screen.getByRole("columnheader", { name: "Column B" }), {
      clientX: 240,
      clientY: 60
    });
    const menu = screen.getByRole("menu", { name: "Column B context menu" });
    await user.click(within(menu).getByRole("menuitem", { name: "Insert column right" }));

    expect(screen.getByLabelText("Status")).toHaveTextContent("Inserted 1 column right");
    expect(screen.getByLabelText("Name box")).toHaveValue("C1:C100");
  });

  it("finds and replaces values from a compact panel", async () => {
    const user = userEvent.setup();
    render(<App />);

    await user.click(screen.getByRole("gridcell", { name: "A1" }));
    fireEvent.paste(screen.getByRole("grid", { name: "Spreadsheet grid" }), {
      clipboardData: {
        getData: () => "Region\tSales\nWest\t10\nWest\t12\nEast\t8"
      }
    });

    await user.click(screen.getByRole("button", { name: "Find and replace" }));
    await user.type(screen.getByLabelText("Find text"), "West");
    await user.click(screen.getByRole("button", { name: "Find next" }));

    expect(screen.getByLabelText("Name box")).toHaveValue("A2");

    await user.type(screen.getByLabelText("Replace text"), "North");
    await user.click(screen.getByRole("button", { name: "Replace all" }));

    expect(screen.getByRole("gridcell", { name: "A2 North" })).toHaveTextContent("North");
    expect(screen.getByRole("gridcell", { name: "A3 North" })).toHaveTextContent("North");
    expect(screen.getByLabelText("Status")).toHaveTextContent("Replaced 2 cells");
  });

  it("shows formula suggestions when a formula starts", async () => {
    const user = userEvent.setup();
    render(<App />);

    await user.click(screen.getByLabelText("Formula input"));
    await user.type(screen.getByLabelText("Formula input"), "=s");

    expect(screen.getByRole("listbox", { name: "Formula suggestions" })).toBeInTheDocument();
    expect(screen.getByRole("option", { name: /SUM/ })).toBeInTheDocument();

    await user.click(screen.getByRole("option", { name: /SUM/ }));
    expect(screen.getByLabelText("Formula input")).toHaveValue("=SUM(");
  });

  it("shows an Excel-like formula helper as soon as equals is typed", async () => {
    const user = userEvent.setup();
    render(<App />);

    await user.click(screen.getByLabelText("Formula input"));
    await user.type(screen.getByLabelText("Formula input"), "=");

    const suggestions = screen.getByRole("listbox", { name: "Formula suggestions" });
    expect(suggestions).toHaveClass("formula-suggestions--excel");
    expect(suggestions).toHaveAttribute("aria-orientation", "vertical");
    expect(within(suggestions).getAllByRole("option")).toHaveLength(6);
    expect(within(suggestions).getByRole("option", { name: "SUM" })).toHaveClass("formula-suggestion-option--active");
    expect(within(suggestions).getByText("SUM(number1, [number2], ...)")).toBeInTheDocument();
    expect(within(suggestions).getByText("Adds numbers or ranges.")).toBeInTheDocument();
  });

  it("accepts formula suggestions from the keyboard in the formula bar", async () => {
    const user = userEvent.setup();
    render(<App />);

    await user.click(screen.getByLabelText("Formula input"));
    await user.type(screen.getByLabelText("Formula input"), "=av{Tab}");

    expect(screen.getByLabelText("Formula input")).toHaveValue("=AVERAGE(");
  });

  it("opens a searchable function library and inserts formulas", async () => {
    const user = userEvent.setup();
    render(<App />);

    await openRibbonTab(user, "Formulas");
    await user.click(screen.getByRole("button", { name: "Function library" }));

    const panel = screen.getByRole("complementary", { name: "Function library" });
    expect(within(panel).getByText(/418 functions/)).toBeInTheDocument();

    await user.type(within(panel).getByLabelText("Search functions"), "lookup");
    expect(within(panel).getByRole("button", { name: "Insert XLOOKUP" })).toBeInTheDocument();

    await user.click(within(panel).getByRole("button", { name: "Insert XLOOKUP" }));

    expect(screen.getByLabelText("Formula input")).toHaveValue("=XLOOKUP(");
    expect(screen.getByLabelText("Status")).toHaveTextContent("Inserted XLOOKUP function");
  });

  it("commits the highlighted formula suggestion from the formula bar", async () => {
    const user = userEvent.setup();
    render(<App />);

    const formulaInput = screen.getByLabelText("Formula input");
    await user.click(formulaInput);
    await user.type(formulaInput, "=av");
    fireEvent.keyDown(formulaInput, { key: "ArrowDown" });

    expect(screen.getByRole("option", { name: "AVEDEV" })).toHaveAttribute("aria-selected", "true");

    fireEvent.keyDown(formulaInput, { key: "Tab" });

    expect(formulaInput).toHaveValue("=AVEDEV(");
  });

  it("tracks hovered formula bar suggestions as the active option", async () => {
    const user = userEvent.setup();
    render(<App />);

    const formulaInput = screen.getByLabelText("Formula input");
    await user.click(formulaInput);
    await user.type(formulaInput, "=av");

    const suggestions = screen.getByRole("listbox", { name: "Formula suggestions" });
    expect(formulaInput).toHaveAttribute("aria-controls", "formula-bar-suggestions");
    expect(formulaInput).toHaveAttribute("aria-expanded", "true");
    expect(formulaInput).toHaveAttribute("aria-activedescendant", "formula-bar-suggestions-option-average");
    expect(suggestions).toHaveAttribute("id", "formula-bar-suggestions");

    fireEvent.mouseMove(screen.getByRole("option", { name: "AVEDEV" }));

    expect(screen.getByRole("option", { name: "AVEDEV" })).toHaveAttribute("aria-selected", "true");
    expect(formulaInput).toHaveAttribute("aria-activedescendant", "formula-bar-suggestions-option-avedev");

    fireEvent.keyDown(formulaInput, { key: "Tab" });

    expect(formulaInput).toHaveValue("=AVEDEV(");
  });

  it("dismisses formula bar suggestions with Escape without clearing the formula", async () => {
    const user = userEvent.setup();
    render(<App />);

    const formulaInput = screen.getByLabelText("Formula input");
    await user.click(formulaInput);
    await user.type(formulaInput, "=av");

    expect(screen.getByRole("listbox", { name: "Formula suggestions" })).toBeInTheDocument();

    fireEvent.keyDown(formulaInput, { key: "Escape" });

    expect(formulaInput).toHaveValue("=av");
    expect(screen.queryByRole("listbox", { name: "Formula suggestions" })).not.toBeInTheDocument();

    await user.type(formulaInput, "e");

    expect(screen.getByRole("listbox", { name: "Formula suggestions" })).toBeInTheDocument();
  });

  it("shows formula suggestions while editing directly in a cell", async () => {
    const user = userEvent.setup();
    render(<App />);

    await user.dblClick(screen.getByRole("gridcell", { name: "A1" }));
    const editor = screen.getByLabelText("Cell editor A1");
    await user.type(editor, "=s");

    expect(screen.getByRole("listbox", { name: "Formula suggestions" })).toBeInTheDocument();
    expect(screen.getByRole("option", { name: /SUM/ })).toBeInTheDocument();

    await user.click(screen.getByRole("option", { name: /SUM/ }));
    expect(screen.getByLabelText("Cell editor A1")).toHaveValue("=SUM(");
  });

  it("commits non-function formulas with Enter without autocomplete hijacking", async () => {
    const user = userEvent.setup();
    render(<App />);

    await user.dblClick(screen.getByRole("gridcell", { name: "A1" }));
    const editor = screen.getByLabelText("Cell editor A1");
    await user.type(editor, "=1+2");

    // No function-name prefix is being typed, so no suggestions may appear.
    expect(screen.queryByRole("listbox", { name: "Formula suggestions" })).not.toBeInTheDocument();

    fireEvent.keyDown(editor, { key: "Enter" });

    expect(await screen.findByRole("gridcell", { name: "A1 3" })).toHaveTextContent("3");
  });

  it("keeps Enter committing the raw text even while suggestions are visible", async () => {
    const user = userEvent.setup();
    render(<App />);

    await user.dblClick(screen.getByRole("gridcell", { name: "A1" }));
    const editor = screen.getByLabelText("Cell editor A1");
    await user.type(editor, "=av");

    expect(screen.getByRole("listbox", { name: "Formula suggestions" })).toBeInTheDocument();

    fireEvent.keyDown(editor, { key: "Enter" });

    // Enter commits what was typed (an unknown name evaluates to #NAME?) instead of
    // silently replacing the input with the highlighted suggestion.
    expect(await screen.findByRole("gridcell", { name: "A1 #NAME?" })).toBeInTheDocument();
  });

  it("accepts formula suggestions from the keyboard in a cell editor", async () => {
    const user = userEvent.setup();
    render(<App />);

    await user.dblClick(screen.getByRole("gridcell", { name: "A1" }));
    await user.type(screen.getByLabelText("Cell editor A1"), "=av{Tab}");

    expect(screen.getByLabelText("Cell editor A1")).toHaveValue("=AVERAGE(");
  });

  it("commits the highlighted formula suggestion from a cell editor", async () => {
    const user = userEvent.setup();
    render(<App />);

    await user.dblClick(screen.getByRole("gridcell", { name: "A1" }));
    const editor = screen.getByLabelText("Cell editor A1");
    await user.type(editor, "=av");
    fireEvent.keyDown(editor, { key: "ArrowDown" });

    expect(screen.getByRole("option", { name: "AVEDEV" })).toHaveAttribute("aria-selected", "true");

    fireEvent.keyDown(editor, { key: "Tab" });

    expect(editor).toHaveValue("=AVEDEV(");
  });

  it("tracks hovered cell editor suggestions as the active option", async () => {
    const user = userEvent.setup();
    render(<App />);

    await user.dblClick(screen.getByRole("gridcell", { name: "A1" }));
    const editor = screen.getByLabelText("Cell editor A1");
    await user.type(editor, "=av");

    const suggestions = screen.getByRole("listbox", { name: "Formula suggestions" });
    expect(editor).toHaveAttribute("aria-controls", "cell-editor-a1-formula-suggestions");
    expect(editor).toHaveAttribute("aria-expanded", "true");
    expect(editor).toHaveAttribute("aria-activedescendant", "cell-editor-a1-formula-suggestions-option-average");
    expect(suggestions).toHaveAttribute("id", "cell-editor-a1-formula-suggestions");

    fireEvent.mouseMove(screen.getByRole("option", { name: "AVEDEV" }));

    expect(screen.getByRole("option", { name: "AVEDEV" })).toHaveAttribute("aria-selected", "true");
    expect(editor).toHaveAttribute("aria-activedescendant", "cell-editor-a1-formula-suggestions-option-avedev");

    fireEvent.keyDown(editor, { key: "Tab" });

    expect(editor).toHaveValue("=AVEDEV(");
  });

  it("dismisses cell editor suggestions with Escape before canceling the edit", async () => {
    const user = userEvent.setup();
    render(<App />);

    await user.dblClick(screen.getByRole("gridcell", { name: "A1" }));
    const editor = screen.getByLabelText("Cell editor A1");
    await user.type(editor, "=av");

    expect(screen.getByRole("listbox", { name: "Formula suggestions" })).toBeInTheDocument();

    fireEvent.keyDown(editor, { key: "Escape" });

    expect(editor).toHaveValue("=av");
    expect(screen.queryByRole("listbox", { name: "Formula suggestions" })).not.toBeInTheDocument();
    expect(screen.getByLabelText("Cell editor A1")).toBeInTheDocument();

    await user.type(editor, "e");

    expect(screen.getByRole("listbox", { name: "Formula suggestions" })).toBeInTheDocument();

    fireEvent.keyDown(screen.getByLabelText("Cell editor A1"), { key: "Escape" });

    expect(screen.queryByRole("listbox", { name: "Formula suggestions" })).not.toBeInTheDocument();
    expect(screen.getByLabelText("Cell editor A1")).toBeInTheDocument();

    fireEvent.keyDown(screen.getByLabelText("Cell editor A1"), { key: "Escape" });

    expect(screen.queryByLabelText("Cell editor A1")).not.toBeInTheDocument();
  });

  it("commits an autocomplete formula selected from a cell editor", async () => {
    const user = userEvent.setup();
    render(<App />);

    await editCell(user, "A2", "10");
    await editCell(user, "A3", "15");

    await user.dblClick(screen.getByRole("gridcell", { name: "A1" }));
    const editor = screen.getByLabelText("Cell editor A1");
    await user.type(editor, "=s");
    await user.click(screen.getByRole("option", { name: /SUM/ }));
    await user.type(screen.getByLabelText("Cell editor A1"), "A2:A3){Enter}");

    expect(screen.getByRole("gridcell", { name: "A1 25" })).toHaveTextContent("25");
  });

  it("applies text and color formatting to selected cells", async () => {
    const user = userEvent.setup();
    render(<App />);

    await editCell(user, "A1", "Styled");
    await user.click(screen.getByRole("gridcell", { name: "A1 Styled" }));
    await user.click(screen.getByRole("button", { name: "Bold" }));
    await user.click(screen.getByRole("button", { name: "Italic" }));
    fireEvent.input(screen.getByLabelText("Text color"), { target: { value: "#ffffff" } });
    fireEvent.input(screen.getByLabelText("Fill color"), { target: { value: "#1f6feb" } });

    const cell = screen.getByRole("gridcell", { name: "A1 Styled" });
    expect(cell).toHaveStyle({ fontWeight: "700", fontStyle: "italic" });
    expect(cell).toHaveStyle({ color: "#ffffff", backgroundColor: "#1f6feb" });
  });

  it("copies active cell formatting with Format Painter", async () => {
    const user = userEvent.setup();
    render(<App />);

    await editCell(user, "A1", "Styled");
    await editCell(user, "B1", "Target");
    await user.click(screen.getByRole("gridcell", { name: "A1 Styled" }));
    await user.click(screen.getByRole("button", { name: "Bold" }));
    await user.click(screen.getByRole("button", { name: "Italic" }));
    fireEvent.input(screen.getByLabelText("Text color"), { target: { value: "#ffffff" } });
    fireEvent.input(screen.getByLabelText("Fill color"), { target: { value: "#1f6feb" } });

    await user.click(screen.getByRole("button", { name: "Format painter" }));
    expect(screen.getByRole("button", { name: "Format painter" })).toHaveClass("active-toolbar-button");

    await user.click(screen.getByRole("gridcell", { name: "B1 Target" }));

    const targetCell = screen.getByRole("gridcell", { name: "B1 Target" });
    expect(targetCell).toHaveTextContent("Target");
    expect(targetCell).toHaveStyle({
      fontWeight: "700",
      fontStyle: "italic",
      color: "#ffffff",
      backgroundColor: "#1f6feb"
    });
    expect(screen.getByRole("button", { name: "Format painter" })).not.toHaveClass("active-toolbar-button");
    expect(screen.getByLabelText("Status")).toHaveTextContent("Painted format to B1");
  });

  it("keeps Format Painter armed when the target is read-only", async () => {
    const user = userEvent.setup();
    render(<App />);

    await editCell(user, "A1", "Styled");
    await editCell(user, "B1", "Locked");
    await user.click(screen.getByRole("gridcell", { name: "A1 Styled" }));
    await user.click(screen.getByRole("button", { name: "Bold" }));
    await user.click(screen.getByRole("button", { name: "Format painter" }));
    await openRibbonTab(user, "Review");
    await user.click(screen.getByRole("button", { name: "Protect sheet" }));
    await openRibbonTab(user, "Home");

    await user.click(screen.getByRole("gridcell", { name: "B1 Locked" }));

    expect(screen.getByRole("gridcell", { name: "B1 Locked" })).not.toHaveStyle({ fontWeight: "700" });
    expect(screen.getByRole("button", { name: "Format painter" })).toHaveClass("active-toolbar-button");
    expect(screen.getByLabelText("Status")).toHaveTextContent("B1 is read-only");
  });

  it("opens a cell context menu and applies wrapped text", async () => {
    const user = userEvent.setup();
    render(<App />);

    await editCell(user, "A1", "Context menu note");
    fireEvent.contextMenu(screen.getByRole("gridcell", { name: "A1 Context menu note" }), {
      clientX: 120,
      clientY: 160
    });

    const menu = screen.getByRole("menu", { name: "Cell context menu" });
    expect(menu).toHaveTextContent("A1");

    await user.click(within(menu).getByRole("menuitem", { name: "Wrap text" }));

    expect(screen.queryByRole("menu", { name: "Cell context menu" })).not.toBeInTheDocument();
    expect(screen.getByRole("gridcell", { name: "A1 Context menu note" })).toHaveClass("wrapped-cell");
    expect(screen.getByLabelText("Status")).toHaveTextContent("Wrapped text");
  });

  it("copies and pastes formats from the cell context menu", async () => {
    const user = userEvent.setup();
    render(<App />);

    await editCell(user, "A1", "Styled");
    await user.click(screen.getByRole("gridcell", { name: "A1 Styled" }));
    await user.click(screen.getByRole("button", { name: "Bold" }));
    await editCell(user, "B1", "Plain");

    fireEvent.contextMenu(screen.getByRole("gridcell", { name: "A1 Styled" }), {
      clientX: 120,
      clientY: 160
    });
    expect(within(screen.getByRole("menu", { name: "Cell context menu" })).getByRole("menuitem", { name: "Cut" })).toBeInTheDocument();
    await user.click(within(screen.getByRole("menu", { name: "Cell context menu" })).getByRole("menuitem", { name: "Copy" }));

    fireEvent.contextMenu(screen.getByRole("gridcell", { name: "B1 Plain" }), {
      clientX: 220,
      clientY: 160
    });
    await user.click(within(screen.getByRole("menu", { name: "Cell context menu" })).getByRole("menuitem", { name: "Paste formats" }));

    expect(screen.getByRole("gridcell", { name: "B1 Plain" })).toHaveTextContent("Plain");
    expect(screen.getByRole("gridcell", { name: "B1 Plain" })).toHaveStyle({ fontWeight: "700" });
    expect(screen.getByLabelText("Status")).toHaveTextContent("Pasted formats");
  });

  it("clears formats from the cell context menu without clearing content", async () => {
    const user = userEvent.setup();
    render(<App />);

    await editCell(user, "A1", "Styled");
    await user.click(screen.getByRole("gridcell", { name: "A1 Styled" }));
    await user.click(screen.getByRole("button", { name: "Bold" }));
    const styledCell = screen.getByRole("gridcell", { name: "A1 Styled" });
    expect(styledCell).toHaveStyle({ fontWeight: "700" });

    fireEvent.contextMenu(styledCell, {
      clientX: 120,
      clientY: 160
    });
    await user.click(within(screen.getByRole("menu", { name: "Cell context menu" })).getByRole("menuitem", { name: "Clear formats" }));

    const clearedCell = screen.getByRole("gridcell", { name: "A1 Styled" });
    expect(clearedCell).toHaveTextContent("Styled");
    expect(clearedCell).not.toHaveStyle({ fontWeight: "700" });
    expect(screen.getByLabelText("Status")).toHaveTextContent("Cleared formats");
  });

  it("clears conditional formats from the cell context menu without clearing direct formats", async () => {
    const user = userEvent.setup();
    render(<App />);

    await editCell(user, "A1", "12");
    await user.click(screen.getByRole("gridcell", { name: "A1 12" }));
    await user.click(screen.getByRole("button", { name: "Italic" }));
    await user.click(screen.getByRole("button", { name: "Conditional formatting" }));
    await user.selectOptions(screen.getByLabelText("Conditional rule"), "greaterThan");
    await user.clear(screen.getByLabelText("Conditional value"));
    await user.type(screen.getByLabelText("Conditional value"), "10");
    await user.click(screen.getByRole("button", { name: "Apply conditional format" }));

    const highlightedCell = screen.getByRole("gridcell", { name: "A1 12" });
    expect(highlightedCell).toHaveClass("conditional-format-cell");
    expect(highlightedCell).toHaveStyle({ fontStyle: "italic" });

    fireEvent.contextMenu(highlightedCell, {
      clientX: 120,
      clientY: 160
    });
    await user.click(
      within(screen.getByRole("menu", { name: "Cell context menu" })).getByRole("menuitem", {
        name: "Clear conditional formats"
      })
    );

    const clearedCell = screen.getByRole("gridcell", { name: "A1 12" });
    expect(clearedCell).toHaveTextContent("12");
    expect(clearedCell).not.toHaveClass("conditional-format-cell");
    expect(clearedCell).toHaveStyle({ fontStyle: "italic" });
    expect(screen.getByLabelText("Status")).toHaveTextContent("Cleared conditional formatting");
  });

  it("clears hyperlinks from the cell context menu without clearing content", async () => {
    const user = userEvent.setup();
    vi.spyOn(window, "prompt").mockReturnValueOnce("example.com/report");
    render(<App />);

    await editCell(user, "A1", "Report");
    await user.click(screen.getByRole("gridcell", { name: "A1 Report" }));
    await openRibbonTab(user, "Review");
    await user.click(screen.getByRole("button", { name: "Link" }));
    expect(screen.getByRole("link", { name: "Report" })).toHaveAttribute("href", "https://example.com/report");

    fireEvent.contextMenu(screen.getByRole("gridcell", { name: "A1 Report" }), {
      clientX: 120,
      clientY: 160
    });
    await user.click(within(screen.getByRole("menu", { name: "Cell context menu" })).getByRole("menuitem", { name: "Clear hyperlinks" }));

    expect(screen.queryByRole("link", { name: "Report" })).not.toBeInTheDocument();
    expect(screen.getByRole("gridcell", { name: "A1 Report" })).toHaveTextContent("Report");
    expect(screen.getByLabelText("Status")).toHaveTextContent("Cleared hyperlinks");
  });

  it("clears data validation from the cell context menu without clearing content", async () => {
    const user = userEvent.setup();
    render(<App />);

    await editCell(user, "A1", "20");
    await user.click(screen.getByRole("gridcell", { name: "A1 20" }));
    await user.click(screen.getByRole("button", { name: "Data validation" }));
    await user.selectOptions(screen.getByLabelText("Validation type"), "number");
    await user.clear(screen.getByLabelText("Minimum"));
    await user.type(screen.getByLabelText("Minimum"), "1");
    await user.clear(screen.getByLabelText("Maximum"));
    await user.type(screen.getByLabelText("Maximum"), "10");
    await user.click(screen.getByRole("button", { name: "Apply validation" }));

    const invalidCell = screen.getByRole("gridcell", { name: "A1 20" });
    expect(invalidCell).toHaveClass("invalid-validation-cell");

    fireEvent.contextMenu(invalidCell, {
      clientX: 120,
      clientY: 160
    });
    await user.click(within(screen.getByRole("menu", { name: "Cell context menu" })).getByRole("menuitem", { name: "Clear validation" }));

    const clearedCell = screen.getByRole("gridcell", { name: "A1 20" });
    expect(clearedCell).toHaveTextContent("20");
    expect(clearedCell).not.toHaveClass("invalid-validation-cell");
    expect(screen.getByLabelText("Status")).toHaveTextContent("Cleared data validation");
  });

  it("chooses list validation values from a selected cell dropdown", async () => {
    const user = userEvent.setup();
    render(<App />);

    await user.click(screen.getByRole("gridcell", { name: "A1" }));
    await user.click(screen.getByRole("button", { name: "Data validation" }));
    await user.clear(screen.getByLabelText("List values"));
    await user.type(screen.getByLabelText("List values"), "Open, Closed, Blocked");
    await user.click(screen.getByRole("button", { name: "Apply validation" }));

    await user.click(screen.getByRole("button", { name: "Open validation choices for A1" }));
    await user.click(within(screen.getByRole("listbox", { name: "Validation choices for A1" })).getByRole("option", { name: "Closed" }));

    expect(screen.getByRole("gridcell", { name: "A1 Closed" })).toHaveTextContent("Closed");
    expect(screen.queryByRole("listbox", { name: "Validation choices for A1" })).not.toBeInTheDocument();
  });

  it("does not show validation dropdowns for read-only list cells", async () => {
    const user = userEvent.setup();
    render(<App />);

    await user.click(screen.getByRole("gridcell", { name: "A1" }));
    await user.click(screen.getByRole("button", { name: "Data validation" }));
    await user.clear(screen.getByLabelText("List values"));
    await user.type(screen.getByLabelText("List values"), "Open, Closed");
    await user.click(screen.getByRole("button", { name: "Apply validation" }));
    await openRibbonTab(user, "Review");
    await user.click(screen.getByRole("button", { name: "Protect sheet" }));

    expect(screen.queryByRole("button", { name: "Open validation choices for A1" })).not.toBeInTheDocument();
  });

  it("clears all content, formats, and hyperlinks from the cell context menu", async () => {
    const user = userEvent.setup();
    vi.spyOn(window, "prompt").mockReturnValueOnce("example.com/report");
    render(<App />);

    await editCell(user, "A1", "Report");
    await user.click(screen.getByRole("gridcell", { name: "A1 Report" }));
    await user.click(screen.getByRole("button", { name: "Bold" }));
    await openRibbonTab(user, "Review");
    await user.click(screen.getByRole("button", { name: "Link" }));
    const linkedCell = screen.getByRole("gridcell", { name: "A1 Report" });
    expect(screen.getByRole("link", { name: "Report" })).toHaveAttribute("href", "https://example.com/report");
    expect(linkedCell).toHaveStyle({ fontWeight: "700" });

    fireEvent.contextMenu(linkedCell, {
      clientX: 120,
      clientY: 160
    });
    await user.click(within(screen.getByRole("menu", { name: "Cell context menu" })).getByRole("menuitem", { name: "Clear all" }));

    const clearedCell = screen.getByRole("gridcell", { name: "A1" });
    expect(clearedCell).toHaveTextContent("");
    expect(clearedCell).not.toHaveStyle({ fontWeight: "700" });
    expect(screen.queryByRole("link", { name: "Report" })).not.toBeInTheDocument();
    expect(screen.getByLabelText("Status")).toHaveTextContent("Cleared all");
  });

  it("clears comments from the cell context menu without clearing content", async () => {
    const user = userEvent.setup();
    vi.spyOn(window, "prompt").mockReturnValueOnce("Review the forecast");
    render(<App />);

    await editCell(user, "A1", "Forecast");
    await user.click(screen.getByRole("gridcell", { name: "A1 Forecast" }));
    await openRibbonTab(user, "Review");
    await user.click(screen.getByRole("button", { name: "Comment" }));
    const commentedCell = screen.getByRole("gridcell", { name: "A1 Forecast" });
    expect(commentedCell).toHaveClass("commented-cell");
    expect(commentedCell).toHaveAttribute("title", "Comment: Review the forecast");

    fireEvent.contextMenu(commentedCell, {
      clientX: 120,
      clientY: 160
    });
    await user.click(within(screen.getByRole("menu", { name: "Cell context menu" })).getByRole("menuitem", { name: "Clear comments" }));

    expect(screen.getByRole("gridcell", { name: "A1 Forecast" })).toHaveTextContent("Forecast");
    expect(screen.getByRole("gridcell", { name: "A1 Forecast" })).not.toHaveClass("commented-cell");
    expect(screen.getByRole("gridcell", { name: "A1 Forecast" })).not.toHaveAttribute("title");
    expect(screen.getByLabelText("Status")).toHaveTextContent("Cleared comments");
  });

  it("inserts rows and columns from the cell context menu", async () => {
    const user = userEvent.setup();
    render(<App />);

    await editCell(user, "A1", "Name");
    await editCell(user, "B1", "Amount");
    await editCell(user, "A2", "West");

    fireEvent.contextMenu(screen.getByRole("gridcell", { name: "A2 West" }), {
      clientX: 120,
      clientY: 188
    });
    await user.click(within(screen.getByRole("menu", { name: "Cell context menu" })).getByRole("menuitem", { name: "Insert row above" }));

    expect(screen.getByRole("gridcell", { name: "A2" })).toBeInTheDocument();
    expect(screen.getByRole("gridcell", { name: "A3 West" })).toHaveTextContent("West");
    expect(screen.getByLabelText("Status")).toHaveTextContent("Inserted 1 row");

    fireEvent.contextMenu(screen.getByRole("gridcell", { name: "B1 Amount" }), {
      clientX: 220,
      clientY: 160
    });
    await user.click(within(screen.getByRole("menu", { name: "Cell context menu" })).getByRole("menuitem", { name: "Insert column left" }));

    expect(screen.getByRole("gridcell", { name: "B1" })).toBeInTheDocument();
    expect(screen.getByRole("gridcell", { name: "C1 Amount" })).toHaveTextContent("Amount");
    expect(screen.getByLabelText("Status")).toHaveTextContent("Inserted 1 column left");
  });

  it("inserts a column to the right from the cell context menu", async () => {
    const user = userEvent.setup();
    render(<Spreadsheet storage={false} />);

    await editCell(user, "B1", "Amount");
    await editCell(user, "C1", "Tail");
    fireEvent.contextMenu(screen.getByRole("gridcell", { name: "B1 Amount" }), {
      clientX: 220,
      clientY: 160
    });
    await user.click(
      within(screen.getByRole("menu", { name: "Cell context menu" })).getByRole("menuitem", {
        name: "Insert column right"
      })
    );

    expect(screen.getByRole("gridcell", { name: "C1" })).toHaveTextContent("");
    expect(screen.getByRole("gridcell", { name: "D1 Tail" })).toHaveTextContent("Tail");
    expect(screen.getByLabelText("Status")).toHaveTextContent("Inserted 1 column right");
  });

  it("applies and clears cell borders from the toolbar", async () => {
    const user = userEvent.setup();
    render(<App />);

    await editCell(user, "A1", "Bordered");
    await user.click(screen.getByRole("gridcell", { name: "A1 Bordered" }));
    await user.click(screen.getByRole("button", { name: "All borders" }));

    const borderedCell = screen.getByRole("gridcell", { name: "A1 Bordered" });
    expect(borderedCell).toHaveStyle({ borderTopStyle: "solid", borderRightStyle: "solid" });
    expect(screen.getByLabelText("Status")).toHaveTextContent("Applied all borders to A1");

    await user.click(screen.getByRole("button", { name: "All borders options" }));
    await user.click(screen.getByRole("menuitem", { name: "Clear borders" }));

    expect(screen.getByRole("gridcell", { name: "A1 Bordered" })).not.toHaveStyle({ borderTopStyle: "solid" });
    expect(screen.getByLabelText("Status")).toHaveTextContent("Cleared borders from A1");
  });

  it("applies font family and size from the toolbar and renders them in the grid", async () => {
    const user = userEvent.setup();
    render(<App />);

    await editCell(user, "A1", "Styled");
    await user.click(screen.getByRole("gridcell", { name: "A1 Styled" }));
    await user.selectOptions(screen.getByLabelText("Font family"), "Georgia");
    await user.selectOptions(screen.getByLabelText("Font size"), "18");

    const styledCell = screen.getByRole("gridcell", { name: "A1 Styled" });
    expect(styledCell).toHaveStyle({ fontFamily: "Georgia" });
    expect(styledCell).toHaveStyle({ fontSize: "18pt" });
    expect(screen.getByLabelText("Font family")).toHaveValue("Georgia");
    expect(screen.getByLabelText("Font size")).toHaveValue("18");
  });

  it("shows mixed formatting state for a selection that disagrees", async () => {
    const user = userEvent.setup();
    render(<App />);

    await editCell(user, "A1", "bolded");
    await user.click(screen.getByRole("gridcell", { name: "A1 bolded" }));
    await user.click(screen.getByRole("button", { name: "Bold" }));
    await editCell(user, "A2", "plain");

    selectRange("A1 bolded", "A2 plain");
    const boldButton = screen.getByRole("button", { name: "Bold" });
    expect(boldButton).toHaveAttribute("aria-pressed", "mixed");

    // Toggling from mixed applies bold to the whole selection, like Excel.
    await user.click(boldButton);
    expect(screen.getByRole("button", { name: "Bold" })).toHaveAttribute("aria-pressed", "true");
    expect(screen.getByRole("gridcell", { name: "A2 plain" })).toHaveStyle({ fontWeight: "700" });
  });

  it("treats an explicit General number format the same as unformatted cells", async () => {
    const user = userEvent.setup();
    render(<App />);

    await editCell(user, "A1", "1");
    await user.click(screen.getByRole("gridcell", { name: "A1 1" }));
    await user.selectOptions(screen.getByLabelText("Number format"), "currency");
    await user.selectOptions(screen.getByLabelText("Number format"), "general");
    await editCell(user, "A2", "2");

    // A1 now stores numberFormat "general" explicitly while A2 has no format
    // entry at all; the selection is effectively uniform, not mixed.
    selectRange("A1 1", "A2 2");
    expect(screen.getByLabelText("Number format")).toHaveValue("general");
  });

  it("infers a date format only when number format is absent and preserves explicit General", async () => {
    const user = userEvent.setup();
    render(<App />);

    await editCell(user, "B1", "2026-07-08");
    expect(screen.getByRole("gridcell", { name: "B1 Jul 8, 2026" })).toHaveTextContent("Jul 8, 2026");

    await user.click(screen.getByRole("gridcell", { name: addressNamePattern("A1") }));
    await user.selectOptions(screen.getByLabelText("Number format"), "currency");
    await user.selectOptions(screen.getByLabelText("Number format"), "general");
    await editCell(user, "A1", "2026-07-08");

    expect(screen.getByRole("gridcell", { name: "A1 46211" })).toHaveTextContent("46211");
    expect(screen.getByLabelText("Number format")).toHaveValue("general");
  });

  it("marks panel-launching toolbar buttons as expanded while their panel is open", async () => {
    const user = userEvent.setup();
    render(<App />);

    const filterButton = screen.getByRole("button", { name: "Filter" });
    expect(filterButton).toHaveAttribute("aria-expanded", "false");

    await user.click(filterButton);
    expect(screen.getByRole("button", { name: "Filter" })).toHaveAttribute("aria-expanded", "true");
    expect(screen.getByRole("button", { name: "Filter" })).toHaveClass("active-toolbar-button");

    await user.click(screen.getByRole("button", { name: "Filter" }));
    expect(screen.getByRole("button", { name: "Filter" })).toHaveAttribute("aria-expanded", "false");
  });

  it("applies quick number formats from the toolbar buttons", async () => {
    const user = userEvent.setup();
    render(<App />);

    await editCell(user, "A1", "1234.5");
    await user.click(screen.getByRole("gridcell", { name: "A1 1234.5" }));
    await user.click(screen.getByRole("button", { name: "Currency format" }));

    expect(screen.getByRole("gridcell", { name: "A1 $1,234.50" })).toBeInTheDocument();
    expect(screen.getByRole("button", { name: "Currency format" })).toHaveAttribute("aria-pressed", "true");

    await user.click(screen.getByRole("button", { name: "Percent format" }));
    expect(screen.getByRole("gridcell", { name: "A1 123,450%" })).toBeInTheDocument();
  });

  it("applies number formats and alignment to selected cells", async () => {
    const user = userEvent.setup();
    render(<App />);

    await editCell(user, "A1", "1234.5");
    await user.click(screen.getByRole("gridcell", { name: "A1 1234.5" }));
    await user.selectOptions(screen.getByLabelText("Number format"), "currency");
    await user.click(screen.getByRole("button", { name: "Align right" }));
    await user.click(screen.getByRole("button", { name: "Align bottom" }));

    const currencyCell = screen.getByRole("gridcell", { name: "A1 $1,234.50" });
    expect(currencyCell).toHaveTextContent("$1,234.50");
    expect(currencyCell).toHaveStyle({ textAlign: "right" });
    expect(currencyCell).toHaveStyle({ alignItems: "flex-end" });

    await editCell(user, "B1", "0.25");
    await user.click(screen.getByRole("gridcell", { name: "B1 0.25" }));
    await user.selectOptions(screen.getByLabelText("Number format"), "percent");

    expect(screen.getByRole("gridcell", { name: "B1 25%" })).toHaveTextContent("25%");
  });

  it("applies list and number data validation to selected cells", async () => {
    const user = userEvent.setup();
    render(<App />);

    await user.click(screen.getByRole("gridcell", { name: "A1" }));
    await user.click(screen.getByRole("button", { name: "Data validation" }));
    await user.selectOptions(screen.getByLabelText("Validation type"), "list");
    await user.clear(screen.getByLabelText("List values"));
    await user.type(screen.getByLabelText("List values"), "Open, Closed");
    await user.click(screen.getByRole("button", { name: "Apply validation" }));

    await user.dblClick(screen.getByRole("gridcell", { name: "A1" }));
    await user.selectOptions(screen.getByLabelText("Cell editor A1"), "Closed");
    fireEvent.blur(screen.getByLabelText("Cell editor A1"));
    expect(screen.getByRole("gridcell", { name: "A1 Closed" })).toHaveTextContent("Closed");

    await user.click(screen.getByRole("gridcell", { name: "B1" }));
    await user.click(screen.getByRole("button", { name: "Data validation" }));
    await user.selectOptions(screen.getByLabelText("Validation type"), "number");
    await user.clear(screen.getByLabelText("Minimum"));
    await user.type(screen.getByLabelText("Minimum"), "1");
    await user.clear(screen.getByLabelText("Maximum"));
    await user.type(screen.getByLabelText("Maximum"), "10");
    await user.click(screen.getByRole("button", { name: "Apply validation" }));

    await user.dblClick(screen.getByRole("gridcell", { name: "B1" }));
    await user.type(screen.getByLabelText("Cell editor B1"), "20{Enter}");

    expect(screen.getByLabelText("Cell editor B1")).toHaveValue("20");
    expect(screen.getByLabelText("Status")).toHaveTextContent("Enter a number between 1 and 10");
  });

  it("does not build candidate engines while a validated formula draft is being typed", async () => {
    const user = userEvent.setup();
    render(<App />);

    await applyNumberValidation(user, "A1", { max: 10 });
    await editCell(user, "A1", "=5+5");
    await user.dblClick(screen.getByRole("gridcell", { name: "A1 10" }));
    const editor = screen.getByLabelText("Cell editor A1");
    const buildSpy = vi.spyOn(HyperFormula, "buildFromSheets");
    const destroySpy = vi.spyOn(HyperFormula.prototype, "destroy");

    await user.clear(editor);
    await user.type(editor, "=4+6");

    expect(buildSpy).not.toHaveBeenCalled();

    await user.keyboard("{Enter}");
    expect(buildSpy).toHaveBeenCalledTimes(1);
    expect(destroySpy).toHaveBeenCalledTimes(1);
    expect(screen.getByRole("gridcell", { name: "A1 10" })).toBeInTheDocument();
  });

  it("does not build a candidate engine for a formula without a validation rule", async () => {
    const user = userEvent.setup();
    render(<App />);
    const buildSpy = vi.spyOn(HyperFormula, "buildFromSheets");

    await editCell(user, "A1", "=1+1");

    expect(screen.getByRole("gridcell", { name: "A1 2" })).toBeInTheDocument();
    expect(buildSpy).not.toHaveBeenCalled();
  });

  it("validates the proposed formula result without publishing a rejected candidate", async () => {
    const user = userEvent.setup();
    render(<App />);

    await user.click(screen.getByRole("gridcell", { name: "A1" }));
    await user.click(screen.getByRole("button", { name: "Data validation" }));
    await user.selectOptions(screen.getByLabelText("Validation type"), "number");
    await user.clear(screen.getByLabelText("Minimum"));
    await user.type(screen.getByLabelText("Minimum"), "1");
    await user.clear(screen.getByLabelText("Maximum"));
    await user.type(screen.getByLabelText("Maximum"), "10");
    await user.click(screen.getByRole("button", { name: "Apply validation" }));

    await editCell(user, "A1", "=5+5");
    expect(screen.getByRole("gridcell", { name: "A1 10" })).not.toHaveClass("invalid-validation-cell");
    await editCell(user, "B1", "=A1*2");
    expect(screen.getByRole("gridcell", { name: "B1 20" })).toBeInTheDocument();

    await user.dblClick(screen.getByRole("gridcell", { name: "A1 10" }));
    const editor = screen.getByLabelText("Cell editor A1");
    await user.clear(editor);
    await user.type(editor, "=5+6{Enter}");

    expect(editor).toHaveValue("=5+6");
    expect(screen.getByRole("gridcell", { name: /^A1(?:\s|$)/ })).toBeInTheDocument();
    expect(screen.getByRole("gridcell", { name: "B1 20" })).toBeInTheDocument();
    expect(screen.getByLabelText("Status")).toHaveTextContent("Enter a number between 1 and 10");

    await user.keyboard("{Escape}");
    expect(screen.getByRole("gridcell", { name: "A1 10" })).toBeInTheDocument();
    expect(screen.getByRole("gridcell", { name: "B1 20" })).toBeInTheDocument();
  });

  it("applies text-length data validation to selected cells", async () => {
    const user = userEvent.setup();
    render(<App />);

    await user.click(screen.getByRole("gridcell", { name: "A1" }));
    await user.click(screen.getByRole("button", { name: "Data validation" }));
    await user.selectOptions(screen.getByLabelText("Validation type"), "textLength");
    await user.clear(screen.getByLabelText("Minimum length"));
    await user.type(screen.getByLabelText("Minimum length"), "2");
    await user.clear(screen.getByLabelText("Maximum length"));
    await user.type(screen.getByLabelText("Maximum length"), "5");
    await user.click(screen.getByRole("button", { name: "Apply validation" }));

    await user.dblClick(screen.getByRole("gridcell", { name: "A1" }));
    await user.type(screen.getByLabelText("Cell editor A1"), "Western{Enter}");

    expect(screen.getByLabelText("Cell editor A1")).toHaveValue("Western");
    expect(screen.getByLabelText("Status")).toHaveTextContent("Enter text between 2 and 5 characters");

    await user.clear(screen.getByLabelText("Cell editor A1"));
    await user.type(screen.getByLabelText("Cell editor A1"), "West{Enter}");

    expect(screen.getByRole("gridcell", { name: "A1 West" })).not.toHaveClass("invalid-validation-cell");

    await user.click(screen.getByRole("button", { name: "Data validation" }));
    const rules = screen.getByLabelText("Data validation rules");
    expect(within(rules).getByText("A1")).toBeInTheDocument();
    expect(within(rules).getByText("Text length: 2 to 5")).toBeInTheDocument();
  });

  it("lists and deletes data validation rules from the panel", async () => {
    const user = userEvent.setup();
    render(<App />);

    selectRange("A1", "A2");
    await user.click(screen.getByRole("button", { name: "Data validation" }));
    await user.clear(screen.getByLabelText("List values"));
    await user.type(screen.getByLabelText("List values"), "Open, Closed");
    await user.click(screen.getByRole("button", { name: "Apply validation" }));

    await user.click(screen.getByRole("gridcell", { name: "A1" }));
    await user.click(screen.getByRole("button", { name: "Open validation choices for A1" }));
    expect(screen.getByRole("option", { name: "Closed" })).toBeInTheDocument();
    await user.click(screen.getByRole("option", { name: "Closed" }));

    await user.click(screen.getByRole("button", { name: "Data validation" }));
    expect(screen.getByRole("heading", { name: "Existing rules" })).toBeInTheDocument();
    expect(screen.getByText("A1:A2")).toBeInTheDocument();
    expect(screen.getByText("List: Open, Closed")).toBeInTheDocument();

    await user.click(screen.getByRole("button", { name: "Delete data validation rule A1:A2 List: Open, Closed" }));

    expect(screen.queryByRole("button", { name: "Open validation choices for A1" })).not.toBeInTheDocument();
    expect(screen.getByText("No data validation rules")).toBeInTheDocument();
    expect(screen.getByLabelText("Status")).toHaveTextContent("Deleted data validation rule");
  });

  it("prevents invalid number validation bounds", async () => {
    const user = userEvent.setup();
    render(<App />);

    await user.click(screen.getByRole("gridcell", { name: "A1" }));
    await user.click(screen.getByRole("button", { name: "Data validation" }));
    await user.selectOptions(screen.getByLabelText("Validation type"), "number");

    const applyButton = screen.getByRole("button", { name: "Apply validation" });
    await user.type(screen.getByLabelText("Minimum"), "abc");
    expect(applyButton).toBeDisabled();

    await user.clear(screen.getByLabelText("Minimum"));
    await user.type(screen.getByLabelText("Minimum"), "10");
    await user.type(screen.getByLabelText("Maximum"), "1");
    expect(applyButton).toBeDisabled();

    await user.clear(screen.getByLabelText("Maximum"));
    await user.type(screen.getByLabelText("Maximum"), "20");
    expect(applyButton).toBeEnabled();
    await user.click(applyButton);

    expect(screen.getByLabelText("Status")).toHaveTextContent("Applied data validation");
  });

  it("applies conditional formatting to matching values in a selected range", async () => {
    const user = userEvent.setup();
    render(<App />);

    await user.click(screen.getByRole("gridcell", { name: "A1" }));
    fireEvent.paste(screen.getByRole("grid", { name: "Spreadsheet grid" }), {
      clipboardData: {
        getData: () => "5\n12\n8"
      }
    });

    selectRange("A1 5", "A3 8");
    await user.click(screen.getByRole("button", { name: "Conditional formatting" }));
    await user.selectOptions(screen.getByLabelText("Conditional rule"), "greaterThan");
    await user.clear(screen.getByLabelText("Conditional value"));
    await user.type(screen.getByLabelText("Conditional value"), "10");
    await user.click(screen.getByRole("button", { name: "Apply conditional format" }));

    expect(screen.getByRole("gridcell", { name: "A2 12" })).toHaveClass("conditional-format-cell");
    expect(screen.getByRole("gridcell", { name: "A2 12" })).toHaveStyle({
      backgroundColor: "#fff1d6",
      color: "#8a4b00"
    });
    expect(screen.getByRole("gridcell", { name: "A1 5" })).not.toHaveClass("conditional-format-cell");
    expect(screen.getByLabelText("Status")).toHaveTextContent("Applied conditional formatting");
  });

  it("applies blank conditional formatting without requiring a value", async () => {
    const user = userEvent.setup();
    render(<App />);

    await user.click(screen.getByRole("gridcell", { name: "A1" }));
    fireEvent.paste(screen.getByRole("grid", { name: "Spreadsheet grid" }), {
      clipboardData: {
        getData: () => "Task\tStatus\nBuild\tDone\nTest\t"
      }
    });

    selectRange("B1 Status", "B3");
    await user.click(screen.getByRole("button", { name: "Conditional formatting" }));
    await user.selectOptions(screen.getByLabelText("Conditional rule"), "blank");

    expect(screen.queryByLabelText("Conditional value")).not.toBeInTheDocument();
    expect(screen.getByRole("button", { name: "Apply conditional format" })).toBeEnabled();

    await user.click(screen.getByRole("button", { name: "Apply conditional format" }));

    expect(screen.getByRole("gridcell", { name: "B3" })).toHaveClass("conditional-format-cell");
    expect(screen.getByRole("gridcell", { name: "B2 Done" })).not.toHaveClass("conditional-format-cell");

    await user.click(screen.getByRole("button", { name: "Conditional formatting" }));
    const rules = screen.getByLabelText("Conditional format rules");
    expect(within(rules).getByText("Blank")).toBeInTheDocument();
  });

  it("applies duplicate-value conditional formatting across the selected range", async () => {
    const user = userEvent.setup();
    render(<App />);

    await user.click(screen.getByRole("gridcell", { name: "A1" }));
    fireEvent.paste(screen.getByRole("grid", { name: "Spreadsheet grid" }), {
      clipboardData: {
        getData: () => "Region\nWest\nEast\nWest"
      }
    });

    selectRange("A1 Region", "A4 West");
    await user.click(screen.getByRole("button", { name: "Conditional formatting" }));
    await user.selectOptions(screen.getByLabelText("Conditional rule"), "duplicate");

    expect(screen.queryByLabelText("Conditional value")).not.toBeInTheDocument();
    expect(screen.getByRole("button", { name: "Apply conditional format" })).toBeEnabled();

    await user.click(screen.getByRole("button", { name: "Apply conditional format" }));

    expect(screen.getByRole("gridcell", { name: "A2 West" })).toHaveClass("conditional-format-cell");
    expect(screen.getByRole("gridcell", { name: "A4 West" })).toHaveClass("conditional-format-cell");
    expect(screen.getByRole("gridcell", { name: "A3 East" })).not.toHaveClass("conditional-format-cell");

    await user.click(screen.getByRole("button", { name: "Conditional formatting" }));
    const rules = screen.getByLabelText("Conditional format rules");
    expect(within(rules).getByText("Duplicate values")).toBeInTheDocument();
  });

  it("applies top-value conditional formatting with a rank input", async () => {
    const user = userEvent.setup();
    render(<App />);

    await user.click(screen.getByRole("gridcell", { name: "A1" }));
    fireEvent.paste(screen.getByRole("grid", { name: "Spreadsheet grid" }), {
      clipboardData: {
        getData: () => "Score\n10\n30\n20\n30"
      }
    });

    selectRange("A1 Score", "A5 30");
    await user.click(screen.getByRole("button", { name: "Conditional formatting" }));
    await user.selectOptions(screen.getByLabelText("Conditional rule"), "top");
    await user.clear(screen.getByLabelText("Conditional rank"));
    await user.type(screen.getByLabelText("Conditional rank"), "2");

    expect(screen.queryByLabelText("Conditional value")).not.toBeInTheDocument();
    expect(screen.getByRole("button", { name: "Apply conditional format" })).toBeEnabled();

    await user.click(screen.getByRole("button", { name: "Apply conditional format" }));

    expect(screen.getByRole("gridcell", { name: "A3 30" })).toHaveClass("conditional-format-cell");
    expect(screen.getByRole("gridcell", { name: "A5 30" })).toHaveClass("conditional-format-cell");
    expect(screen.getByRole("gridcell", { name: "A4 20" })).not.toHaveClass("conditional-format-cell");

    await user.click(screen.getByRole("button", { name: "Conditional formatting" }));
    const rules = screen.getByLabelText("Conditional format rules");
    expect(within(rules).getByText("Top 2 values")).toBeInTheDocument();
  });

  it("renders conditional formatting data bars for numeric ranges", async () => {
    const user = userEvent.setup();
    render(<App />);

    await user.click(screen.getByRole("gridcell", { name: "A1" }));
    fireEvent.paste(screen.getByRole("grid", { name: "Spreadsheet grid" }), {
      clipboardData: {
        getData: () => "Score\n10\n20\n30"
      }
    });

    selectRange("A2 10", "A4 30");
    await user.click(screen.getByRole("button", { name: "Conditional formatting" }));
    await user.selectOptions(screen.getByLabelText("Conditional rule"), "dataBar");

    expect(screen.queryByLabelText("Conditional value")).not.toBeInTheDocument();
    expect(screen.queryByLabelText("Conditional fill color")).not.toBeInTheDocument();
    expect(screen.getByLabelText("Data bar color")).toBeInTheDocument();

    fireEvent.input(screen.getByLabelText("Data bar color"), { target: { value: "#2f7d9f" } });
    await user.click(screen.getByRole("button", { name: "Apply conditional format" }));

    const firstBar = screen.getByRole("gridcell", { name: "A2 10" }).querySelector<HTMLElement>(".cell-data-bar");
    const lastBar = screen.getByRole("gridcell", { name: "A4 30" }).querySelector<HTMLElement>(".cell-data-bar");

    expect(screen.getByRole("gridcell", { name: "A2 10" })).toHaveClass("conditional-data-bar-cell");
    expect(firstBar).toHaveStyle({ width: "33%" });
    expect(lastBar).toHaveStyle({ width: "100%" });

    await user.click(screen.getByRole("button", { name: "Conditional formatting" }));
    const rules = screen.getByLabelText("Conditional format rules");
    expect(within(rules).getByText("Data bars")).toBeInTheDocument();
  });

  it("renders conditional formatting color scales for numeric ranges", async () => {
    const user = userEvent.setup();
    render(<App />);

    await user.click(screen.getByRole("gridcell", { name: "A1" }));
    fireEvent.paste(screen.getByRole("grid", { name: "Spreadsheet grid" }), {
      clipboardData: {
        getData: () => "Score\n10\n20\n30"
      }
    });

    selectRange("A2 10", "A4 30");
    await user.click(screen.getByRole("button", { name: "Conditional formatting" }));
    await user.selectOptions(screen.getByLabelText("Conditional rule"), "colorScale");

    expect(screen.queryByLabelText("Conditional value")).not.toBeInTheDocument();
    expect(screen.queryByLabelText("Conditional fill color")).not.toBeInTheDocument();
    expect(screen.getByLabelText("Minimum color")).toBeInTheDocument();
    expect(screen.getByLabelText("Maximum color")).toBeInTheDocument();

    fireEvent.input(screen.getByLabelText("Minimum color"), { target: { value: "#ffffff" } });
    fireEvent.input(screen.getByLabelText("Maximum color"), { target: { value: "#000000" } });
    await user.click(screen.getByRole("button", { name: "Apply conditional format" }));

    expect(screen.getByRole("gridcell", { name: "A2 10" })).toHaveClass("conditional-format-cell");
    expect(screen.getByRole("gridcell", { name: "A2 10" })).toHaveStyle({ backgroundColor: "#ffffff" });
    expect(screen.getByRole("gridcell", { name: "A3 20" })).toHaveStyle({ backgroundColor: "#808080" });
    expect(screen.getByRole("gridcell", { name: "A4 30" })).toHaveStyle({ backgroundColor: "#000000" });

    await user.click(screen.getByRole("button", { name: "Conditional formatting" }));
    const rules = screen.getByLabelText("Conditional format rules");
    expect(within(rules).getByText("Color scale")).toBeInTheDocument();
  });

  it("lists and deletes conditional format rules from the panel", async () => {
    const user = userEvent.setup();
    render(<App />);

    await user.click(screen.getByRole("gridcell", { name: "A1" }));
    fireEvent.paste(screen.getByRole("grid", { name: "Spreadsheet grid" }), {
      clipboardData: {
        getData: () => "5\n12\n8"
      }
    });

    selectRange("A1 5", "A3 8");
    await user.click(screen.getByRole("button", { name: "Conditional formatting" }));
    await user.clear(screen.getByLabelText("Conditional value"));
    await user.type(screen.getByLabelText("Conditional value"), "10");
    await user.click(screen.getByRole("button", { name: "Apply conditional format" }));

    expect(screen.getByRole("gridcell", { name: "A2 12" })).toHaveClass("conditional-format-cell");

    await user.click(screen.getByRole("button", { name: "Conditional formatting" }));
    expect(screen.getByRole("heading", { name: "Existing rules" })).toBeInTheDocument();
    expect(screen.getByText("A1:A3")).toBeInTheDocument();
    expect(screen.getByText("Greater than 10")).toBeInTheDocument();

    await user.click(screen.getByRole("button", { name: "Delete conditional format rule A1:A3 Greater than 10" }));

    expect(screen.getByRole("gridcell", { name: "A2 12" })).not.toHaveClass("conditional-format-cell");
    expect(screen.getByText("No conditional format rules")).toBeInTheDocument();
    expect(screen.getByLabelText("Status")).toHaveTextContent("Deleted conditional format rule");
  });

  it("filters a selected range and clears hidden rows", async () => {
    const user = userEvent.setup();
    render(<App />);

    await user.click(screen.getByRole("gridcell", { name: "A1" }));
    fireEvent.paste(screen.getByRole("grid", { name: "Spreadsheet grid" }), {
      clipboardData: {
        getData: () => "Region\tSales\nWest\t10\nEast\t8\nWest\t12"
      }
    });

    selectRange("A1 Region", "B4 12");
    await user.click(screen.getByRole("button", { name: "Filter" }));
    await user.selectOptions(screen.getByLabelText("Filter operator"), "equals");
    await user.clear(screen.getByLabelText("Filter value"));
    await user.type(screen.getByLabelText("Filter value"), "West");
    await user.click(screen.getByRole("button", { name: "Apply filter" }));

    expect(screen.getByRole("gridcell", { name: "A1 Region" })).toBeInTheDocument();
    expect(screen.getByRole("gridcell", { name: "A2 West" })).toBeInTheDocument();
    expect(screen.queryByRole("gridcell", { name: "A3 East" })).not.toBeInTheDocument();
    expect(screen.getByRole("gridcell", { name: "A4 West" })).toBeInTheDocument();
    expect(screen.getByLabelText("Status")).toHaveTextContent("Filter applied");

    await user.click(screen.getByRole("button", { name: "Filter" }));
    await user.click(screen.getByRole("button", { name: "Clear filters" }));

    expect(screen.getByRole("gridcell", { name: "A3 East" })).toBeInTheDocument();
    expect(screen.getByLabelText("Status")).toHaveTextContent("Filters cleared");
  });

  it("creates and removes an embedded chart from the selected range", async () => {
    const user = userEvent.setup();
    render(<App />);

    await user.click(screen.getByRole("gridcell", { name: "A1" }));
    fireEvent.paste(screen.getByRole("grid", { name: "Spreadsheet grid" }), {
      clipboardData: {
        getData: () => "Region\tSales\nWest\t10\nEast\t8\nWest\t12"
      }
    });

    selectRange("A1 Region", "B4 12");
    await openRibbonTab(user, "Insert");
    await user.click(screen.getByRole("button", { name: "Chart" }));
    await user.selectOptions(screen.getByLabelText("Chart type"), "bar");
    await user.type(screen.getByLabelText("Chart title"), "Sales overview");
    await user.click(screen.getByRole("button", { name: "Create chart" }));

    expect(screen.getByRole("figure", { name: "Chart Sales overview" })).toBeInTheDocument();
    expect(screen.getByText("3 points")).toBeInTheDocument();
    expect(screen.getByLabelText("Status")).toHaveTextContent("Created chart");

    await user.click(screen.getByRole("button", { name: "Delete chart Sales overview" }));

    expect(screen.queryByRole("figure", { name: "Chart Sales overview" })).not.toBeInTheDocument();
    expect(screen.getByLabelText("Status")).toHaveTextContent("Deleted chart");
  });

  it("creates an embedded pie chart from the selected range", async () => {
    const user = userEvent.setup();
    render(<App />);

    await user.click(screen.getByRole("gridcell", { name: "A1" }));
    fireEvent.paste(screen.getByRole("grid", { name: "Spreadsheet grid" }), {
      clipboardData: {
        getData: () => "Category\tShare\nProduct\t45\nServices\t35\nSupport\t20"
      }
    });

    selectRange("A1 Category", "B4 20");
    await openRibbonTab(user, "Insert");
    await user.click(screen.getByRole("button", { name: "Chart" }));
    await user.selectOptions(screen.getByLabelText("Chart type"), "pie");
    await user.type(screen.getByLabelText("Chart title"), "Revenue mix");
    await user.click(screen.getByRole("button", { name: "Create chart" }));

    const chart = screen.getByRole("figure", { name: "Chart Revenue mix" });
    expect(chart).toBeInTheDocument();
    expect(within(chart).getByRole("img", { name: "Pie chart" })).toBeInTheDocument();
    expect(within(chart).getByText("3 points")).toBeInTheDocument();
  });

  it("creates a pivot table sheet from the selected source range", async () => {
    const user = userEvent.setup();
    render(<App />);

    await user.click(screen.getByRole("gridcell", { name: "A1" }));
    fireEvent.paste(screen.getByRole("grid", { name: "Spreadsheet grid" }), {
      clipboardData: {
        getData: () => "Region\tProduct\tSales\nWest\tHardware\t10\nWest\tSoftware\t20\nEast\tHardware\t8\nEast\tSoftware\t7"
      }
    });

    selectRange("A1 Region", "C5 7");
    await openRibbonTab(user, "Insert");
    await user.click(screen.getByRole("button", { name: "Pivot table" }));
    await user.click(screen.getByRole("button", { name: "Create pivot table" }));

    expect(screen.getByRole("tab", { name: "Pivot 1" })).toBeInTheDocument();
    expect(screen.getByRole("gridcell", { name: "A1 Region" })).toHaveTextContent("Region");
    expect(screen.getByRole("gridcell", { name: "B1 Hardware" })).toHaveTextContent("Hardware");
    expect(screen.getByRole("gridcell", { name: "A2 East" })).toHaveTextContent("East");
    expect(screen.getByRole("gridcell", { name: "B2 8" })).toHaveTextContent("8");
  });

  it("creates a nested row pivot table from a secondary row field", async () => {
    const user = userEvent.setup();
    render(<App />);

    await user.click(screen.getByRole("gridcell", { name: "A1" }));
    fireEvent.paste(screen.getByRole("grid", { name: "Spreadsheet grid" }), {
      clipboardData: {
        getData: () => "Region\tProduct\tSales\nWest\tHardware\t10\nWest\tSoftware\t20\nEast\tHardware\t8\nEast\tSoftware\t7"
      }
    });

    selectRange("A1 Region", "C5 7");
    await openRibbonTab(user, "Insert");
    await user.click(screen.getByRole("button", { name: "Pivot table" }));
    await user.selectOptions(screen.getByLabelText("Pivot columns"), "");
    await user.selectOptions(screen.getByLabelText("Pivot row detail"), "Product");
    await user.click(screen.getByRole("button", { name: "Create pivot table" }));

    expect(screen.getByRole("tab", { name: "Pivot 1" })).toBeInTheDocument();
    expect(screen.getByRole("gridcell", { name: "A1 Region" })).toHaveTextContent("Region");
    expect(screen.getByRole("gridcell", { name: "B1 Product" })).toHaveTextContent("Product");
    expect(screen.getByRole("gridcell", { name: "C1 SUM of Sales" })).toHaveTextContent("SUM of Sales");
    expect(screen.getByRole("gridcell", { name: "A2 East" })).toHaveTextContent("East");
    expect(screen.getByRole("gridcell", { name: "B2 Hardware" })).toHaveTextContent("Hardware");
    expect(screen.getByRole("gridcell", { name: "C2 8" })).toHaveTextContent("8");
    expect(screen.getByRole("gridcell", { name: "B3 Software" })).toHaveTextContent("Software");
    expect(screen.getByRole("gridcell", { name: "C3 7" })).toHaveTextContent("7");
  });
});

function structuredTableWorkbook(): ReturnType<typeof createBlankWorkbook> {
  const workbook = createBlankWorkbook();
  const sheet = workbook.sheets[0];
  return {
    ...workbook,
    sheets: [{
      ...sheet,
      cells: {
        ...sheet.cells,
        A1: "Region",
        B1: "Sales",
        A2: "West",
        B2: 10,
        A3: "East",
        B3: 8,
        D1: "Department",
        E1: "Cost",
        D2: "Operations",
        E2: 5,
        D3: "Technology",
        E3: 7
      }
    }],
    tables: [
      {
        id: "table-sales",
        name: "SalesTable",
        sheetId: sheet.id,
        range: { start: { row: 0, column: 0 }, end: { row: 2, column: 1 } },
        headerRow: true,
        totalsRow: false,
        columns: [
          { id: "sales-region", name: "Region", sheetColumn: 0 },
          { id: "sales-value", name: "Sales", sheetColumn: 1 }
        ],
        rowIds: ["sales-west", "sales-east"],
        style: { theme: "TableStyleLight1", showRowStripes: true }
      },
      {
        id: "table-costs",
        name: "CostsTable",
        sheetId: sheet.id,
        range: { start: { row: 0, column: 3 }, end: { row: 2, column: 4 } },
        headerRow: true,
        totalsRow: false,
        columns: [
          { id: "cost-department", name: "Department", sheetColumn: 3 },
          { id: "cost-value", name: "Cost", sheetColumn: 4 }
        ],
        rowIds: ["cost-operations", "cost-technology"],
        style: { theme: "TableStyleMedium2", showRowStripes: true }
      }
    ]
  };
}

function columnBoundaryContextWorkbook(): ReturnType<typeof createBlankWorkbook> {
  const workbook = createBlankWorkbook();
  const sheet = workbook.sheets[0];
  const otherSheetId = "sheet-other";
  return {
    ...workbook,
    sheets: [
      {
        ...sheet,
        name: "Active",
        cells: {
          A1: "Region",
          B1: "Sales",
          A2: "West",
          B2: 10,
          A3: "East",
          B3: 8,
          C1: "Department",
          D1: "Cost",
          C2: "Operations",
          D2: 5,
          C3: "Technology",
          D3: 7,
          A5: "Team",
          B5: "Value",
          A6: "North",
          B6: 3,
          A7: "South",
          B7: 4
        }
      },
      {
        ...sheet,
        id: otherSheetId,
        name: "Other",
        cells: {
          A1: "OtherRegion",
          B1: "OtherSales",
          A2: "West",
          B2: 12,
          A3: "East",
          B3: 9
        }
      }
    ],
    tables: [
      {
        id: "table-selected",
        name: "SelectedTable",
        sheetId: sheet.id,
        range: { start: { row: 0, column: 0 }, end: { row: 2, column: 1 } },
        headerRow: true,
        totalsRow: false,
        columns: [
          { id: "selected-region", name: "Region", sheetColumn: 0 },
          { id: "selected-sales", name: "Sales", sheetColumn: 1 }
        ],
        rowIds: ["selected-west", "selected-east"]
      },
      {
        id: "table-adjacent",
        name: "AdjacentTable",
        sheetId: sheet.id,
        range: { start: { row: 0, column: 2 }, end: { row: 2, column: 3 } },
        headerRow: true,
        totalsRow: false,
        columns: [
          { id: "adjacent-department", name: "Department", sheetColumn: 2 },
          { id: "adjacent-cost", name: "Cost", sheetColumn: 3 }
        ],
        rowIds: ["adjacent-operations", "adjacent-technology"]
      },
      {
        id: "table-lower",
        name: "LowerTable",
        sheetId: sheet.id,
        range: { start: { row: 4, column: 0 }, end: { row: 6, column: 1 } },
        headerRow: true,
        totalsRow: false,
        columns: [
          { id: "lower-team", name: "Team", sheetColumn: 0 },
          { id: "lower-value", name: "Value", sheetColumn: 1 }
        ],
        rowIds: ["lower-north", "lower-south"]
      },
      {
        id: "table-other",
        name: "OtherTable",
        sheetId: otherSheetId,
        range: { start: { row: 0, column: 0 }, end: { row: 2, column: 1 } },
        headerRow: true,
        totalsRow: false,
        columns: [
          { id: "other-region", name: "OtherRegion", sheetColumn: 0 },
          { id: "other-sales", name: "OtherSales", sheetColumn: 1 }
        ],
        rowIds: ["other-west", "other-east"]
      }
    ]
  };
}

type FormulaFillOperation = "Fill Down" | "Fill Right" | "AutoFill";

async function applyNumberValidation(
  user: ReturnType<typeof userEvent.setup>,
  address: string,
  bounds: { min?: number; max?: number }
) {
  await user.click(screen.getByRole("gridcell", { name: addressNamePattern(address) }));
  await user.click(screen.getByRole("button", { name: "Data validation" }));
  await user.selectOptions(screen.getByLabelText("Validation type"), "number");
  await user.clear(screen.getByLabelText("Minimum"));
  await user.clear(screen.getByLabelText("Maximum"));
  if (bounds.min !== undefined) {
    await user.type(screen.getByLabelText("Minimum"), String(bounds.min));
  }
  if (bounds.max !== undefined) {
    await user.type(screen.getByLabelText("Maximum"), String(bounds.max));
  }
  await user.click(screen.getByRole("button", { name: "Apply validation" }));
}

async function setupFormulaFill(
  user: ReturnType<typeof userEvent.setup>,
  operation: FormulaFillOperation,
  translatedValue: number
): Promise<{ address: string; formula: string }> {
  if (operation === "Fill Right") {
    await editCell(user, "A2", "5");
    await editCell(user, "B2", String(translatedValue));
    await editCell(user, "B1", "=A2");
    await applyNumberValidation(user, "C1", { max: 10 });
    return { address: "C1", formula: "=B2" };
  }

  await editCell(user, "A1", "5");
  await editCell(user, "A2", String(translatedValue));
  await editCell(user, "B1", "=A1");
  await applyNumberValidation(user, operation === "AutoFill" ? "B1" : "B2", { max: 10 });
  return { address: "B2", formula: "=A2" };
}

async function runFormulaFill(user: ReturnType<typeof userEvent.setup>, operation: FormulaFillOperation) {
  if (operation === "Fill Down") {
    selectRange("B1 5", "B2");
    await user.click(screen.getByRole("button", { name: "Fill down" }));
    return;
  }
  if (operation === "Fill Right") {
    selectRange("B1 5", "C1");
    await user.click(screen.getByRole("button", { name: "Fill right" }));
    return;
  }

  await user.click(screen.getByRole("gridcell", { name: "B1 5" }));
  fireEvent.mouseDown(screen.getByRole("button", { name: "AutoFill selection" }));
  fireEvent.mouseEnter(screen.getByRole("gridcell", { name: "B2" }));
  fireEvent.mouseUp(screen.getByRole("grid", { name: "Spreadsheet grid" }));
}

async function editCell(user: ReturnType<typeof userEvent.setup>, address: string, value: string) {
  await user.dblClick(screen.getByRole("gridcell", { name: addressNamePattern(address) }));
  const editor = screen.getByLabelText(`Cell editor ${address}`);
  await user.clear(editor);
  await user.type(editor, `${value}{Enter}`);
}

async function openRibbonTab(user: ReturnType<typeof userEvent.setup>, name: string) {
  const tab = within(screen.getByRole("tablist", { name: "Ribbon tabs" })).getByRole("tab", { name });
  if (tab.getAttribute("aria-selected") !== "true") {
    await user.click(tab);
  }
}

function addressNamePattern(address: string): RegExp {
  return new RegExp(`^${address.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}(?:\\s|$)`);
}

function sheetTabNames(): string[] {
  return within(screen.getByRole("tablist", { name: "Sheet tabs" }))
    .getAllByRole("tab")
    .map((tab) => tab.textContent ?? "");
}

function selectRange(startName: string, endName: string) {
  const grid = screen.getByRole("grid", { name: "Spreadsheet grid" });
  fireEvent.mouseDown(screen.getByRole("gridcell", { name: startName }));
  fireEvent.mouseEnter(screen.getByRole("gridcell", { name: endName }));
  fireEvent.mouseUp(grid);
}
