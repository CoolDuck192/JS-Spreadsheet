import { expect, test } from "@playwright/test";
import { exportWorkbookToXlsx } from "../src/lib/xlsx";
import {
  addConditionalFormatRule,
  createBlankWorkbook,
  defineNamedRange,
  setCellContent,
  setCellFormat,
  setCellValidation,
  setRangeReadOnly,
  setSheetFreezePanes,
  setSheetProtection
} from "../src/lib/workbook";

test.beforeEach(async ({ page }) => {
  await page.addInitScript(() => window.localStorage.clear());
});

test("edits cells, recalculates formulas, and manages sheets", async ({ page }) => {
  await page.goto("/");

  await editCell(page, "A1", "10");
  await editCell(page, "A2", "20");
  await editCell(page, "A3", "=SUM(A1:A2)");

  await expect(page.getByRole("gridcell", { name: "A3 30", exact: true })).toBeVisible();
  await page.getByRole("gridcell", { name: "A3 30", exact: true }).click();
  await expect(page.getByLabel("Formula input")).toHaveValue("=SUM(A1:A2)");

  await page.getByRole("button", { name: "Add sheet", exact: true }).click();
  await expect(page.getByRole("tab", { name: "Sheet2", exact: true })).toBeVisible();
  await openRibbonTab(page, "File");
  await expect(page.getByRole("button", { name: "Export CSV", exact: true })).toBeVisible();
});

test("prints the workbook with worksheet chrome", async ({ page }) => {
  await page.addInitScript(() => {
    const printWindow = window as typeof window & { __printCalls: number };
    printWindow.__printCalls = 0;
    window.print = () => {
      printWindow.__printCalls += 1;
    };
  });
  await page.goto("/");

  await editCell(page, "A1", "Printable");
  await openRibbonTab(page, "File");
  await page.getByRole("button", { name: "Print workbook", exact: true }).click();

  await expect
    .poll(() => page.evaluate(() => (window as typeof window & { __printCalls: number }).__printCalls))
    .toBe(1);
  await expect(page.getByLabel("Status", { exact: true })).toContainText("Opened print dialog");

  await page.emulateMedia({ media: "print" });

  await expect(page.getByRole("toolbar", { name: "Toolbar", exact: true })).toBeHidden();
  await expect(page.getByLabel("Formula bar", { exact: true })).toBeHidden();
  await expect(page.getByLabel("Sheet tabs", { exact: true })).toBeHidden();
  await expect(page.getByLabel("Status", { exact: true })).toBeHidden();
  await expect(page.getByRole("gridcell", { name: "A1 Printable", exact: true })).toBeVisible();
});

test("toggles worksheet gridlines from the toolbar", async ({ page }) => {
  await page.goto("/");

  const grid = page.getByRole("grid", { name: "Spreadsheet grid", exact: true });
  const cell = page.getByRole("gridcell", { name: "A1", exact: true });

  await expect(grid).toHaveAttribute("data-gridlines", "visible");
  await expect(cell).toHaveCSS("border-bottom-color", "rgb(221, 229, 238)");

  await openRibbonTab(page, "View");
  await page.getByRole("button", { name: "Hide gridlines", exact: true }).click();

  await expect(grid).toHaveAttribute("data-gridlines", "hidden");
  await expect(cell).toHaveCSS("border-bottom-color", "rgba(0, 0, 0, 0)");
  await expect(page.getByLabel("Status", { exact: true })).toContainText("Gridlines hidden");

  await page.getByRole("button", { name: "Show gridlines", exact: true }).click();

  await expect(grid).toHaveAttribute("data-gridlines", "visible");
  await expect(cell).toHaveCSS("border-bottom-color", "rgb(221, 229, 238)");
  await expect(page.getByLabel("Status", { exact: true })).toContainText("Gridlines shown");
});

test("organizes the ribbon into Excel-style tabs without horizontal overflow", async ({ page }) => {
  await page.goto("/");

  const toolbar = page.getByRole("toolbar", { name: "Toolbar", exact: true });
  const tabs = toolbar.getByRole("tablist", { name: "Ribbon tabs", exact: true });
  await expect(tabs.getByRole("tab")).toHaveText(["File", "Home", "Insert", "Formulas", "Data", "Review", "View"]);
  await expect(tabs.getByRole("tab", { name: "Home", exact: true })).toHaveAttribute("aria-selected", "true");
  await expect(toolbar.getByRole("tabpanel", { name: "Home", exact: true })).toBeVisible();
  await expect(toolbar.getByRole("group", { name: "Clipboard", exact: true })).toBeVisible();
  await expect(toolbar.getByRole("group", { name: "Font", exact: true })).toBeVisible();
  await expect(toolbar.getByRole("button", { name: "Import CSV", exact: true })).toHaveCount(0);

  await tabs.getByRole("tab", { name: "File", exact: true }).click();

  await expect(tabs.getByRole("tab", { name: "File", exact: true })).toHaveAttribute("aria-selected", "true");
  await expect(toolbar.getByRole("tabpanel", { name: "File", exact: true })).toBeVisible();
  await expect(toolbar.getByRole("group", { name: "Workbook", exact: true })).toBeVisible();
  await expect(toolbar.getByRole("button", { name: "Import CSV", exact: true })).toBeVisible();
  await expect(toolbar.getByRole("button", { name: "Export XLSX", exact: true })).toBeVisible();
  await expect(toolbar.getByRole("button", { name: "Bold", exact: true })).toHaveCount(0);

  await tabs.getByRole("tab", { name: "Insert", exact: true }).click();
  await expect(toolbar.getByRole("tabpanel", { name: "Insert", exact: true })).toBeVisible();
  await expect(toolbar.getByRole("button", { name: "Chart", exact: true })).toBeVisible();
  await expect(toolbar.getByRole("button", { name: "Pivot table", exact: true })).toBeVisible();

  const toolbarOverflow = await toolbar.evaluate((element) => ({
    clientWidth: element.clientWidth,
    scrollWidth: element.scrollWidth
  }));
  expect(toolbarOverflow.scrollWidth).toBeLessThanOrEqual(toolbarOverflow.clientWidth + 1);
});

test("toggles worksheet headers from the toolbar", async ({ page }) => {
  await page.goto("/");

  const grid = page.getByRole("grid", { name: "Spreadsheet grid", exact: true });
  const cell = page.getByRole("gridcell", { name: "A1", exact: true });

  await expect(grid).toHaveAttribute("data-headers", "visible");
  await expect(page.getByRole("columnheader", { name: "Column A", exact: true })).toBeVisible();
  await expect(page.getByRole("rowheader", { name: "Row 1", exact: true })).toBeVisible();

  const gridBox = await grid.boundingBox();
  const cellBoxWithHeaders = await cell.boundingBox();
  expect(gridBox).not.toBeNull();
  expect(cellBoxWithHeaders).not.toBeNull();
  expect(cellBoxWithHeaders!.x).toBeGreaterThanOrEqual(gridBox!.x + 48);
  expect(cellBoxWithHeaders!.y).toBeGreaterThanOrEqual(gridBox!.y + 28);

  await openRibbonTab(page, "View");
  await page.getByRole("button", { name: "Hide headers", exact: true }).click();

  await expect(grid).toHaveAttribute("data-headers", "hidden");
  await expect(page.getByRole("columnheader", { name: "Column A", exact: true })).toHaveCount(0);
  await expect(page.getByRole("rowheader", { name: "Row 1", exact: true })).toHaveCount(0);
  await expect(page.getByLabel("Status", { exact: true })).toContainText("Headers hidden");

  const cellBoxWithoutHeaders = await cell.boundingBox();
  expect(cellBoxWithoutHeaders).not.toBeNull();
  expect(cellBoxWithoutHeaders!.x).toBeLessThanOrEqual(gridBox!.x + 1);
  expect(cellBoxWithoutHeaders!.y).toBeLessThanOrEqual(gridBox!.y + 1);

  await page.getByRole("button", { name: "Show headers", exact: true }).click();

  await expect(grid).toHaveAttribute("data-headers", "visible");
  await expect(page.getByRole("columnheader", { name: "Column A", exact: true })).toBeVisible();
  await expect(page.getByRole("rowheader", { name: "Row 1", exact: true })).toBeVisible();
  await expect(page.getByLabel("Status", { exact: true })).toContainText("Headers shown");
});

test("toggles the formula bar while preserving direct cell formula editing", async ({ page }) => {
  await page.goto("/");

  const grid = page.getByRole("grid", { name: "Spreadsheet grid", exact: true });
  const formulaBar = page.getByLabel("Formula bar", { exact: true });

  await expect(formulaBar).toBeVisible();
  const gridBoxWithFormulaBar = await grid.boundingBox();
  expect(gridBoxWithFormulaBar).not.toBeNull();

  await editCell(page, "A1", "10");
  await openRibbonTab(page, "View");
  await page.getByRole("button", { name: "Hide formula bar", exact: true }).click();

  await expect(formulaBar).toHaveCount(0);
  await expect(page.getByLabel("Formula input", { exact: true })).toHaveCount(0);
  await expect(page.getByLabel("Status", { exact: true })).toContainText("Formula bar hidden");

  const gridBoxWithoutFormulaBar = await grid.boundingBox();
  expect(gridBoxWithoutFormulaBar).not.toBeNull();
  expect(gridBoxWithoutFormulaBar!.y).toBeLessThan(gridBoxWithFormulaBar!.y);

  await editCell(page, "A2", "=A1*2");

  await expect(page.getByRole("gridcell", { name: "A2 20", exact: true })).toBeVisible();

  await page.getByRole("button", { name: "Show formula bar", exact: true }).click();
  await page.getByRole("gridcell", { name: "A2 20", exact: true }).click();

  await expect(page.getByLabel("Formula bar", { exact: true })).toBeVisible();
  await expect(page.getByLabel("Formula input", { exact: true })).toHaveValue("=A1*2");
  await expect(page.getByLabel("Status", { exact: true })).toContainText("Formula bar shown");
});

test("shows formula text in worksheet cells without changing formula editing", async ({ page }) => {
  await page.goto("/");

  await editCell(page, "A1", "10");
  await editCell(page, "A2", "20");
  await editCell(page, "A3", "=SUM(A1:A2)");
  await page.getByRole("gridcell", { name: "A3 30", exact: true }).click();

  await expect(page.getByRole("gridcell", { name: "A3 30", exact: true })).toBeVisible();
  await expect(page.getByLabel("Formula input", { exact: true })).toHaveValue("=SUM(A1:A2)");

  await openRibbonTab(page, "View");
  await page.getByRole("button", { name: "Show formulas", exact: true }).click();

  await expect(page.getByRole("button", { name: "Show formula results", exact: true })).toHaveClass(/active-toolbar-button/);
  await expect(page.getByRole("gridcell", { name: "A3 =SUM(A1:A2)", exact: true })).toBeVisible();
  await expect(page.getByRole("gridcell", { name: "A1 10", exact: true })).toBeVisible();
  await expect(page.getByLabel("Formula input", { exact: true })).toHaveValue("=SUM(A1:A2)");
  await expect(page.getByLabel("Status", { exact: true })).toContainText("Formulas shown");

  await page.getByRole("button", { name: "Show formula results", exact: true }).click();

  await expect(page.getByRole("gridcell", { name: "A3 30", exact: true })).toBeVisible();
  await expect(page.getByLabel("Formula input", { exact: true })).toHaveValue("=SUM(A1:A2)");
  await expect(page.getByLabel("Status", { exact: true })).toContainText("Formula results shown");
});

test("audits formula precedents and dependents from the toolbar", async ({ page }) => {
  await page.goto("/");

  await editCell(page, "A1", "10");
  await editCell(page, "A2", "20");
  await editCell(page, "C1", "5");
  await editCell(page, "D1", "=SUM(A1:A2)+C1");

  await page.getByRole("gridcell", { name: "D1 35", exact: true }).click();
  await openRibbonTab(page, "Formulas");
  await page.getByRole("button", { name: "Audit formulas", exact: true }).click();

  const panel = page.getByRole("complementary", { name: "Formula audit", exact: true });
  await expect(panel).toContainText("D1");
  await expect(panel).toContainText("=SUM(A1:A2)+C1");
  await expect(panel.getByRole("button", { name: "Go to A1:A2", exact: true })).toBeVisible();
  await expect(panel.getByRole("button", { name: "Go to C1", exact: true })).toBeVisible();
  await expect(panel.getByText("No dependents", { exact: true })).toBeVisible();
  await expect(page.getByLabel("Status", { exact: true })).toContainText("Formula audit opened");

  await panel.getByRole("button", { name: "Go to A1:A2", exact: true }).click();

  await expect(page.getByRole("gridcell", { name: "A1 10", exact: true })).toHaveAttribute("aria-selected", "true");
  await expect(page.getByRole("gridcell", { name: "A2 20", exact: true })).toHaveAttribute("aria-selected", "true");
  await expect(page.getByLabel("Status", { exact: true })).toContainText("Selected A1:A2");
  await expect(panel.getByRole("button", { name: "Go to D1", exact: true })).toBeVisible();
});

test("toggles sheet tabs while preserving sheet creation", async ({ page }) => {
  await page.goto("/");

  const tabs = page.getByRole("tablist", { name: "Sheet tabs", exact: true });

  await expect(tabs).toBeVisible();

  await openRibbonTab(page, "View");
  await page.getByRole("button", { name: "Hide sheet tabs", exact: true }).click();

  await expect(tabs).toHaveCount(0);
  await expect(page.getByLabel("Status", { exact: true })).toContainText("Sheet tabs hidden");

  await openRibbonTab(page, "Home");
  await page.getByRole("button", { name: "Add sheet", exact: true }).click();

  await expect(tabs).toHaveCount(0);
  await expect(page.getByLabel("Status", { exact: true })).toContainText("Added sheet");

  await openRibbonTab(page, "View");
  await page.getByRole("button", { name: "Show sheet tabs", exact: true }).click();

  await expectSheetTabs(page, ["Sheet1", "Sheet2"]);
  await expect(page.getByRole("tab", { name: "Sheet2", exact: true })).toHaveAttribute("aria-selected", "true");
  await expect(page.getByLabel("Status", { exact: true })).toContainText("Sheet tabs shown");
});

test("hides and restores sheet tabs from the toolbar", async ({ page }) => {
  await page.goto("/");

  await page.getByRole("button", { name: "Add sheet", exact: true }).click();
  await expect(page.getByRole("tab", { name: "Sheet2", exact: true })).toHaveAttribute("aria-selected", "true");

  await page.getByRole("button", { name: "Hide sheet", exact: true }).click();

  await expect(page.getByRole("tab", { name: "Sheet2", exact: true })).toHaveCount(0);
  await expect(page.getByRole("tab", { name: "Sheet1", exact: true })).toHaveAttribute("aria-selected", "true");
  await expect(page.getByLabel("Status", { exact: true })).toContainText("Hid Sheet2");

  await page.getByRole("button", { name: "Hide sheet", exact: true }).click();

  await expect(page.getByRole("tab", { name: "Sheet1", exact: true })).toBeVisible();
  await expect(page.getByLabel("Status", { exact: true })).toContainText("Cannot hide the only visible sheet");

  await page.getByRole("button", { name: "Unhide sheets", exact: true }).click();

  await expect(page.getByRole("tab", { name: "Sheet2", exact: true })).toBeVisible();
  await expect(page.getByLabel("Status", { exact: true })).toContainText("Unhid sheets");
});

test("moves the active sheet left and right from the toolbar", async ({ page }) => {
  await page.goto("/");

  await page.getByRole("button", { name: "Add sheet", exact: true }).click();
  await page.getByRole("button", { name: "Add sheet", exact: true }).click();

  await expectSheetTabs(page, ["Sheet1", "Sheet2", "Sheet3"]);
  await expect(page.getByRole("tab", { name: "Sheet3", exact: true })).toHaveAttribute("aria-selected", "true");

  await page.getByRole("button", { name: "Move sheet left", exact: true }).click();

  await expectSheetTabs(page, ["Sheet1", "Sheet3", "Sheet2"]);
  await expect(page.getByRole("tab", { name: "Sheet3", exact: true })).toHaveAttribute("aria-selected", "true");
  await expect(page.getByLabel("Status", { exact: true })).toContainText("Moved Sheet3 left");

  await page.getByRole("button", { name: "Move sheet left", exact: true }).click();
  await page.getByRole("button", { name: "Move sheet left", exact: true }).click();

  await expectSheetTabs(page, ["Sheet3", "Sheet1", "Sheet2"]);
  await expect(page.getByLabel("Status", { exact: true })).toContainText("Sheet3 is already first");

  await page.getByRole("button", { name: "Move sheet right", exact: true }).click();

  await expectSheetTabs(page, ["Sheet1", "Sheet3", "Sheet2"]);
  await expect(page.getByLabel("Status", { exact: true })).toContainText("Moved Sheet3 right");
});

test("colors active sheet tabs from the toolbar", async ({ page }) => {
  await page.goto("/");

  await setColor(page.getByLabel("Sheet tab color", { exact: true }), "#0f766e");
  await expect(page.getByRole("tab", { name: "Sheet1", exact: true })).toHaveCSS("border-top-color", "rgb(15, 118, 110)");
  await expect(page.getByLabel("Status", { exact: true })).toContainText("Changed Sheet1 tab color");

  await page.getByRole("button", { name: "Add sheet", exact: true }).click();
  await setColor(page.getByLabel("Sheet tab color", { exact: true }), "#7c3aed");

  await expect(page.getByRole("tab", { name: "Sheet1", exact: true })).toHaveCSS("border-top-color", "rgb(15, 118, 110)");
  await expect(page.getByRole("tab", { name: "Sheet2", exact: true })).toHaveCSS("border-top-color", "rgb(124, 58, 237)");
  await expect(page.getByLabel("Status", { exact: true })).toContainText("Changed Sheet2 tab color");
});

test("suggests formulas from the formula bar", async ({ page }) => {
  await page.goto("/");

  await page.getByLabel("Formula input").fill("=");
  const suggestions = page.getByRole("listbox", { name: "Formula suggestions" });
  await expect(suggestions).toBeVisible();
  await expect(suggestions).toHaveAttribute("aria-orientation", "vertical");
  await expect(suggestions.getByRole("option")).toHaveCount(6);
  await expect(suggestions.getByText("SUM(number1, [number2], ...)", { exact: true })).toBeVisible();
  await expect(suggestions.getByText("Adds numbers or ranges.", { exact: true })).toBeVisible();
  const inputBox = await page.getByLabel("Formula input").boundingBox();
  const suggestionBox = await suggestions.boundingBox();
  expect(inputBox).not.toBeNull();
  expect(suggestionBox).not.toBeNull();
  expect(suggestionBox!.y).toBeGreaterThanOrEqual(inputBox!.y + inputBox!.height - 1);
  expect(suggestionBox!.x).toBeGreaterThanOrEqual(inputBox!.x - 1);
  expect(suggestionBox!.width).toBeGreaterThanOrEqual(300);
  expect(suggestionBox!.height).toBeGreaterThanOrEqual(180);
  const formulaPopupStyle = await suggestions.evaluate((element) => {
    const styles = getComputedStyle(element);
    return {
      backgroundAlpha: Number.parseFloat(styles.backgroundColor.match(/rgba?\([^,]+,[^,]+,[^,]+,\s*([^)]+)\)/)?.[1] ?? "1"),
      borderTopColor: styles.borderTopColor,
      maxHeight: Number.parseFloat(styles.maxHeight),
      borderRadius: Number.parseFloat(styles.borderTopLeftRadius),
      boxShadow: styles.boxShadow,
      opacity: Number.parseFloat(styles.opacity)
    };
  });
  expect(formulaPopupStyle.backgroundAlpha).toBeGreaterThanOrEqual(0.95);
  expect(formulaPopupStyle.borderTopColor).toBe("rgb(203, 213, 225)");
  expect(formulaPopupStyle.maxHeight).toBeGreaterThanOrEqual(220);
  expect(formulaPopupStyle.borderRadius).toBeGreaterThanOrEqual(6);
  expect(formulaPopupStyle.boxShadow).not.toBe("none");
  expect(formulaPopupStyle.opacity).toBe(1);
  await page.getByRole("option", { name: /SUM/ }).first().click();

  await expect(page.getByLabel("Formula input")).toHaveValue("=SUM(");
});

test("suggests advanced HyperFormula functions from the formula bar", async ({ page }) => {
  await page.goto("/");

  await page.getByLabel("Formula input").fill("=xl");
  await page.getByRole("option", { name: "XLOOKUP", exact: true }).click();

  await expect(page.getByLabel("Formula input")).toHaveValue("=XLOOKUP(");

  await page.getByLabel("Formula input").fill("=networkdays.");
  await page.getByRole("option", { name: "NETWORKDAYS.INTL", exact: true }).click();

  await expect(page.getByLabel("Formula input")).toHaveValue("=NETWORKDAYS.INTL(");
});

test("opens the function library and inserts a selected function", async ({ page }) => {
  await page.goto("/");

  await openRibbonTab(page, "Formulas");
  await page.getByRole("button", { name: "Function library", exact: true }).click();
  const panel = page.getByRole("complementary", { name: "Function library", exact: true });

  await expect(panel).toBeVisible();
  await expect(panel.getByText(/418 functions/)).toBeVisible();

  await panel.getByLabel("Search functions", { exact: true }).fill("lookup");
  await expect(panel.getByRole("button", { name: "Insert XLOOKUP", exact: true })).toBeVisible();
  await panel.getByRole("button", { name: "Insert XLOOKUP", exact: true }).click();

  await expect(page.getByLabel("Formula input", { exact: true })).toHaveValue("=XLOOKUP(");
  await expect(page.getByLabel("Status", { exact: true })).toContainText("Inserted XLOOKUP function");
});

test("navigates formula suggestions from the keyboard", async ({ page }) => {
  await page.goto("/");

  const formulaInput = page.getByLabel("Formula input");
  await formulaInput.fill("=av");
  await expect(page.getByRole("option", { name: "AVERAGE", exact: true })).toHaveAttribute("aria-selected", "true");

  await formulaInput.press("ArrowDown");
  await expect(page.getByRole("option", { name: "AVEDEV", exact: true })).toHaveAttribute("aria-selected", "true");

  await formulaInput.press("Tab");
  await expect(formulaInput).toHaveValue("=AVEDEV(");

  await page.getByRole("gridcell", { name: "A1", exact: true }).dblclick();
  const cellEditor = page.getByLabel("Cell editor A1");
  await cellEditor.fill("=av");
  await expect(page.getByRole("option", { name: "AVERAGE", exact: true })).toHaveAttribute("aria-selected", "true");

  await cellEditor.press("ArrowDown");
  await expect(page.getByRole("option", { name: "AVEDEV", exact: true })).toHaveAttribute("aria-selected", "true");

  await cellEditor.press("Tab");
  await expect(cellEditor).toHaveValue("=AVEDEV(");
});

test("tracks hovered formula suggestions as the active choice", async ({ page }) => {
  await page.goto("/");

  const formulaInput = page.getByLabel("Formula input");
  await formulaInput.fill("=av");
  await page.getByRole("listbox", { name: "Formula suggestions" }).hover();
  await page.getByRole("option", { name: "AVEDEV", exact: true }).hover();
  await expect(page.getByRole("option", { name: "AVEDEV", exact: true })).toHaveAttribute("aria-selected", "true");
  await expect(formulaInput).toHaveAttribute("aria-activedescendant", "formula-bar-suggestions-option-avedev");
  await formulaInput.press("Tab");
  await expect(formulaInput).toHaveValue("=AVEDEV(");

  await page.getByRole("gridcell", { name: "A1", exact: true }).dblclick();
  const cellEditor = page.getByLabel("Cell editor A1");
  await cellEditor.fill("=av");
  await page.getByRole("listbox", { name: "Formula suggestions" }).hover();
  await page.getByRole("option", { name: "AVEDEV", exact: true }).hover();
  await expect(page.getByRole("option", { name: "AVEDEV", exact: true })).toHaveAttribute("aria-selected", "true");
  await expect(cellEditor).toHaveAttribute("aria-activedescendant", "cell-editor-a1-formula-suggestions-option-avedev");
  await cellEditor.press("Tab");
  await expect(cellEditor).toHaveValue("=AVEDEV(");
});

test("dismisses formula suggestions with Escape without losing typed formulas", async ({ page }) => {
  await page.goto("/");

  const suggestions = page.getByRole("listbox", { name: "Formula suggestions" });
  const formulaInput = page.getByLabel("Formula input");
  await formulaInput.fill("=av");
  await expect(suggestions).toBeVisible();

  await formulaInput.press("Escape");
  await expect(formulaInput).toHaveValue("=av");
  await expect(suggestions).toBeHidden();

  await formulaInput.press("e");
  await expect(suggestions).toBeVisible();

  await page.getByRole("gridcell", { name: "A1", exact: true }).dblclick();
  const cellEditor = page.getByLabel("Cell editor A1");
  await cellEditor.fill("=av");
  await expect(suggestions).toBeVisible();

  await cellEditor.press("Escape");
  await expect(cellEditor).toHaveValue("=av");
  await expect(cellEditor).toBeVisible();
  await expect(suggestions).toBeHidden();

  await cellEditor.press("e");
  await expect(suggestions).toBeVisible();

  await cellEditor.press("Escape");
  await expect(suggestions).toBeHidden();
  await cellEditor.press("Escape");
  await expect(cellEditor).toHaveCount(0);
});

test("imports XLSX workbook files", async ({ page }) => {
  await page.goto("/");

  let workbook = createBlankWorkbook();
  workbook = setCellContent(workbook, workbook.activeSheetId, "A1", "Project");
  workbook = setCellContent(workbook, workbook.activeSheetId, "B1", "=LEN(A1)");
  workbook = setCellContent(workbook, workbook.activeSheetId, "C1", "20");
  workbook = setCellContent(workbook, workbook.activeSheetId, "D1", "10");
  workbook = setCellContent(workbook, workbook.activeSheetId, "D2", "20");
  workbook = setCellContent(workbook, workbook.activeSheetId, "E1", "=SUM(Sales)");
  workbook = setCellContent(workbook, workbook.activeSheetId, "F1", "20");
  workbook = setCellContent(workbook, workbook.activeSheetId, "G1", "Locked");
  workbook = setCellContent(workbook, workbook.activeSheetId, "H1", "Input");
  workbook = defineNamedRange(
    workbook,
    workbook.activeSheetId,
    "Sales",
    { start: { row: 0, column: 3 }, end: { row: 1, column: 3 } }
  );
  workbook = setCellFormat(
    workbook,
    workbook.activeSheetId,
    { start: { row: 0, column: 0 }, end: { row: 0, column: 0 } },
    { bold: true, textColor: "#17634a", backgroundColor: "#eaf7f2", horizontalAlign: "center" }
  );
  workbook = setCellValidation(
    workbook,
    workbook.activeSheetId,
    { start: { row: 0, column: 2 }, end: { row: 0, column: 2 } },
    { type: "number", min: 1, max: 10 }
  );
  workbook = addConditionalFormatRule(
    workbook,
    workbook.activeSheetId,
    { start: { row: 0, column: 5 }, end: { row: 0, column: 5 } },
    {
      condition: { type: "greaterThan", value: "10" },
      format: { backgroundColor: "#fff1d6", textColor: "#8a4b00", bold: true }
    }
  );
  workbook = setRangeReadOnly(
    workbook,
    workbook.activeSheetId,
    { start: { row: 0, column: 0 }, end: { row: 0, column: 0 } },
    false
  );
  workbook = setRangeReadOnly(
    workbook,
    workbook.activeSheetId,
    { start: { row: 0, column: 5 }, end: { row: 0, column: 5 } },
    false
  );
  workbook = setRangeReadOnly(
    workbook,
    workbook.activeSheetId,
    { start: { row: 0, column: 7 }, end: { row: 0, column: 7 } },
    false
  );
  workbook = setSheetFreezePanes(workbook, workbook.activeSheetId, { freezeTopRow: true, freezeFirstColumn: true });
  workbook = setSheetProtection(workbook, workbook.activeSheetId, true);
  const bytes = await exportWorkbookToXlsx(workbook);

  await page.getByLabel("XLSX file").setInputFiles({
    name: "budget.xlsx",
    mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    buffer: Buffer.from(bytes)
  });

  const importedCell = page.getByRole("gridcell", { name: "A1 Project", exact: true });
  await expect(importedCell).toBeVisible();
  await expect(importedCell).toHaveCSS("font-weight", "700");
  await expect(importedCell).toHaveCSS("color", "rgb(23, 99, 74)");
  await expect(importedCell).toHaveCSS("background-color", "rgb(234, 247, 242)");
  await expect(importedCell).toHaveCSS("text-align", "center");
  await expect(page.getByRole("gridcell", { name: "B1 7", exact: true })).toBeVisible();
  await expect(page.getByRole("gridcell", { name: "C1 20", exact: true })).toHaveClass(/invalid-validation-cell/);
  await expect(page.getByRole("gridcell", { name: "E1 30", exact: true })).toBeVisible();
  const conditionalCell = page.getByRole("gridcell", { name: "F1 20", exact: true });
  await expect(conditionalCell).toHaveClass(/conditional-format-cell/);
  await expect(conditionalCell).toHaveCSS("background-color", "rgb(255, 241, 214)");
  await expect(conditionalCell).toHaveCSS("color", "rgb(138, 75, 0)");
  await expect(conditionalCell).toHaveCSS("font-weight", "700");
  await expect(page.getByRole("gridcell", { name: "G1 Locked", exact: true })).toHaveClass(/read-only-cell/);
  await expect(page.getByRole("gridcell", { name: "H1 Input", exact: true })).not.toHaveClass(/read-only-cell/);
  await expect(page.getByRole("button", { name: "Freeze top row", exact: true })).toHaveClass(/active-toolbar-button/);
  await expect(page.getByRole("button", { name: "Freeze first column", exact: true })).toHaveClass(/active-toolbar-button/);
  await expect(page.getByLabel("Status")).toContainText("Imported budget.xlsx");
});

test("suggests formulas while editing directly in a cell", async ({ page }) => {
  await page.goto("/");

  await page.getByRole("gridcell", { name: "A1", exact: true }).dblclick();
  await page.getByLabel("Cell editor A1").fill("=");
  const editor = page.getByLabel("Cell editor A1");
  const suggestions = page.getByRole("listbox", { name: "Formula suggestions" });
  await expect(suggestions).toBeVisible();
  await expect(suggestions.getByText("SUM(number1, [number2], ...)", { exact: true })).toBeVisible();
  const cellSuggestionMainStyle = await suggestions.locator(".formula-suggestion-main").first().evaluate((element) => {
    const styles = getComputedStyle(element);
    return {
      display: styles.display,
      columnGap: Number.parseFloat(styles.columnGap)
    };
  });
  expect(cellSuggestionMainStyle.display).toBe("flex");
  expect(cellSuggestionMainStyle.columnGap).toBeGreaterThanOrEqual(6);

  await editor.fill("=av");
  await expect(suggestions.getByRole("option")).toHaveCount(4);
  const editorBox = await editor.boundingBox();
  const suggestionBox = await suggestions.boundingBox();
  expect(editorBox).not.toBeNull();
  expect(suggestionBox).not.toBeNull();
  expect(suggestionBox!.y).toBeGreaterThanOrEqual(editorBox!.y + editorBox!.height - 1);
  expect(suggestionBox!.height).toBeGreaterThanOrEqual(120);
  expect(suggestionBox!.width).toBeGreaterThanOrEqual(260);
  const cellPopupStyle = await suggestions.evaluate((element) => {
    const styles = getComputedStyle(element);
    return {
      backgroundAlpha: Number.parseFloat(styles.backgroundColor.match(/rgba?\([^,]+,[^,]+,[^,]+,\s*([^)]+)\)/)?.[1] ?? "1"),
      borderTopWidth: styles.borderTopWidth,
      maxHeight: Number.parseFloat(styles.maxHeight),
      borderRadius: Number.parseFloat(styles.borderTopLeftRadius),
      boxShadow: styles.boxShadow,
      opacity: Number.parseFloat(styles.opacity)
    };
  });
  expect(cellPopupStyle.backgroundAlpha).toBeGreaterThanOrEqual(0.95);
  expect(cellPopupStyle.borderTopWidth).toBe("1px");
  expect(cellPopupStyle.maxHeight).toBeGreaterThanOrEqual(220);
  expect(cellPopupStyle.borderRadius).toBeGreaterThanOrEqual(6);
  expect(cellPopupStyle.boxShadow).not.toBe("none");
  expect(cellPopupStyle.opacity).toBe(1);
  await page.keyboard.press("Tab");

  await expect(page.getByLabel("Cell editor A1")).toHaveValue("=AVERAGE(");
});

test("keeps large pasted datasets responsive with bounded rendering and pivot creation", async ({ page }) => {
  test.setTimeout(60000);
  await page.goto("/");

  const rowCount = 1500;
  const regions = ["East", "North", "South", "West"];
  const products = ["Hardware", "Software", "Services"];
  const rows = ["Region\tProduct\tSales\tQuantity\tMonth\tRep"];
  for (let index = 1; index <= rowCount; index += 1) {
    rows.push(
      [
        regions[index % regions.length],
        products[index % products.length],
        String((index % 17) + 1),
        String((index % 5) + 1),
        `Month ${index % 12}`,
        `Rep ${index % 20}`
      ].join("\t")
    );
  }

  await page.getByRole("gridcell", { name: "A1", exact: true }).click();
  const pasteStartedAt = await page.evaluate(() => performance.now());
  await pasteGridData(page, rows.join("\n"));
  await expect(page.getByRole("gridcell", { name: "F1 Rep", exact: true })).toBeVisible();
  const pasteDuration = await page.evaluate((startedAt) => performance.now() - startedAt, pasteStartedAt);
  expect(pasteDuration).toBeLessThan(12000);

  const renderedCells = await page.getByRole("gridcell").count();
  expect(renderedCells).toBeLessThan(1800);

  const grid = page.getByRole("grid", { name: "Spreadsheet grid", exact: true });
  await grid.evaluate((element) => {
    element.scrollTop = element.scrollHeight;
    element.dispatchEvent(new Event("scroll", { bubbles: true }));
  });
  await expect(page.getByRole("gridcell", { name: "A1501 East", exact: true })).toBeVisible();
  expect(await page.getByRole("gridcell").count()).toBeLessThan(1800);

  await page.getByLabel("Name box", { exact: true }).fill("A1:F1501");
  await page.getByLabel("Name box", { exact: true }).press("Enter");
  await openRibbonTab(page, "Insert");
  await page.getByRole("button", { name: "Pivot table", exact: true }).click();
  await page.getByRole("button", { name: "Create pivot table", exact: true }).click();

  await expect(page.getByRole("tab", { name: "Pivot 1", exact: true })).toBeVisible();
  await expect(page.getByRole("gridcell", { name: "A1 Region", exact: true })).toBeVisible();
  await expect(page.getByRole("gridcell", { name: "E1 Grand Total", exact: true })).toBeVisible();
});

test("inserts AutoSum formulas from selected ranges", async ({ page }) => {
  await page.goto("/");

  await page.getByRole("gridcell", { name: "A1", exact: true }).click();
  await pasteGridData(page, "10\n20\n30");
  await selectRange(page, "A1 10", "A3 30");

  await page.getByRole("button", { name: "AutoSum", exact: true }).click();

  await expect(page.getByRole("gridcell", { name: "A4 60", exact: true })).toBeVisible();
  await expect(page.getByLabel("Formula input")).toHaveValue("=SUM(A1:A3)");
  await expect(page.getByLabel("Status")).toContainText("Inserted AutoSum");
});

test("inserts AutoAverage formulas from the toolbar function picker", async ({ page }) => {
  await page.goto("/");

  await page.getByRole("gridcell", { name: "A1", exact: true }).click();
  await pasteGridData(page, "10\n20\n30");
  await selectRange(page, "A1 10", "A3 30");

  await openRibbonTab(page, "Formulas");
  await page.getByLabel("Auto function", { exact: true }).selectOption("AVERAGE");

  await expect(page.getByRole("gridcell", { name: "A4 20", exact: true })).toBeVisible();
  await expect(page.getByLabel("Formula input")).toHaveValue("=AVERAGE(A1:A3)");
  await expect(page.getByLabel("Auto function", { exact: true })).toHaveValue("");
  await expect(page.getByLabel("Status")).toContainText("Inserted Average for A1:A3");
});

test("defines a named range from the name box and uses it in formulas", async ({ page }) => {
  await page.goto("/");

  await page.getByRole("gridcell", { name: "A1", exact: true }).click();
  await pasteGridData(page, "10\n20\n30");
  await selectRange(page, "A1 10", "A3 30");

  await expect(page.getByLabel("Name box")).toHaveValue("A1:A3");
  await page.getByLabel("Name box").fill("Sales");
  await page.keyboard.press("Enter");

  await expect(page.getByLabel("Name box")).toHaveValue("Sales");
  await expect(page.getByLabel("Status")).toContainText("Named A1:A3 as Sales");

  await editCell(page, "B1", "=SUM(Sales)");

  await expect(page.getByRole("gridcell", { name: "B1 60", exact: true })).toBeVisible();
});

test("manages named ranges from the toolbar", async ({ page }) => {
  await page.goto("/");

  await page.getByRole("gridcell", { name: "A1", exact: true }).click();
  await pasteGridData(page, "10\n20\n30");
  await selectRange(page, "A1 10", "A3 30");

  await page.getByLabel("Name box").fill("Sales");
  await page.keyboard.press("Enter");

  await openRibbonTab(page, "Formulas");
  await page.getByRole("button", { name: "Named ranges", exact: true }).click();
  const panel = page.getByRole("complementary", { name: "Named ranges", exact: true });

  await expect(panel.getByText("Sales", { exact: true })).toBeVisible();
  await expect(panel.getByText("Sheet1!A1:A3", { exact: true })).toBeVisible();

  await panel.getByRole("button", { name: "Select Sales", exact: true }).click();

  await expect(page.getByLabel("Name box")).toHaveValue("Sales");
  await expect(page.getByLabel("Status")).toContainText("Selected Sales");

  await panel.getByRole("button", { name: "Delete Sales", exact: true }).click();

  await expect(panel.getByText("Sales", { exact: true })).toHaveCount(0);
  await expect(page.getByLabel("Status")).toContainText("Deleted named range Sales");
});

test("selects typed cell addresses and ranges from the name box", async ({ page }) => {
  await page.goto("/");

  await page.getByLabel("Name box").fill("C5");
  await page.keyboard.press("Enter");

  await expect(page.getByLabel("Name box")).toHaveValue("C5");
  await expect(page.getByRole("gridcell", { name: "C5", exact: true })).toHaveAttribute("aria-selected", "true");
  await expect(page.getByLabel("Status")).toContainText("Selected C5");

  await page.getByLabel("Name box").fill("B2:C3");
  await page.keyboard.press("Enter");

  await expect(page.getByLabel("Name box")).toHaveValue("B2:C3");
  await expect(page.getByRole("gridcell", { name: "B2", exact: true })).toHaveAttribute("aria-selected", "true");
  await expect(page.getByRole("gridcell", { name: "C3", exact: true })).toHaveAttribute("aria-selected", "true");
  await expect(page.getByLabel("Status")).toContainText("Selected B2:C3");
});

test("jumps to references and named ranges from the Go To panel", async ({ page }) => {
  await page.goto("/");

  await page.getByRole("button", { name: "Go to", exact: true }).click();
  const goToPanel = page.getByRole("complementary", { name: "Go to", exact: true });
  await goToPanel.getByLabel("Go to reference", { exact: true }).fill("C5");
  await goToPanel.getByRole("button", { name: "Go", exact: true }).click();

  await expect(page.getByLabel("Name box")).toHaveValue("C5");
  await expect(page.getByRole("gridcell", { name: "C5", exact: true })).toHaveAttribute("aria-selected", "true");
  await expect(page.getByLabel("Status")).toContainText("Selected C5");

  await page.getByRole("gridcell", { name: "A1", exact: true }).click();
  await pasteGridData(page, "10\n20\n30");
  await selectRange(page, "A1 10", "A3 30");
  await page.getByLabel("Name box").fill("Sales");
  await page.keyboard.press("Enter");

  await page.getByRole("button", { name: "Go to", exact: true }).click();
  await expect(goToPanel.getByText("Sales", { exact: true })).toBeVisible();
  await goToPanel.getByRole("button", { name: "Go to Sales", exact: true }).click();

  await expect(page.getByLabel("Name box")).toHaveValue("Sales");
  await expect(page.getByRole("gridcell", { name: "A1 10", exact: true })).toHaveAttribute("aria-selected", "true");
  await expect(page.getByRole("gridcell", { name: "A3 30", exact: true })).toHaveAttribute("aria-selected", "true");
  await expect(page.getByLabel("Status")).toContainText("Selected Sales");
});

test("selects full rows, columns, and sheets from grid headers", async ({ page }) => {
  await page.goto("/");

  await page.getByRole("columnheader", { name: "Column B", exact: true }).click();
  await expect(page.getByLabel("Name box")).toHaveValue("B1:B100");
  await expect(page.getByLabel("Status")).toContainText("100 selected");
  await expect(page.getByRole("gridcell", { name: "B1", exact: true })).toHaveAttribute("aria-selected", "true");

  await page.getByRole("rowheader", { name: "Row 3", exact: true }).click();
  await expect(page.getByLabel("Name box")).toHaveValue("A3:Z3");
  await expect(page.getByLabel("Status")).toContainText("26 selected");
  await expect(page.getByRole("gridcell", { name: "A3", exact: true })).toHaveAttribute("aria-selected", "true");

  await page.getByRole("button", { name: "Select sheet", exact: true }).click();
  await expect(page.getByLabel("Name box")).toHaveValue("A1:Z100");
  await expect(page.getByLabel("Status")).toContainText("2600 selected");
});

test("zooms the worksheet from the status bar", async ({ page }) => {
  await page.goto("/");

  const cell = page.getByRole("gridcell", { name: "A1", exact: true });
  const originalBox = await cell.boundingBox();
  expect(originalBox).not.toBeNull();
  await expect(page.getByRole("button", { name: "Reset zoom", exact: true })).toHaveText("100%");

  await page.getByRole("button", { name: "Zoom in", exact: true }).click();

  await expect(page.getByRole("grid", { name: "Spreadsheet grid", exact: true })).toHaveAttribute("data-zoom-level", "125");
  await expect(page.getByRole("button", { name: "Reset zoom", exact: true })).toHaveText("125%");
  await expect(page.getByLabel("Status", { exact: true })).toContainText("Zoom 125%");
  const zoomedBox = await cell.boundingBox();
  expect(zoomedBox).not.toBeNull();
  expect(zoomedBox!.width).toBeGreaterThan(originalBox!.width);
  expect(zoomedBox!.height).toBeGreaterThan(originalBox!.height);

  await page.getByRole("button", { name: "Reset zoom", exact: true }).click();

  await expect(page.getByRole("grid", { name: "Spreadsheet grid", exact: true })).toHaveAttribute("data-zoom-level", "100");
  await expect(page.getByRole("button", { name: "Reset zoom", exact: true })).toHaveText("100%");
  await expect(page.getByLabel("Status", { exact: true })).toContainText("Zoom reset");
});

test("hides and unhides selected rows and columns", async ({ page }) => {
  await page.goto("/");

  await editCell(page, "A2", "Hidden row");
  await editCell(page, "B1", "Hidden column");

  await page.getByRole("rowheader", { name: "Row 2", exact: true }).click();
  await page.getByRole("button", { name: "Hide rows", exact: true }).click();

  await expect(page.getByRole("rowheader", { name: "Row 2", exact: true })).toHaveCount(0);
  await expect(page.getByRole("gridcell", { name: "A2 Hidden row", exact: true })).toHaveCount(0);
  await expect(page.getByLabel("Status")).toContainText("Hid 1 row");

  await page.getByRole("columnheader", { name: "Column B", exact: true }).click();
  await page.getByRole("button", { name: "Hide columns", exact: true }).click();

  await expect(page.getByRole("columnheader", { name: "Column B", exact: true })).toHaveCount(0);
  await expect(page.getByRole("gridcell", { name: "B1 Hidden column", exact: true })).toHaveCount(0);
  await expect(page.getByLabel("Status")).toContainText("Hid 1 column");

  await page.getByRole("button", { name: "Unhide all", exact: true }).click();

  await expect(page.getByRole("rowheader", { name: "Row 2", exact: true })).toBeVisible();
  await expect(page.getByRole("columnheader", { name: "Column B", exact: true })).toBeVisible();
  await expect(page.getByRole("gridcell", { name: "A2 Hidden row", exact: true })).toBeVisible();
  await expect(page.getByRole("gridcell", { name: "B1 Hidden column", exact: true })).toBeVisible();
});

test("adds and removes cell comments", async ({ page }) => {
  await page.goto("/");

  await page.evaluate(() => {
    const comments = ["Review this assumption", ""];
    window.prompt = () => comments.shift() ?? null;
  });

  await page.getByRole("gridcell", { name: "A1", exact: true }).click();
  await openRibbonTab(page, "Review");
  await page.getByRole("button", { name: "Comment", exact: true }).click();

  await expect(page.getByRole("gridcell", { name: "A1", exact: true })).toHaveClass(/commented-cell/);
  await expect(page.getByRole("gridcell", { name: "A1", exact: true })).toHaveAttribute(
    "title",
    "Comment: Review this assumption"
  );
  await expect(page.getByLabel("Status")).toContainText("Added comment to A1");

  await page.getByRole("button", { name: "Comment", exact: true }).click();

  await expect(page.getByRole("gridcell", { name: "A1", exact: true })).not.toHaveClass(/commented-cell/);
  await expect(page.getByLabel("Status")).toContainText("Removed comment from A1");
});

test("merges and unmerges selected cells", async ({ page }) => {
  await page.goto("/");

  await editCell(page, "A1", "Quarterly report");
  await editCell(page, "B1", "hidden");
  await selectRange(page, "A1 Quarterly report", "B2");

  await page.getByRole("button", { name: "Merge cells", exact: true }).click();

  const mergedCell = page.getByRole("gridcell", { name: "A1 Quarterly report", exact: true });
  await expect(mergedCell).toHaveClass(/merged-cell/);
  await expect(mergedCell).toHaveAttribute("aria-colspan", "2");
  await expect(mergedCell).toHaveAttribute("aria-rowspan", "2");
  await expect(page.getByRole("gridcell", { name: "B1", exact: true })).toHaveClass(/merge-covered-cell/);
  await expect(page.getByLabel("Status")).toContainText("Merged A1:B2");

  const mergedBox = await mergedCell.boundingBox();
  const coveredBox = await page.getByRole("gridcell", { name: "B1", exact: true }).boundingBox();
  expect(mergedBox).not.toBeNull();
  expect(coveredBox).not.toBeNull();
  expect(mergedBox!.width).toBeGreaterThan(coveredBox!.width * 1.8);
  expect(mergedBox!.height).toBeGreaterThan(coveredBox!.height * 1.8);

  await page.getByRole("button", { name: "Unmerge cells", exact: true }).click();

  await expect(page.getByRole("gridcell", { name: "A1 Quarterly report", exact: true })).not.toHaveClass(/merged-cell/);
  await expect(page.getByRole("gridcell", { name: "B1", exact: true })).not.toHaveClass(/merge-covered-cell/);
  await expect(page.getByLabel("Status")).toContainText("Unmerged A1:B2");
});

test("formats cells and creates a pivot table sheet", async ({ page }) => {
  await page.goto("/");

  await editCell(page, "A1", "Styled");
  const styledCell = page.getByRole("gridcell", { name: "A1 Styled", exact: true });
  await styledCell.click();
  await page.getByRole("button", { name: "Bold", exact: true }).click();
  await page.getByRole("button", { name: "Italic", exact: true }).click();
  await setColor(page.getByLabel("Text color"), "#ffffff");
  await setColor(page.getByLabel("Fill color"), "#1f6feb");

  await expect(styledCell).toHaveCSS("font-weight", "700");
  await expect(styledCell).toHaveCSS("font-style", "italic");
  await expect(styledCell).toHaveCSS("color", "rgb(255, 255, 255)");
  await expect(styledCell).toHaveCSS("background-color", "rgb(31, 111, 235)");

  await openRibbonTab(page, "File");
  await page.getByRole("button", { name: "New workbook", exact: true }).click();
  await page.getByRole("gridcell", { name: "A1", exact: true }).click();
  await pasteGridData(
    page,
    "Region\tProduct\tSales\nWest\tHardware\t10\nWest\tSoftware\t20\nEast\tHardware\t8\nEast\tSoftware\t7"
  );

  await expect(page.getByRole("gridcell", { name: "C5 7", exact: true })).toBeVisible();
  await openRibbonTab(page, "Insert");
  await page.getByRole("button", { name: "Pivot table", exact: true }).click();
  await page.getByRole("button", { name: "Create pivot table", exact: true }).click();

  await expect(page.getByRole("tab", { name: "Pivot 1", exact: true })).toBeVisible();
  await expect(page.getByRole("gridcell", { name: "A1 Region", exact: true })).toBeVisible();
  await expect(page.getByRole("gridcell", { name: "B1 Hardware", exact: true })).toBeVisible();
  await expect(page.getByRole("gridcell", { name: "A2 East", exact: true })).toBeVisible();
  await expect(page.getByRole("gridcell", { name: "B2 8", exact: true })).toBeVisible();
});

test("copies cell styling with Format Painter", async ({ page }) => {
  await page.goto("/");

  await editCell(page, "A1", "Styled");
  await editCell(page, "B1", "Target");
  const sourceCell = page.getByRole("gridcell", { name: "A1 Styled", exact: true });
  const targetCell = page.getByRole("gridcell", { name: "B1 Target", exact: true });
  await sourceCell.click();
  await page.getByRole("button", { name: "Bold", exact: true }).click();
  await page.getByRole("button", { name: "Italic", exact: true }).click();
  await setColor(page.getByLabel("Text color"), "#ffffff");
  await setColor(page.getByLabel("Fill color"), "#1f6feb");

  await page.getByRole("button", { name: "Format painter", exact: true }).click();
  await expect(page.getByRole("button", { name: "Format painter", exact: true })).toHaveClass(/active-toolbar-button/);
  await targetCell.click();

  await expect(targetCell).toHaveText("Target");
  await expect(targetCell).toHaveCSS("font-weight", "700");
  await expect(targetCell).toHaveCSS("font-style", "italic");
  await expect(targetCell).toHaveCSS("color", "rgb(255, 255, 255)");
  await expect(targetCell).toHaveCSS("background-color", "rgb(31, 111, 235)");
  await expect(page.getByRole("button", { name: "Format painter", exact: true })).not.toHaveClass(/active-toolbar-button/);
  await expect(page.getByLabel("Status", { exact: true })).toContainText("Painted format to B1");
});

test("creates nested row pivot tables from the builder", async ({ page }) => {
  await page.goto("/");

  await page.getByRole("gridcell", { name: "A1", exact: true }).click();
  await pasteGridData(
    page,
    "Region\tProduct\tSales\nWest\tHardware\t10\nWest\tSoftware\t20\nEast\tHardware\t8\nEast\tSoftware\t7"
  );
  await selectRange(page, "A1 Region", "C5 7");

  await openRibbonTab(page, "Insert");
  await page.getByRole("button", { name: "Pivot table", exact: true }).click();
  await page.getByLabel("Pivot columns").selectOption("");
  await page.getByLabel("Pivot row detail").selectOption("Product");
  await page.getByRole("button", { name: "Create pivot table", exact: true }).click();

  await expect(page.getByRole("tab", { name: "Pivot 1", exact: true })).toBeVisible();
  await expect(page.getByRole("gridcell", { name: "A1 Region", exact: true })).toBeVisible();
  await expect(page.getByRole("gridcell", { name: "B1 Product", exact: true })).toBeVisible();
  await expect(page.getByRole("gridcell", { name: "C1 SUM of Sales", exact: true })).toBeVisible();
  await expect(page.getByRole("gridcell", { name: "A2 East", exact: true })).toBeVisible();
  await expect(page.getByRole("gridcell", { name: "B2 Hardware", exact: true })).toBeVisible();
  await expect(page.getByRole("gridcell", { name: "C2 8", exact: true })).toBeVisible();
  await expect(page.getByRole("gridcell", { name: "B3 Software", exact: true })).toBeVisible();
  await expect(page.getByRole("gridcell", { name: "C3 7", exact: true })).toBeVisible();
});

test("summarizes, sorts, and replaces selected data", async ({ page }) => {
  await page.goto("/");

  await page.getByRole("gridcell", { name: "A1", exact: true }).click();
  await pasteGridData(page, "Delta\t4\nAlpha\t1\nCharlie\t3");
  await selectRange(page, "A1 Delta", "B3 3");

  await expect(page.getByLabel("Status")).toContainText("Count 6");
  await expect(page.getByLabel("Status")).toContainText("Sum 8");
  await expect(page.getByLabel("Status")).toContainText("Min 1");
  await expect(page.getByLabel("Status")).toContainText("Max 4");
  await page.getByRole("button", { name: "Sort A to Z", exact: true }).click();

  await expect(page.getByRole("gridcell", { name: "A1 Alpha", exact: true })).toBeVisible();
  await expect(page.getByRole("gridcell", { name: "B1 1", exact: true })).toBeVisible();
  await expect(page.getByRole("gridcell", { name: "A3 Delta", exact: true })).toBeVisible();

  await page.getByRole("button", { name: "Find and replace", exact: true }).click();
  await page.getByLabel("Find text").fill("Alpha");
  await page.getByLabel("Replace text").fill("Aardvark");
  await page.getByRole("button", { name: "Replace all", exact: true }).click();

  await expect(page.getByRole("gridcell", { name: "A1 Aardvark", exact: true })).toBeVisible();
  await expect(page.getByLabel("Status")).toContainText("Replaced 1 cell");
});

test("sorts AutoFilter table data without moving headers", async ({ page }) => {
  await page.goto("/");

  await page.getByRole("gridcell", { name: "A1", exact: true }).click();
  await pasteGridData(page, "Region\tSales\nWest\t10\nEast\t20\nNorth\t15");
  await selectRange(page, "A1 Region", "B4 15");

  await page.getByRole("button", { name: "Filter", exact: true }).click();
  await page.getByLabel("Filter value").fill("t");
  await page.getByRole("button", { name: "Apply filter", exact: true }).click();
  await selectRange(page, "A1 Region", "B4 15");
  await page.getByRole("button", { name: "Sort A to Z", exact: true }).click();

  await expect(page.getByRole("gridcell", { name: "A1 Region", exact: true })).toBeVisible();
  await expect(page.getByRole("gridcell", { name: "B1 Sales", exact: true })).toBeVisible();
  await expect(page.getByRole("gridcell", { name: "A2 East", exact: true })).toBeVisible();
  await expect(page.getByRole("gridcell", { name: "B2 20", exact: true })).toBeVisible();
  await expect(page.getByRole("gridcell", { name: "A3 North", exact: true })).toBeVisible();
  await expect(page.getByRole("gridcell", { name: "B3 15", exact: true })).toBeVisible();
  await expect(page.getByRole("gridcell", { name: "A4 West", exact: true })).toBeVisible();
  await expect(page.getByRole("gridcell", { name: "B4 10", exact: true })).toBeVisible();
});

test("fills formulas with adjusted relative references", async ({ page }) => {
  await page.goto("/");

  await editCell(page, "A1", "10");
  await editCell(page, "A2", "20");
  await editCell(page, "A3", "30");
  await editCell(page, "B1", "=A1");
  await selectRange(page, "B1 10", "B3");

  await page.getByRole("button", { name: "Fill down", exact: true }).click();

  await expect(page.getByRole("gridcell", { name: "B2 20", exact: true })).toBeVisible();
  await expect(page.getByRole("gridcell", { name: "B3 30", exact: true })).toBeVisible();
  await page.getByRole("gridcell", { name: "B2 20", exact: true }).click();
  await expect(page.getByLabel("Formula input")).toHaveValue("=A2");
});

test("auto-fills a numeric series from the selection handle", async ({ page }) => {
  await page.goto("/");

  await editCell(page, "A1", "1");
  await editCell(page, "A2", "2");
  await selectRange(page, "A1 1", "A2 2");

  await page.getByRole("button", { name: "AutoFill selection", exact: true }).dragTo(
    page.getByRole("gridcell", { name: "A5", exact: true })
  );

  await expect(page.getByRole("gridcell", { name: "A3 3", exact: true })).toBeVisible();
  await expect(page.getByRole("gridcell", { name: "A4 4", exact: true })).toBeVisible();
  await expect(page.getByRole("gridcell", { name: "A5 5", exact: true })).toBeVisible();
  await expect(page.getByLabel("Status", { exact: true })).toContainText("AutoFilled A1:A5");
});

test("auto-fills an ISO date series from the selection handle", async ({ page }) => {
  await page.goto("/");

  // Date entries get a date number format on commit, so cells display the
  // formatted date while the underlying content stays the ISO string.
  await editCell(page, "A1", "2026-06-28");
  await editCell(page, "A2", "2026-06-30");
  await selectRange(page, "A1 Jun 28, 2026", "A2 Jun 30, 2026");

  await page.getByRole("button", { name: "AutoFill selection", exact: true }).dragTo(
    page.getByRole("gridcell", { name: "A4", exact: true })
  );

  await expect(page.getByRole("gridcell", { name: "A3 Jul 2, 2026", exact: true })).toBeVisible();
  await expect(page.getByRole("gridcell", { name: "A4 Jul 4, 2026", exact: true })).toBeVisible();
  await expect(page.getByLabel("Status", { exact: true })).toContainText("AutoFilled A1:A4");
});

test("auto-fills source formatting from the selection handle", async ({ page }) => {
  await page.goto("/");

  await editCell(page, "A1", "North");
  await editCell(page, "A2", "South");
  await selectRange(page, "A1 North", "A2 South");
  await page.getByRole("button", { name: "Bold", exact: true }).click();
  await setColor(page.getByLabel("Fill color"), "#eaf7f2");

  await page.getByRole("button", { name: "AutoFill selection", exact: true }).dragTo(
    page.getByRole("gridcell", { name: "A4", exact: true })
  );

  const filledCell = page.getByRole("gridcell", { name: "A3 North", exact: true });
  await expect(filledCell).toHaveCSS("font-weight", "700");
  await expect(filledCell).toHaveCSS("background-color", "rgb(234, 247, 242)");
});

test("auto-fills month names from the selection handle", async ({ page }) => {
  await page.goto("/");

  await editCell(page, "A1", "Jan");
  await editCell(page, "A2", "Feb");
  await selectRange(page, "A1 Jan", "A2 Feb");

  await page.getByRole("button", { name: "AutoFill selection", exact: true }).dragTo(
    page.getByRole("gridcell", { name: "A5", exact: true })
  );

  await expect(page.getByRole("gridcell", { name: "A3 Mar", exact: true })).toBeVisible();
  await expect(page.getByRole("gridcell", { name: "A4 Apr", exact: true })).toBeVisible();
  await expect(page.getByRole("gridcell", { name: "A5 May", exact: true })).toBeVisible();
});

test("auto-fills upward and clears cells when dragging the handle back inside", async ({ page }) => {
  await page.goto("/");

  await editCell(page, "A3", "3");
  await editCell(page, "A4", "4");
  await selectRange(page, "A3 3", "A4 4");

  await page.getByRole("button", { name: "AutoFill selection", exact: true }).dragTo(
    page.getByRole("gridcell", { name: "A1", exact: true })
  );

  await expect(page.getByRole("gridcell", { name: "A2 2", exact: true })).toBeVisible();
  await expect(page.getByRole("gridcell", { name: "A1 1", exact: true })).toBeVisible();
  await expect(page.getByLabel("Status", { exact: true })).toContainText("AutoFilled A1:A4");

  // Dragging the handle back inside the selection shrinks it, clearing the rest.
  await selectRange(page, "A1 1", "A4 4");
  await page.getByRole("button", { name: "AutoFill selection", exact: true }).dragTo(
    page.getByRole("gridcell", { name: "A2 2", exact: true })
  );

  await expect(page.getByRole("gridcell", { name: "A3", exact: true })).toBeVisible();
  await expect(page.getByRole("gridcell", { name: "A4", exact: true })).toBeVisible();
  await expect(page.getByLabel("Status", { exact: true })).toContainText("Cleared A3:A4");
});

test("dragging from near a selected cell's corner starts a selection, not a fill", async ({ page }) => {
  await page.goto("/");

  await editCell(page, "A5", "precious");
  await editCell(page, "B5", "data");

  await selectRange(page, "D4", "E6");
  await expect(page.getByLabel("Name box", { exact: true })).toHaveValue("D4:E6");

  // Cell drags must not create a native browser text selection: in DOM order it
  // would span unrelated data cells and headers, painting the browser's own
  // highlight over them and hijacking the next drag via native drag-and-drop.
  expect(await page.evaluate(() => window.getSelection()?.toString() ?? "")).toBe("");

  // Press inside E6 a few pixels from its bottom-right corner — close to the
  // fill handle, but still on the cell — and drag left across the data. This
  // must start a NEW selection, not grab the fill handle and overwrite A5:B5.
  const e6 = await page.getByRole("gridcell", { name: "E6", exact: true }).boundingBox();
  const a6 = await page.getByRole("gridcell", { name: "A6", exact: true }).boundingBox();
  expect(e6).not.toBeNull();
  expect(a6).not.toBeNull();
  await page.mouse.move(e6!.x + e6!.width - 7, e6!.y + e6!.height - 7);
  await page.mouse.down();
  await page.mouse.move(a6!.x + a6!.width / 2, a6!.y + a6!.height / 2, { steps: 8 });
  await page.mouse.up();

  await expect(page.getByLabel("Name box", { exact: true })).toHaveValue("A6:E6");
  await expect(page.getByLabel("Status", { exact: true })).not.toContainText("AutoFilled");
  await expect(page.getByRole("gridcell", { name: "A5 precious", exact: true })).toBeVisible();
  await expect(page.getByRole("gridcell", { name: "B5 data", exact: true })).toBeVisible();
});

test("blocks AutoFill when protected target cells are read-only", async ({ page }) => {
  await page.goto("/");

  await editCell(page, "A1", "1");
  await editCell(page, "A2", "2");
  await selectRange(page, "A1 1", "A2 2");
  await openRibbonTab(page, "Review");
  await page.getByRole("button", { name: "Protect sheet", exact: true }).click();

  await page.getByRole("button", { name: "AutoFill selection", exact: true }).dragTo(
    page.getByRole("gridcell", { name: "A5", exact: true })
  );

  await expect(page.getByRole("gridcell", { name: "A3", exact: true })).toBeVisible();
  await expect(page.getByLabel("Status", { exact: true })).toContainText("A3 is read-only");
});

test("pastes internally copied cells with formatting and adjusted formulas", async ({ page }) => {
  await page.goto("/");

  await editCell(page, "A1", "10");
  await editCell(page, "B1", "=A1");
  const sourceCell = page.getByRole("gridcell", { name: "B1 10", exact: true });
  await sourceCell.click();
  await page.getByRole("button", { name: "Bold", exact: true }).click();
  await page.keyboard.press("Control+C");

  await page.getByRole("gridcell", { name: "C1", exact: true }).click();
  await pasteGridData(page, "=A1");

  const pastedCell = page.getByRole("gridcell", { name: "C1 10", exact: true });
  await expect(pastedCell).toHaveCSS("font-weight", "700");
  await pastedCell.click();
  await expect(page.getByLabel("Formula input")).toHaveValue("=B1");
  await expect(page.getByLabel("Status")).toContainText("Pasted cells with formatting");
});

test("uses paste special values and transpose from the toolbar", async ({ page }) => {
  await page.goto("/");

  await editCell(page, "A1", "10");
  await editCell(page, "B1", "=A1");
  await page.getByRole("gridcell", { name: "B1 10", exact: true }).click();
  await page.getByRole("button", { name: "Bold", exact: true }).click();
  await page.keyboard.press("Control+C");

  await page.getByRole("gridcell", { name: "C1", exact: true }).click();
  await page.getByRole("button", { name: "Paste values", exact: true }).click();

  const valuesCell = page.getByRole("gridcell", { name: "C1 10", exact: true });
  await expect(valuesCell).toHaveCSS("font-weight", "400");
  await valuesCell.click();
  await expect(page.getByLabel("Formula input")).toHaveValue("10");
  await expect(page.getByLabel("Status")).toContainText("Pasted values");

  await page.getByRole("gridcell", { name: "A3", exact: true }).click();
  await pasteGridData(page, "A\tB\nC\tD");
  await selectRange(page, "A3 A", "B4 D");
  await page.keyboard.press("Control+C");
  await page.getByRole("gridcell", { name: "D3", exact: true }).click();
  await page.getByRole("button", { name: "Transpose paste", exact: true }).click();

  await expect(page.getByRole("gridcell", { name: "D3 A", exact: true })).toBeVisible();
  await expect(page.getByRole("gridcell", { name: "E3 C", exact: true })).toBeVisible();
  await expect(page.getByRole("gridcell", { name: "D4 B", exact: true })).toBeVisible();
  await expect(page.getByRole("gridcell", { name: "E4 D", exact: true })).toBeVisible();
});

test("adds and removes cell hyperlinks", async ({ page }) => {
  await page.goto("/");

  await editCell(page, "A1", "Report");
  await page.getByRole("gridcell", { name: "A1 Report", exact: true }).click();
  page.once("dialog", async (dialog) => {
    expect(dialog.message()).toBe("Cell link URL");
    await dialog.accept("example.com/report");
  });
  await openRibbonTab(page, "Review");
  await page.getByRole("button", { name: "Link", exact: true }).click();

  const link = page.getByRole("link", { name: "Report", exact: true });
  await expect(link).toHaveAttribute("href", "https://example.com/report");
  await expect(page.getByLabel("Status")).toContainText("Added link to A1");

  await page.getByRole("button", { name: "Unlink", exact: true }).click();

  await expect(page.getByRole("link", { name: "Report", exact: true })).toHaveCount(0);
  await expect(page.getByLabel("Status")).toContainText("Removed link from A1");
});

test("locks and unlocks selected cells", async ({ page }) => {
  await page.goto("/");

  await editCell(page, "A1", "Locked");
  await page.getByRole("gridcell", { name: "A1 Locked", exact: true }).click();
  await openRibbonTab(page, "Review");
  await page.getByRole("button", { name: "Lock cells", exact: true }).click();

  const lockedCell = page.getByRole("gridcell", { name: "A1 Locked", exact: true });
  await expect(lockedCell).toHaveClass(/read-only-cell/);
  await lockedCell.dblclick();
  await expect(page.getByLabel("Cell editor A1")).toHaveCount(0);
  await expect(page.getByLabel("Status")).toContainText("A1 is read-only");

  await page.getByRole("button", { name: "Unlock cells", exact: true }).click();
  await editCell(page, "A1", "Editable");

  await expect(page.getByRole("gridcell", { name: "A1 Editable", exact: true })).not.toHaveClass(/read-only-cell/);
});

test("blocks cut-paste moves when the source becomes protected before paste", async ({ page }) => {
  await page.goto("/");

  await editCell(page, "A1", "Move me");
  await page.getByRole("gridcell", { name: "B1", exact: true }).click();
  await openRibbonTab(page, "Review");
  await page.getByRole("button", { name: "Unlock cells", exact: true }).click();
  await page.getByRole("gridcell", { name: "A1 Move me", exact: true }).click();
  await page.keyboard.press("Control+X");
  await page.getByRole("button", { name: "Protect sheet", exact: true }).click();
  await page.getByRole("gridcell", { name: "B1", exact: true }).click();
  await pasteGridData(page, "Move me");

  await expect(page.getByRole("gridcell", { name: "A1 Move me", exact: true })).toBeVisible();
  await expect(page.getByRole("gridcell", { name: "B1", exact: true })).toBeVisible();
  await expect(page.getByLabel("Status", { exact: true })).toContainText("A1 is read-only");
});

test("inserts, deletes, and freezes spreadsheet structure", async ({ page }) => {
  await page.goto("/");

  await page.getByRole("gridcell", { name: "A1", exact: true }).click();
  await pasteGridData(page, "Name\tAmount\nWest\t10");

  await page.getByRole("gridcell", { name: "A2 West", exact: true }).click();
  await page.getByRole("button", { name: "Insert row above", exact: true }).click();
  await expect(page.getByRole("gridcell", { name: "A2", exact: true })).toBeVisible();
  await expect(page.getByRole("gridcell", { name: "A3 West", exact: true })).toBeVisible();

  await page.getByRole("button", { name: "Delete row", exact: true }).click();
  await expect(page.getByRole("gridcell", { name: "A2 West", exact: true })).toBeVisible();

  await page.getByRole("gridcell", { name: "B1 Amount", exact: true }).click();
  await page.getByRole("button", { name: "Insert column left", exact: true }).click();
  await expect(page.getByRole("gridcell", { name: "B1", exact: true })).toBeVisible();
  await expect(page.getByRole("gridcell", { name: "C1 Amount", exact: true })).toBeVisible();

  await page.getByRole("button", { name: "Delete column", exact: true }).click();
  await expect(page.getByRole("gridcell", { name: "B1 Amount", exact: true })).toBeVisible();

  await page.getByRole("button", { name: "Freeze top row", exact: true }).click();
  await page.getByRole("button", { name: "Freeze first column", exact: true }).click();
  await expect(page.getByRole("button", { name: "Freeze top row", exact: true })).toHaveClass(/active-toolbar-button/);
  await expect(page.getByRole("button", { name: "Freeze first column", exact: true })).toHaveClass(/active-toolbar-button/);

  await page.getByRole("grid", { name: "Spreadsheet grid" }).evaluate((grid) => {
    grid.scrollTop = 900;
    grid.scrollLeft = 400;
    grid.dispatchEvent(new Event("scroll", { bubbles: true }));
  });
  await expect(page.getByRole("gridcell", { name: "A1 Name", exact: true })).toBeVisible();
});

test("resizes columns and rows from header handles", async ({ page }) => {
  await page.goto("/");

  const firstCell = page.getByRole("gridcell", { name: "A1", exact: true });
  const originalCellBox = await firstCell.boundingBox();
  const columnHandle = page.getByRole("button", { name: "Resize column A", exact: true });
  const columnHandleBox = await columnHandle.boundingBox();
  if (!originalCellBox || !columnHandleBox) {
    throw new Error("Could not resolve column resize bounds");
  }

  await page.mouse.move(columnHandleBox.x + columnHandleBox.width / 2, columnHandleBox.y + columnHandleBox.height / 2);
  await page.mouse.down();
  await page.mouse.move(columnHandleBox.x + columnHandleBox.width / 2 + 36, columnHandleBox.y + columnHandleBox.height / 2);
  await page.mouse.up();

  const resizedColumnBox = await firstCell.boundingBox();
  expect(resizedColumnBox?.width).toBeGreaterThan(originalCellBox.width + 28);
  await expect(page.getByLabel("Status")).toContainText("Set column A width");

  const rowHandle = page.getByRole("button", { name: "Resize row 1", exact: true });
  const rowHandleBox = await rowHandle.boundingBox();
  if (!rowHandleBox || !resizedColumnBox) {
    throw new Error("Could not resolve row resize bounds");
  }

  await page.mouse.move(rowHandleBox.x + rowHandleBox.width / 2, rowHandleBox.y + rowHandleBox.height / 2);
  await page.mouse.down();
  await page.mouse.move(rowHandleBox.x + rowHandleBox.width / 2, rowHandleBox.y + rowHandleBox.height / 2 + 16);
  await page.mouse.up();

  const resizedRowBox = await firstCell.boundingBox();
  expect(resizedRowBox?.height).toBeGreaterThan(resizedColumnBox.height + 10);
  await expect(page.getByLabel("Status")).toContainText("Set row 1 height");
});

test("auto-fits selected columns and rows from the toolbar", async ({ page }) => {
  await page.goto("/");

  await editCell(page, "B1", "A much longer customer segment name");
  const targetCell = page.getByRole("gridcell", { name: "B1 A much longer customer segment name", exact: true });
  const originalBox = await targetCell.boundingBox();
  if (!originalBox) {
    throw new Error("Could not resolve original B1 bounds");
  }

  await page.getByRole("columnheader", { name: "Column B", exact: true }).click();
  await page.getByRole("button", { name: "Auto-fit columns", exact: true }).click();

  const autoFitColumnBox = await targetCell.boundingBox();
  expect(autoFitColumnBox?.width).toBeGreaterThan(originalBox.width + 120);
  await expect(page.getByLabel("Status")).toContainText("Auto-fit 1 column");

  const rowHandle = page.getByRole("button", { name: "Resize row 1", exact: true });
  const rowHandleBox = await rowHandle.boundingBox();
  if (!rowHandleBox || !autoFitColumnBox) {
    throw new Error("Could not resolve row resize bounds");
  }

  await page.mouse.move(rowHandleBox.x + rowHandleBox.width / 2, rowHandleBox.y + rowHandleBox.height / 2);
  await page.mouse.down();
  await page.mouse.move(rowHandleBox.x + rowHandleBox.width / 2, rowHandleBox.y + rowHandleBox.height / 2 + 32);
  await page.mouse.up();

  const tallerRowBox = await targetCell.boundingBox();
  expect(tallerRowBox?.height).toBeGreaterThan(autoFitColumnBox.height + 24);

  await page.getByRole("rowheader", { name: "Row 1", exact: true }).click();
  await page.getByRole("button", { name: "Auto-fit rows", exact: true }).click();

  const autoFitRowBox = await targetCell.boundingBox();
  expect(autoFitRowBox?.height).toBeLessThan(tallerRowBox!.height - 24);
  await expect(page.getByLabel("Status")).toContainText("Auto-fit 1 row");
});

test("applies number formats and cell alignment", async ({ page }) => {
  await page.goto("/");

  await editCell(page, "A1", "1234.5");
  await page.getByRole("gridcell", { name: "A1 1234.5", exact: true }).click();
  await page.getByLabel("Number format").selectOption("currency");
  await page.getByLabel("Horizontal align").selectOption("right");
  await page.getByLabel("Vertical align").selectOption("bottom");

  const currencyCell = page.getByRole("gridcell", { name: "A1 $1,234.50", exact: true });
  await expect(currencyCell).toBeVisible();
  await expect(currencyCell).toHaveCSS("text-align", "right");
  await expect(currencyCell).toHaveCSS("align-items", "flex-end");

  await editCell(page, "B1", "0.25");
  await page.getByRole("gridcell", { name: "B1 0.25", exact: true }).click();
  await page.getByLabel("Number format").selectOption("percent");

  await expect(page.getByRole("gridcell", { name: "B1 25%", exact: true })).toBeVisible();
});

test("wraps and unwraps text from the toolbar", async ({ page }) => {
  await page.goto("/");

  await editCell(page, "A1", "A long status note that should wrap inside the cell");
  const cell = page.getByRole("gridcell", { name: "A1 A long status note that should wrap inside the cell", exact: true });
  await cell.click();

  await page.getByRole("button", { name: "Wrap text", exact: true }).click();

  await expect(cell).toHaveClass(/wrapped-cell/);
  await expect(cell).toHaveCSS("white-space", "normal");
  await expect(page.getByLabel("Status", { exact: true })).toContainText("Wrapped text");

  await page.getByRole("button", { name: "Wrap text", exact: true }).click();

  await expect(cell).not.toHaveClass(/wrapped-cell/);
  await expect(page.getByLabel("Status", { exact: true })).toContainText("Unwrapped text");
});

test("uses keyboard shortcuts for formatting links and filters", async ({ page }) => {
  await page.goto("/");

  await editCell(page, "A1", "Report");
  const cell = page.getByRole("gridcell", { name: "A1 Report", exact: true });
  await cell.click();

  await page.keyboard.press("Control+B");
  await page.keyboard.press("Control+I");
  await expect(cell).toHaveCSS("font-weight", "700");
  await expect(cell).toHaveCSS("font-style", "italic");
  await expect(page.getByRole("button", { name: "Bold", exact: true })).toHaveClass(/active-toolbar-button/);
  await expect(page.getByRole("button", { name: "Italic", exact: true })).toHaveClass(/active-toolbar-button/);

  page.once("dialog", async (dialog) => {
    expect(dialog.message()).toBe("Cell link URL");
    await dialog.accept("example.com/report");
  });
  await page.keyboard.press("Control+K");
  await expect(page.getByRole("link", { name: "Report", exact: true })).toHaveAttribute("href", "https://example.com/report");

  await page.keyboard.press("Control+Shift+L");
  await expect(page.getByRole("complementary", { name: "Filter", exact: true })).toBeVisible();
});

test("uses the cell context menu for common cell actions", async ({ page }) => {
  await page.goto("/");

  await editCell(page, "A1", "Context note");
  const cell = page.getByRole("gridcell", { name: "A1 Context note", exact: true });
  await cell.click({ button: "right" });

  const menu = page.getByRole("menu", { name: "Cell context menu", exact: true });
  await expect(menu).toBeVisible();
  await expect(menu).toContainText("A1");

  await menu.getByRole("menuitem", { name: "Wrap text", exact: true }).click();

  await expect(menu).toHaveCount(0);
  await expect(cell).toHaveClass(/wrapped-cell/);
  await expect(page.getByLabel("Status", { exact: true })).toContainText("Wrapped text");

  await cell.click({ button: "right" });
  await page.getByRole("menu", { name: "Cell context menu", exact: true }).getByRole("menuitem", { name: "Clear contents", exact: true }).click();

  await expect(page.getByRole("gridcell", { name: "A1", exact: true })).toBeVisible();
  await expect(page.getByLabel("Status", { exact: true })).toContainText("Cleared selection");
});

test("cuts and pastes cells from the context menu", async ({ page }) => {
  await page.goto("/");

  await editCell(page, "A1", "Move me");
  await page.getByRole("gridcell", { name: "A1 Move me", exact: true }).click({ button: "right" });
  await page.getByRole("menu", { name: "Cell context menu", exact: true }).getByRole("menuitem", { name: "Cut", exact: true }).click();

  await expect(page.getByLabel("Status", { exact: true })).toContainText("Cut selection");

  await page.getByRole("gridcell", { name: "B1", exact: true }).click({ button: "right" });
  await page.getByRole("menu", { name: "Cell context menu", exact: true }).getByRole("menuitem", { name: "Paste", exact: true }).click();

  await expect(page.getByRole("gridcell", { name: "A1", exact: true })).toBeVisible();
  await expect(page.getByRole("gridcell", { name: "B1 Move me", exact: true })).toBeVisible();
  await expect(page.getByLabel("Status", { exact: true })).toContainText("Moved selection");
});

test("inserts and deletes rows and columns from the context menu", async ({ page }) => {
  await page.goto("/");

  await editCell(page, "A1", "Name");
  await editCell(page, "B1", "Amount");
  await editCell(page, "A2", "West");

  await page.getByRole("gridcell", { name: "A2 West", exact: true }).click({ button: "right" });
  await expectCellContextMenuInsideViewport(page);
  await page.getByRole("menu", { name: "Cell context menu", exact: true }).getByRole("menuitem", { name: "Insert row above", exact: true }).click();
  await expect(page.getByRole("gridcell", { name: "A2", exact: true })).toBeVisible();
  await expect(page.getByRole("gridcell", { name: "A3 West", exact: true })).toBeVisible();

  await page.getByRole("gridcell", { name: "A2", exact: true }).click({ button: "right" });
  await page.getByRole("menu", { name: "Cell context menu", exact: true }).getByRole("menuitem", { name: "Delete row", exact: true }).click();
  await expect(page.getByRole("gridcell", { name: "A2 West", exact: true })).toBeVisible();

  await page.getByRole("gridcell", { name: "B1 Amount", exact: true }).click({ button: "right" });
  await page.getByRole("menu", { name: "Cell context menu", exact: true }).getByRole("menuitem", { name: "Insert column left", exact: true }).click();
  await expect(page.getByRole("gridcell", { name: "B1", exact: true })).toBeVisible();
  await expect(page.getByRole("gridcell", { name: "C1 Amount", exact: true })).toBeVisible();

  await page.getByRole("gridcell", { name: "B1", exact: true }).click({ button: "right" });
  await page.getByRole("menu", { name: "Cell context menu", exact: true }).getByRole("menuitem", { name: "Delete column", exact: true }).click();
  await expect(page.getByRole("gridcell", { name: "B1 Amount", exact: true })).toBeVisible();
});

test("applies and clears cell borders", async ({ page }) => {
  await page.goto("/");

  await editCell(page, "A1", "Bordered");
  await page.getByRole("gridcell", { name: "A1 Bordered", exact: true }).click();
  await page.getByLabel("Borders").selectOption("all");

  const borderedCell = page.getByRole("gridcell", { name: "A1 Bordered", exact: true });
  await expect(borderedCell).toHaveCSS("border-top-style", "solid");
  await expect(borderedCell).toHaveCSS("border-right-style", "solid");
  await expect(page.getByLabel("Status")).toContainText("Applied all borders to A1");

  await page.getByLabel("Borders").selectOption("none");

  await expect(borderedCell).toHaveCSS("border-top-style", "none");
  await expect(page.getByLabel("Status")).toContainText("Cleared borders from A1");
});

test("clears formats from the context menu without clearing content", async ({ page }) => {
  await page.goto("/");

  await editCell(page, "A1", "Styled");
  const cell = page.getByRole("gridcell", { name: "A1 Styled", exact: true });
  await cell.click();
  await page.getByRole("button", { name: "Bold", exact: true }).click();

  await expect(cell).toHaveCSS("font-weight", "700");

  await cell.click({ button: "right" });
  await page.getByRole("menu", { name: "Cell context menu", exact: true }).getByRole("menuitem", { name: "Clear formats", exact: true }).click();

  await expect(page.getByRole("gridcell", { name: "A1 Styled", exact: true })).toBeVisible();
  await expect(page.getByRole("gridcell", { name: "A1 Styled", exact: true })).not.toHaveCSS("font-weight", "700");
  await expect(page.getByLabel("Status", { exact: true })).toContainText("Cleared formats");
});

test("clears conditional formats from the context menu without clearing direct formats", async ({ page }) => {
  await page.goto("/");

  await editCell(page, "A1", "12");
  const cell = page.getByRole("gridcell", { name: "A1 12", exact: true });
  await cell.click();
  await page.getByRole("button", { name: "Italic", exact: true }).click();
  await page.getByRole("button", { name: "Conditional formatting", exact: true }).click();
  await page.getByLabel("Conditional rule").selectOption("greaterThan");
  await page.getByLabel("Conditional value").fill("10");
  await page.getByRole("button", { name: "Apply conditional format", exact: true }).click();

  await expect(cell).toHaveClass(/conditional-format-cell/);
  await expect(cell).toHaveCSS("background-color", "rgb(255, 241, 214)");
  await expect(cell).toHaveCSS("font-style", "italic");

  await cell.click({ button: "right" });
  await page
    .getByRole("menu", { name: "Cell context menu", exact: true })
    .getByRole("menuitem", { name: "Clear conditional formats", exact: true })
    .click();

  await expect(page.getByRole("gridcell", { name: "A1 12", exact: true })).toBeVisible();
  await expect(page.getByRole("gridcell", { name: "A1 12", exact: true })).not.toHaveClass(/conditional-format-cell/);
  await expect(page.getByRole("gridcell", { name: "A1 12", exact: true })).toHaveCSS("font-style", "italic");
  await expect(page.getByLabel("Status", { exact: true })).toContainText("Cleared conditional formatting");
});

test("clears hyperlinks from the context menu without clearing content", async ({ page }) => {
  await page.goto("/");

  await editCell(page, "A1", "Report");
  await page.getByRole("gridcell", { name: "A1 Report", exact: true }).click();
  page.once("dialog", async (dialog) => {
    expect(dialog.message()).toBe("Cell link URL");
    await dialog.accept("example.com/report");
  });
  await openRibbonTab(page, "Review");
  await page.getByRole("button", { name: "Link", exact: true }).click();
  await expect(page.getByRole("link", { name: "Report", exact: true })).toHaveAttribute("href", "https://example.com/report");

  await page.getByRole("gridcell", { name: "A1 Report", exact: true }).click({ button: "right" });
  await page.getByRole("menu", { name: "Cell context menu", exact: true }).getByRole("menuitem", { name: "Clear hyperlinks", exact: true }).click();

  await expect(page.getByRole("link", { name: "Report", exact: true })).toHaveCount(0);
  await expect(page.getByRole("gridcell", { name: "A1 Report", exact: true })).toBeVisible();
  await expect(page.getByLabel("Status", { exact: true })).toContainText("Cleared hyperlinks");
});

test("clears data validation from the context menu without clearing content", async ({ page }) => {
  await page.goto("/");

  await editCell(page, "A1", "20");
  await page.getByRole("gridcell", { name: "A1 20", exact: true }).click();
  await page.getByRole("button", { name: "Data validation", exact: true }).click();
  await page.getByLabel("Validation type").selectOption("number");
  await page.getByLabel("Minimum").fill("abc");
  await expect(page.getByRole("button", { name: "Apply validation", exact: true })).toBeDisabled();
  await page.getByLabel("Minimum").fill("");
  await page.getByLabel("Minimum").fill("10");
  await page.getByLabel("Maximum").fill("1");
  await expect(page.getByRole("button", { name: "Apply validation", exact: true })).toBeDisabled();
  await page.getByLabel("Minimum").fill("1");
  await page.getByLabel("Maximum").fill("10");
  await page.getByRole("button", { name: "Apply validation", exact: true }).click();

  const invalidCell = page.getByRole("gridcell", { name: "A1 20", exact: true });
  await expect(invalidCell).toHaveClass(/invalid-validation-cell/);

  await invalidCell.click({ button: "right" });
  await page.getByRole("menu", { name: "Cell context menu", exact: true }).getByRole("menuitem", { name: "Clear validation", exact: true }).click();

  await expect(page.getByRole("gridcell", { name: "A1 20", exact: true })).toBeVisible();
  await expect(page.getByRole("gridcell", { name: "A1 20", exact: true })).not.toHaveClass(/invalid-validation-cell/);
  await expect(page.getByLabel("Status", { exact: true })).toContainText("Cleared data validation");
});

test("chooses list validation values from the selected cell dropdown", async ({ page }) => {
  await page.goto("/");

  await page.getByRole("gridcell", { name: "A1", exact: true }).click();
  await page.getByRole("button", { name: "Data validation", exact: true }).click();
  await page.getByLabel("List values").fill("Open, Closed, Blocked");
  await page.getByRole("button", { name: "Apply validation", exact: true }).click();

  await page.getByRole("button", { name: "Open validation choices for A1", exact: true }).click();
  await page.getByRole("listbox", { name: "Validation choices for A1", exact: true }).getByRole("option", { name: "Closed", exact: true }).click();

  await expect(page.getByRole("gridcell", { name: "A1 Closed", exact: true })).toBeVisible();
  await expect(page.getByRole("listbox", { name: "Validation choices for A1", exact: true })).toHaveCount(0);

  await page.getByRole("button", { name: "Data validation", exact: true }).click();
  const validationRules = page.getByLabel("Data validation rules", { exact: true });
  await expect(page.getByRole("heading", { name: "Existing rules", exact: true })).toBeVisible();
  await expect(validationRules.getByText("A1", { exact: true })).toBeVisible();
  await expect(validationRules.getByText("List: Open, Closed, Blocked", { exact: true })).toBeVisible();

  await page.getByRole("button", { name: "Delete data validation rule A1 List: Open, Closed, Blocked", exact: true }).click();

  await expect(page.getByRole("button", { name: "Open validation choices for A1", exact: true })).toHaveCount(0);
  await expect(page.getByText("No data validation rules", { exact: true })).toBeVisible();
  await expect(page.getByLabel("Status", { exact: true })).toContainText("Deleted data validation rule");
});

test("validates text length from the data validation panel", async ({ page }) => {
  await page.goto("/");

  await page.getByRole("gridcell", { name: "A1", exact: true }).click();
  await page.getByRole("button", { name: "Data validation", exact: true }).click();
  await page.getByLabel("Validation type").selectOption("textLength");
  await page.getByLabel("Minimum length").fill("2");
  await page.getByLabel("Maximum length").fill("5");
  await page.getByRole("button", { name: "Apply validation", exact: true }).click();

  await page.getByRole("gridcell", { name: "A1", exact: true }).dblclick();
  await page.getByLabel("Cell editor A1").fill("Western");
  await page.keyboard.press("Enter");

  await expect(page.getByLabel("Cell editor A1")).toHaveValue("Western");
  await expect(page.getByLabel("Status", { exact: true })).toContainText("Enter text between 2 and 5 characters");

  await page.getByLabel("Cell editor A1").fill("West");
  await page.keyboard.press("Enter");
  await expect(page.getByRole("gridcell", { name: "A1 West", exact: true })).not.toHaveClass(/invalid-validation-cell/);

  await page.getByRole("button", { name: "Data validation", exact: true }).click();
  const validationRules = page.getByLabel("Data validation rules", { exact: true });
  await expect(validationRules.getByText("A1", { exact: true })).toBeVisible();
  await expect(validationRules.getByText("Text length: 2 to 5", { exact: true })).toBeVisible();
});

test("clears all content, formats, and hyperlinks from the context menu", async ({ page }) => {
  await page.goto("/");

  await editCell(page, "A1", "Report");
  const cell = page.getByRole("gridcell", { name: "A1 Report", exact: true });
  await cell.click();
  await page.getByRole("button", { name: "Bold", exact: true }).click();
  page.once("dialog", async (dialog) => {
    expect(dialog.message()).toBe("Cell link URL");
    await dialog.accept("example.com/report");
  });
  await openRibbonTab(page, "Review");
  await page.getByRole("button", { name: "Link", exact: true }).click();

  await expect(page.getByRole("link", { name: "Report", exact: true })).toHaveAttribute("href", "https://example.com/report");
  await expect(cell).toHaveCSS("font-weight", "700");

  await cell.click({ button: "right" });
  await page.getByRole("menu", { name: "Cell context menu", exact: true }).getByRole("menuitem", { name: "Clear all", exact: true }).click();

  await expect(page.getByRole("gridcell", { name: "A1", exact: true })).toBeVisible();
  await expect(page.getByRole("link", { name: "Report", exact: true })).toHaveCount(0);
  await expect(page.getByLabel("Status", { exact: true })).toContainText("Cleared all");
});

test("clears comments from the context menu without clearing content", async ({ page }) => {
  await page.goto("/");

  await editCell(page, "A1", "Forecast");
  await page.getByRole("gridcell", { name: "A1 Forecast", exact: true }).click();
  page.once("dialog", async (dialog) => {
    expect(dialog.message()).toBe("Cell comment");
    await dialog.accept("Review the forecast");
  });
  await openRibbonTab(page, "Review");
  await page.getByRole("button", { name: "Comment", exact: true }).click();

  const cell = page.getByRole("gridcell", { name: "A1 Forecast", exact: true });
  await expect(cell).toHaveClass(/commented-cell/);
  await expect(cell).toHaveAttribute("title", "Comment: Review the forecast");

  await cell.click({ button: "right" });
  await expectCellContextMenuInsideViewport(page);
  await page.getByRole("menu", { name: "Cell context menu", exact: true }).getByRole("menuitem", { name: "Clear comments", exact: true }).click();

  await expect(page.getByRole("gridcell", { name: "A1 Forecast", exact: true })).toBeVisible();
  await expect(page.getByRole("gridcell", { name: "A1 Forecast", exact: true })).not.toHaveClass(/commented-cell/);
  await expect(page.getByLabel("Status", { exact: true })).toContainText("Cleared comments");
});

test("applies conditional formatting to matching selected cells", async ({ page }) => {
  await page.goto("/");

  await page.getByRole("gridcell", { name: "A1", exact: true }).click();
  await pasteGridData(page, "5\n12\n8");
  await selectRange(page, "A1 5", "A3 8");

  await page.getByRole("button", { name: "Conditional formatting", exact: true }).click();
  await page.getByLabel("Conditional rule").selectOption("greaterThan");
  await page.getByLabel("Conditional value").fill("10");
  await page.getByRole("button", { name: "Apply conditional format", exact: true }).click();

  const matchingCell = page.getByRole("gridcell", { name: "A2 12", exact: true });
  await expect(matchingCell).toHaveClass(/conditional-format-cell/);
  await expect(matchingCell).toHaveCSS("background-color", "rgb(255, 241, 214)");
  await expect(matchingCell).toHaveCSS("color", "rgb(138, 75, 0)");
  await expect(page.getByRole("gridcell", { name: "A1 5", exact: true })).not.toHaveClass(/conditional-format-cell/);

  await page.getByRole("button", { name: "Conditional formatting", exact: true }).click();
  await expect(page.getByRole("heading", { name: "Existing rules", exact: true })).toBeVisible();
  await expect(page.getByText("A1:A3", { exact: true })).toBeVisible();
  await expect(page.getByText("Greater than 10", { exact: true })).toBeVisible();

  await page.getByRole("button", { name: "Delete conditional format rule A1:A3 Greater than 10", exact: true }).click();

  await expect(matchingCell).not.toHaveClass(/conditional-format-cell/);
  await expect(page.getByText("No conditional format rules", { exact: true })).toBeVisible();
  await expect(page.getByLabel("Status", { exact: true })).toContainText("Deleted conditional format rule");
});

test("highlights blank cells with conditional formatting", async ({ page }) => {
  await page.goto("/");

  await page.getByRole("gridcell", { name: "A1", exact: true }).click();
  await pasteGridData(page, "Task\tStatus\nBuild\tDone\nTest\t");
  await selectRange(page, "B1 Status", "B3");

  await page.getByRole("button", { name: "Conditional formatting", exact: true }).click();
  await page.getByLabel("Conditional rule").selectOption("blank");

  await expect(page.getByLabel("Conditional value")).toHaveCount(0);
  await expect(page.getByRole("button", { name: "Apply conditional format", exact: true })).toBeEnabled();

  await page.getByRole("button", { name: "Apply conditional format", exact: true }).click();

  const blankCell = page.getByRole("gridcell", { name: "B3", exact: true });
  await expect(blankCell).toHaveClass(/conditional-format-cell/);
  await expect(page.getByRole("gridcell", { name: "B2 Done", exact: true })).not.toHaveClass(/conditional-format-cell/);

  await page.getByRole("button", { name: "Conditional formatting", exact: true }).click();
  const rules = page.getByLabel("Conditional format rules", { exact: true });
  await expect(rules.getByText("B1:B3", { exact: true })).toBeVisible();
  await expect(rules.getByText("Blank", { exact: true })).toBeVisible();
});

test("highlights duplicate values with conditional formatting", async ({ page }) => {
  await page.goto("/");

  await page.getByRole("gridcell", { name: "A1", exact: true }).click();
  await pasteGridData(page, "Region\nWest\nEast\nWest");
  await selectRange(page, "A1 Region", "A4 West");

  await page.getByRole("button", { name: "Conditional formatting", exact: true }).click();
  await page.getByLabel("Conditional rule").selectOption("duplicate");

  await expect(page.getByLabel("Conditional value")).toHaveCount(0);
  await expect(page.getByRole("button", { name: "Apply conditional format", exact: true })).toBeEnabled();

  await page.getByRole("button", { name: "Apply conditional format", exact: true }).click();

  await expect(page.getByRole("gridcell", { name: "A2 West", exact: true })).toHaveClass(/conditional-format-cell/);
  await expect(page.getByRole("gridcell", { name: "A4 West", exact: true })).toHaveClass(/conditional-format-cell/);
  await expect(page.getByRole("gridcell", { name: "A3 East", exact: true })).not.toHaveClass(/conditional-format-cell/);

  await page.getByRole("button", { name: "Conditional formatting", exact: true }).click();
  const rules = page.getByLabel("Conditional format rules", { exact: true });
  await expect(rules.getByText("A1:A4", { exact: true })).toBeVisible();
  await expect(rules.getByText("Duplicate values", { exact: true })).toBeVisible();
});

test("highlights top ranked values with conditional formatting", async ({ page }) => {
  await page.goto("/");

  await page.getByRole("gridcell", { name: "A1", exact: true }).click();
  await pasteGridData(page, "Score\n10\n30\n20\n30");
  await selectRange(page, "A1 Score", "A5 30");

  await page.getByRole("button", { name: "Conditional formatting", exact: true }).click();
  await page.getByLabel("Conditional rule").selectOption("top");
  await page.getByLabel("Conditional rank").fill("2");

  await expect(page.getByLabel("Conditional value")).toHaveCount(0);
  await expect(page.getByRole("button", { name: "Apply conditional format", exact: true })).toBeEnabled();

  await page.getByRole("button", { name: "Apply conditional format", exact: true }).click();

  await expect(page.getByRole("gridcell", { name: "A3 30", exact: true })).toHaveClass(/conditional-format-cell/);
  await expect(page.getByRole("gridcell", { name: "A5 30", exact: true })).toHaveClass(/conditional-format-cell/);
  await expect(page.getByRole("gridcell", { name: "A4 20", exact: true })).not.toHaveClass(/conditional-format-cell/);

  await page.getByRole("button", { name: "Conditional formatting", exact: true }).click();
  const rules = page.getByLabel("Conditional format rules", { exact: true });
  await expect(rules.getByText("A1:A5", { exact: true })).toBeVisible();
  await expect(rules.getByText("Top 2 values", { exact: true })).toBeVisible();
});

test("renders conditional formatting data bars for numeric ranges", async ({ page }) => {
  await page.goto("/");

  await page.getByRole("gridcell", { name: "A1", exact: true }).click();
  await pasteGridData(page, "Score\n10\n20\n30");
  await selectRange(page, "A2 10", "A4 30");

  await page.getByRole("button", { name: "Conditional formatting", exact: true }).click();
  await page.getByLabel("Conditional rule").selectOption("dataBar");

  await expect(page.getByLabel("Conditional value")).toHaveCount(0);
  await expect(page.getByLabel("Conditional fill color")).toHaveCount(0);
  await expect(page.getByLabel("Data bar color")).toBeVisible();

  await page.getByLabel("Data bar color").fill("#2f7d9f");
  await page.getByRole("button", { name: "Apply conditional format", exact: true }).click();

  const firstCell = page.getByRole("gridcell", { name: "A2 10", exact: true });
  const lastCell = page.getByRole("gridcell", { name: "A4 30", exact: true });

  await expect(firstCell).toHaveClass(/conditional-data-bar-cell/);
  await expect(firstCell.locator(".cell-data-bar")).toHaveAttribute("style", /width:\s*33%;/);
  await expect(lastCell.locator(".cell-data-bar")).toHaveAttribute("style", /width:\s*100%;/);

  await page.getByRole("button", { name: "Conditional formatting", exact: true }).click();
  const rules = page.getByLabel("Conditional format rules", { exact: true });
  await expect(rules.getByText("Data bars", { exact: true })).toBeVisible();
});

test("renders conditional formatting color scales for numeric ranges", async ({ page }) => {
  await page.goto("/");

  await page.getByRole("gridcell", { name: "A1", exact: true }).click();
  await pasteGridData(page, "Score\n10\n20\n30");
  await selectRange(page, "A2 10", "A4 30");

  await page.getByRole("button", { name: "Conditional formatting", exact: true }).click();
  await page.getByLabel("Conditional rule").selectOption("colorScale");

  await expect(page.getByLabel("Conditional value")).toHaveCount(0);
  await expect(page.getByLabel("Conditional fill color")).toHaveCount(0);
  await expect(page.getByLabel("Minimum color")).toBeVisible();
  await expect(page.getByLabel("Maximum color")).toBeVisible();

  await page.getByLabel("Minimum color").fill("#ffffff");
  await page.getByLabel("Maximum color").fill("#000000");
  await page.getByRole("button", { name: "Apply conditional format", exact: true }).click();

  await page.getByRole("gridcell", { name: "B1", exact: true }).click();

  const lowCell = page.getByRole("gridcell", { name: "A2 10", exact: true });
  const midCell = page.getByRole("gridcell", { name: "A3 20", exact: true });
  const highCell = page.getByRole("gridcell", { name: "A4 30", exact: true });

  await expect(lowCell).toHaveClass(/conditional-format-cell/);
  await expect(lowCell).toHaveCSS("background-color", "rgb(255, 255, 255)");
  await expect(midCell).toHaveCSS("background-color", "rgb(128, 128, 128)");
  await expect(highCell).toHaveCSS("background-color", "rgb(0, 0, 0)");

  await page.getByRole("button", { name: "Conditional formatting", exact: true }).click();
  const rules = page.getByLabel("Conditional format rules", { exact: true });
  await expect(rules.getByText("Color scale", { exact: true })).toBeVisible();
});

test("filters selected table rows and clears filters", async ({ page }) => {
  await page.goto("/");

  await page.getByRole("gridcell", { name: "A1", exact: true }).click();
  await pasteGridData(page, "Region\tSales\nWest\t10\nEast\t8\nNorth\t11\nWest\t12");
  await selectRange(page, "A1 Region", "B5 12");

  await page.getByRole("button", { name: "Filter", exact: true }).click();
  await page.getByLabel("Filter operator").selectOption("equals");
  await page.getByLabel("Filter value").fill("West");
  await page.getByRole("button", { name: "Apply filter", exact: true }).click();

  await expect(page.getByRole("gridcell", { name: "A1 Region", exact: true })).toBeVisible();
  await expect(page.getByRole("gridcell", { name: "A2 West", exact: true })).toBeVisible();
  await expect(page.getByRole("gridcell", { name: "A3 East", exact: true })).toHaveCount(0);
  await expect(page.getByRole("gridcell", { name: "A5 West", exact: true })).toBeVisible();

  await page.getByRole("button", { name: "Filter", exact: true }).click();
  await page.getByRole("button", { name: "Clear filters", exact: true }).click();
  await expect(page.getByRole("gridcell", { name: "A3 East", exact: true })).toBeVisible();
});

test("removes duplicate rows from the selected range", async ({ page }) => {
  await page.goto("/");

  await page.getByRole("gridcell", { name: "A1", exact: true }).click();
  await pasteGridData(page, "East\tHardware\nWest\tSoftware\nEast\tHardware\nNorth\tServices\nWest\tSoftware");
  await selectRange(page, "A1 East", "B5 Software");

  await page.getByRole("button", { name: "Remove duplicates", exact: true }).click();

  await expect(page.getByRole("gridcell", { name: "A1 East", exact: true })).toBeVisible();
  await expect(page.getByRole("gridcell", { name: "B1 Hardware", exact: true })).toBeVisible();
  await expect(page.getByRole("gridcell", { name: "A2 West", exact: true })).toBeVisible();
  await expect(page.getByRole("gridcell", { name: "B2 Software", exact: true })).toBeVisible();
  await expect(page.getByRole("gridcell", { name: "A3 North", exact: true })).toBeVisible();
  await expect(page.getByRole("gridcell", { name: "B3 Services", exact: true })).toBeVisible();
  await expect(page.getByRole("gridcell", { name: "A4", exact: true })).toBeVisible();
  await expect(page.getByRole("gridcell", { name: "B5", exact: true })).toBeVisible();
  await expect(page.getByLabel("Status", { exact: true })).toContainText("Removed 2 duplicate rows");
});

test("filters table data from AutoFilter header menus", async ({ page }) => {
  await page.goto("/");

  await page.getByRole("gridcell", { name: "A1", exact: true }).click();
  await pasteGridData(page, "Region\tSales\nWest\t10\nEast\t8\nNorth\t11\nWest\t12");
  await selectRange(page, "A1 Region", "B5 12");

  await page.getByRole("button", { name: "Filter", exact: true }).click();
  await page.getByLabel("Filter operator").selectOption("equals");
  await page.getByLabel("Filter value").fill("West");
  await page.getByRole("button", { name: "Apply filter", exact: true }).click();

  await page.getByRole("button", { name: "Open AutoFilter menu for Region", exact: true }).click();
  await page.getByRole("menu", { name: "AutoFilter menu for Region", exact: true }).getByRole("menuitem", { name: "Clear filter from Region", exact: true }).click();
  await expect(page.getByRole("gridcell", { name: "A3 East", exact: true })).toBeVisible();

  await page.getByRole("button", { name: "Open AutoFilter menu for Region", exact: true }).click();
  const menu = page.getByRole("menu", { name: "AutoFilter menu for Region", exact: true });
  await menu.getByRole("menuitemcheckbox", { name: "East", exact: true }).click();
  await menu.getByRole("menuitemcheckbox", { name: "West", exact: true }).click();
  await menu.getByRole("menuitem", { name: "Apply selected values", exact: true }).click();

  await expect(page.getByRole("gridcell", { name: "A1 Region", exact: true })).toBeVisible();
  await expect(page.getByRole("gridcell", { name: "A2 West", exact: true })).toBeVisible();
  await expect(page.getByRole("gridcell", { name: "A3 East", exact: true })).toBeVisible();
  await expect(page.getByRole("gridcell", { name: "A4 North", exact: true })).toHaveCount(0);
  await expect(page.getByRole("gridcell", { name: "A5 West", exact: true })).toBeVisible();
  await expect(page.getByRole("menu", { name: "AutoFilter menu for Region", exact: true })).toHaveCount(0);
});

test("creates and deletes an embedded chart from selected data", async ({ page }) => {
  await page.goto("/");

  await page.getByRole("gridcell", { name: "A1", exact: true }).click();
  await pasteGridData(page, "Region\tSales\nWest\t10\nEast\t8\nWest\t12");
  await selectRange(page, "A1 Region", "B4 12");

  await openRibbonTab(page, "Insert");
  await page.getByRole("button", { name: "Chart", exact: true }).click();
  await page.getByLabel("Chart type").selectOption("line");
  await page.getByLabel("Chart title").fill("Sales overview");
  await page.getByRole("button", { name: "Create chart", exact: true }).click();

  await expect(page.getByRole("figure", { name: "Chart Sales overview", exact: true })).toBeVisible();
  await expect(page.getByText("3 points")).toBeVisible();

  await page.getByRole("button", { name: "Delete chart Sales overview", exact: true }).click();
  await expect(page.getByRole("figure", { name: "Chart Sales overview", exact: true })).toHaveCount(0);
});

test("creates an embedded pie chart from selected data", async ({ page }) => {
  await page.goto("/");

  await page.getByRole("gridcell", { name: "A1", exact: true }).click();
  await pasteGridData(page, "Category\tShare\nProduct\t45\nServices\t35\nSupport\t20");
  await selectRange(page, "A1 Category", "B4 20");

  await openRibbonTab(page, "Insert");
  await page.getByRole("button", { name: "Chart", exact: true }).click();
  await page.getByLabel("Chart type").selectOption("pie");
  await page.getByLabel("Chart title").fill("Revenue mix");
  await page.getByRole("button", { name: "Create chart", exact: true }).click();

  const chart = page.getByRole("figure", { name: "Chart Revenue mix", exact: true });
  await expect(chart).toBeVisible();
  await expect(chart.getByRole("img", { name: "Pie chart", exact: true })).toBeVisible();
  await expect(chart.getByText("3 points", { exact: true })).toBeVisible();
});

async function editCell(page: import("@playwright/test").Page, address: string, value: string) {
  await page.getByRole("gridcell", { name: addressNamePattern(address) }).dblclick();
  await page.getByLabel(`Cell editor ${address}`).fill(value);
  await page.keyboard.press("Enter");
}

function addressNamePattern(address: string): RegExp {
  return new RegExp(`^${address.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}(?:\\s|$)`);
}

async function setColor(locator: import("@playwright/test").Locator, value: string) {
  await locator.evaluate((element, nextValue) => {
    const input = element as HTMLInputElement;
    input.value = nextValue;
    input.dispatchEvent(new Event("input", { bubbles: true }));
    input.dispatchEvent(new Event("change", { bubbles: true }));
  }, value);
}

async function expectSheetTabs(page: import("@playwright/test").Page, names: string[]) {
  await expect(page.getByRole("tablist", { name: "Sheet tabs", exact: true }).getByRole("tab")).toHaveText(names);
}

async function expectCellContextMenuInsideViewport(page: import("@playwright/test").Page) {
  const menu = page.getByRole("menu", { name: "Cell context menu", exact: true });
  await expect(menu).toBeVisible();
  const menuBox = await menu.boundingBox();
  const viewport = page.viewportSize();
  expect(menuBox).not.toBeNull();
  expect(viewport).not.toBeNull();
  expect(menuBox!.x).toBeGreaterThanOrEqual(0);
  expect(menuBox!.y).toBeGreaterThanOrEqual(0);
  expect(menuBox!.x + menuBox!.width).toBeLessThanOrEqual(viewport!.width);
  expect(menuBox!.y + menuBox!.height).toBeLessThanOrEqual(viewport!.height);
}

async function openRibbonTab(page: import("@playwright/test").Page, name: string) {
  const tab = page.getByRole("tablist", { name: "Ribbon tabs", exact: true }).getByRole("tab", { name, exact: true });
  if ((await tab.getAttribute("aria-selected")) !== "true") {
    await tab.click();
  }
}

async function pasteGridData(page: import("@playwright/test").Page, text: string) {
  await page.evaluate((clipboardText) => {
    const grid = document.querySelector('[role="grid"][aria-label="Spreadsheet grid"]');
    if (!grid) {
      throw new Error("Spreadsheet grid not found");
    }

    const data = new DataTransfer();
    data.setData("text/plain", clipboardText);
    grid.dispatchEvent(new ClipboardEvent("paste", { clipboardData: data, bubbles: true, cancelable: true }));
  }, text);
}

async function selectRange(page: import("@playwright/test").Page, startName: string, endName: string) {
  const start = page.getByRole("gridcell", { name: startName, exact: true });
  const end = page.getByRole("gridcell", { name: endName, exact: true });
  const startBox = await start.boundingBox();
  const endBox = await end.boundingBox();
  if (!startBox || !endBox) {
    throw new Error("Could not resolve range bounds");
  }

  await page.mouse.move(startBox.x + startBox.width / 2, startBox.y + startBox.height / 2);
  await page.mouse.down();
  await page.mouse.move(endBox.x + endBox.width / 2, endBox.y + endBox.height / 2);
  await page.mouse.up();
}
