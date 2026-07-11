import { expect, test } from "@playwright/test";

test("keeps the rightmost column menu inside desktop and narrow viewports", async ({ page }) => {
  for (const viewport of [{ width: 1280, height: 800 }, { width: 390, height: 844 }]) {
    await page.setViewportSize(viewport);
    await page.goto("/datatable");
    const grid = page.getByRole("grid", { name: "Employee directory", exact: true });
    await grid.evaluate((node) => {
      node.scrollLeft = node.scrollWidth;
    });

    const trigger = page.getByRole("button", { name: "Column options for Active", exact: true });
    await trigger.click();
    const menu = page.getByRole("menu", { name: "Active column menu", exact: true });
    await expect(menu).toBeVisible();
    const box = await menu.boundingBox();
    expect(box).not.toBeNull();
    expect(box!.x).toBeGreaterThanOrEqual(8);
    expect(box!.y).toBeGreaterThanOrEqual(8);
    expect(box!.x + box!.width).toBeLessThanOrEqual(viewport.width - 8);
    expect(box!.y + box!.height).toBeLessThanOrEqual(viewport.height - 8);

    await page.keyboard.press("Escape");
    await expect(trigger).toBeFocused();
  }
});

test("collapses a tree row on the first narrow-screen click without scrolling", async ({ page }) => {
  await page.setViewportSize({ width: 390, height: 844 });
  await page.goto("/datatable");
  const grid = page.getByRole("grid", { name: "Employee directory", exact: true });
  const before = await grid.evaluate((node) => node.scrollLeft);

  await page.getByRole("button", { name: "Collapse Engineering team", exact: true }).click();

  await expect(page.getByRole("gridcell", { name: /e-ada Employee/ })).toHaveCount(0);
  expect(await grid.evaluate((node) => node.scrollLeft)).toBe(before);

  await page.getByRole("button", { name: "Expand Engineering team", exact: true }).click();
  await expect(page.getByRole("gridcell", { name: /e-ada Employee/ })).toBeVisible();
});
