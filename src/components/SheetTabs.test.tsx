import { render } from "@testing-library/react";
import { afterEach, describe, expect, it, vi } from "vitest";
import { createBlankWorkbook } from "../lib/workbook";
import { SheetTabs } from "./SheetTabs";

const originalScrollIntoView = Object.getOwnPropertyDescriptor(
  HTMLElement.prototype,
  "scrollIntoView"
);

describe("SheetTabs", () => {
  afterEach(() => {
    if (originalScrollIntoView) {
      Object.defineProperty(HTMLElement.prototype, "scrollIntoView", originalScrollIntoView);
    } else {
      Reflect.deleteProperty(HTMLElement.prototype, "scrollIntoView");
    }
  });

  it("scrolls only the active sheet tab with nearest alignment", () => {
    const scrollIntoView = vi.fn();
    Object.defineProperty(HTMLElement.prototype, "scrollIntoView", {
      configurable: true,
      value: scrollIntoView
    });
    const firstSheet = createBlankWorkbook().sheets[0];
    const sheets = [
      firstSheet,
      { ...firstSheet, id: "sheet-2", name: "Sheet2" }
    ];
    const { rerender } = render(
      <SheetTabs sheets={sheets} activeSheetId={firstSheet.id} onSelect={vi.fn()} onAdd={vi.fn()} />
    );

    scrollIntoView.mockClear();
    rerender(
      <SheetTabs sheets={sheets} activeSheetId="sheet-2" onSelect={vi.fn()} onAdd={vi.fn()} />
    );

    expect(scrollIntoView).toHaveBeenCalledTimes(1);
    expect(scrollIntoView).toHaveBeenCalledWith({ block: "nearest", inline: "nearest" });
  });
});
