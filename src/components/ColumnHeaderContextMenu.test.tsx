import { fireEvent, render, screen } from "@testing-library/react";
import { afterEach, describe, expect, it, vi } from "vitest";
import { ColumnHeaderContextMenu, type ColumnHeaderContextMenuProps } from "./ColumnHeaderContextMenu";

describe("ColumnHeaderContextMenu", () => {
  afterEach(() => {
    document.querySelectorAll("[data-column-menu-opener]").forEach((element) => element.remove());
    vi.restoreAllMocks();
  });

  it("renders accessible actions, focuses the first item, and keeps dense targets usable", () => {
    const opener = createOpener();
    renderMenu({ opener });

    const menu = screen.getByRole("menu", { name: "Column B context menu" });
    const items = screen.getAllByRole("menuitem");
    expect(menu).toBeInTheDocument();
    expect(items.map((item) => item.textContent)).toEqual([
      "Insert column left",
      "Insert column right",
      "Delete column"
    ]);
    expect(items[0]).toHaveFocus();
    for (const item of items) {
      expect(item).toHaveStyle({ minHeight: "40px", minWidth: "40px" });
    }
  });

  it.each([
    ["Insert column left", "onInsertLeft"],
    ["Insert column right", "onInsertRight"],
    ["Delete column", "onDelete"]
  ] as const)("runs %s, closes, and restores focus", (itemName, callbackName) => {
    const opener = createOpener();
    const onClose = vi.fn();
    const onInsertLeft = vi.fn();
    const onInsertRight = vi.fn();
    const onDelete = vi.fn();
    const callbacks = { onInsertLeft, onInsertRight, onDelete };
    renderMenu({ opener, onClose, ...callbacks });

    fireEvent.click(screen.getByRole("menuitem", { name: itemName }));

    expect(callbacks[callbackName]).toHaveBeenCalledOnce();
    expect(onClose).toHaveBeenCalledOnce();
    expect(opener).toHaveFocus();
  });

  it("closes on Escape and restores focus to the supplied opener", () => {
    const opener = createOpener();
    const onClose = vi.fn();
    renderMenu({ opener, onClose });

    fireEvent.keyDown(screen.getByRole("menuitem", { name: "Insert column left" }), { key: "Escape" });

    expect(onClose).toHaveBeenCalledOnce();
    expect(opener).toHaveFocus();
  });

  it("dismisses on an outside pointer without treating an inside pointer as outside", () => {
    const opener = createOpener();
    const onClose = vi.fn();
    renderMenu({ opener, onClose });

    fireEvent.pointerDown(screen.getByRole("menu", { name: "Column B context menu" }));
    expect(onClose).not.toHaveBeenCalled();

    fireEvent.pointerDown(document.body);
    expect(onClose).toHaveBeenCalledOnce();
    expect(opener).toHaveFocus();
  });

  it("clamps its measured position to an eight-pixel viewport margin and recomputes on resize", () => {
    const width = vi.spyOn(window, "innerWidth", "get").mockReturnValue(320);
    vi.spyOn(window, "innerHeight", "get").mockReturnValue(200);
    vi.spyOn(HTMLElement.prototype, "getBoundingClientRect").mockReturnValue({
      x: 0,
      y: 0,
      top: 0,
      right: 180,
      bottom: 120,
      left: 0,
      width: 180,
      height: 120,
      toJSON: () => ({})
    });
    renderMenu({ opener: createOpener(), x: 500, y: 500 });

    const menu = screen.getByRole("menu", { name: "Column B context menu" });
    expect(menu).toHaveStyle({ left: "132px", top: "72px" });

    width.mockReturnValue(280);
    fireEvent(window, new Event("resize"));
    expect(menu).toHaveStyle({ left: "92px", top: "72px" });
  });
});

function createOpener(): HTMLButtonElement {
  const opener = document.createElement("button");
  opener.dataset.columnMenuOpener = "true";
  opener.textContent = "Column B";
  document.body.append(opener);
  opener.focus();
  return opener;
}

function renderMenu(overrides: Partial<ColumnHeaderContextMenuProps> = {}) {
  const props: ColumnHeaderContextMenuProps = {
    label: "B",
    x: 120,
    y: 60,
    opener: overrides.opener ?? createOpener(),
    onClose: vi.fn(),
    onInsertLeft: vi.fn(),
    onInsertRight: vi.fn(),
    onDelete: vi.fn(),
    ...overrides
  };
  return render(<ColumnHeaderContextMenu {...props} />);
}
