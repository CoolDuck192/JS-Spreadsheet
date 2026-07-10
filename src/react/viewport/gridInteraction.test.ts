import { describe, expect, it } from "vitest";
import { createGridInteractionState, reduceGridInteraction } from "./gridInteraction";

const rowIds = ["row-1", "row-2", "row-3"] as const;
const columnIds = ["name", "salary", "active"] as const;
const model = { rowIds, columnIds };

describe("grid interaction reducer", () => {
  it("moves and extends selection using stable ids", () => {
    const initial = createGridInteractionState({ rowId: "row-1", columnId: "name" });
    const moved = reduceGridInteraction(
      initial,
      { type: "move", rowDelta: 1, columnDelta: 1, extend: false },
      model
    );
    const extended = reduceGridInteraction(
      moved,
      { type: "move", rowDelta: 1, columnDelta: 0, extend: true },
      model
    );

    expect(moved.selection).toEqual({
      anchor: { rowId: "row-2", columnId: "salary" },
      focus: { rowId: "row-2", columnId: "salary" }
    });
    expect(extended.selection).toEqual({
      anchor: { rowId: "row-2", columnId: "salary" },
      focus: { rowId: "row-3", columnId: "salary" }
    });
  });

  it("clamps navigation at all four model edges", () => {
    const first = createGridInteractionState({ rowId: "row-1", columnId: "name" });
    const beforeFirst = reduceGridInteraction(
      first,
      { type: "move", rowDelta: -20, columnDelta: -20, extend: false },
      model
    );
    const last = createGridInteractionState({ rowId: "row-3", columnId: "active" });
    const afterLast = reduceGridInteraction(
      last,
      { type: "move", rowDelta: 20, columnDelta: 20, extend: false },
      model
    );

    expect(beforeFirst.selection.focus).toEqual({ rowId: "row-1", columnId: "name" });
    expect(afterLast.selection.focus).toEqual({ rowId: "row-3", columnId: "active" });
  });

  it("selects and extends pointer drags without replacing the stable anchor", () => {
    const initial = createGridInteractionState({ rowId: "row-1", columnId: "name" });
    const started = reduceGridInteraction(
      initial,
      { type: "pointer-down", cell: { rowId: "row-2", columnId: "salary" }, extend: false },
      model
    );
    const dragged = reduceGridInteraction(
      started,
      { type: "pointer-enter", cell: { rowId: "row-3", columnId: "active" } },
      model
    );
    const ended = reduceGridInteraction(dragged, { type: "pointer-up" }, model);

    expect(dragged.selection).toEqual({
      anchor: { rowId: "row-2", columnId: "salary" },
      focus: { rowId: "row-3", columnId: "active" }
    });
    expect(dragged.dragging).toBe(true);
    expect(ended.dragging).toBe(false);
  });

  it("starts, updates, commits, and cancels raw editor state", () => {
    const initial = createGridInteractionState({ rowId: "row-1", columnId: "name" });
    const editing = reduceGridInteraction(initial, { type: "start-edit", rawText: "A" }, model);
    const changed = reduceGridInteraction(editing, { type: "change-edit", rawText: "Ada" }, model);
    const committed = reduceGridInteraction(changed, { type: "commit-edit" }, model);
    const cancelled = reduceGridInteraction(changed, { type: "cancel-edit" }, model);

    expect(changed.editing).toEqual({ rowId: "row-1", columnId: "name", rawText: "Ada" });
    expect(committed.editing).toBeNull();
    expect(cancelled.editing).toBeNull();
    expect(cancelled.selection).toEqual(initial.selection);
  });
});
