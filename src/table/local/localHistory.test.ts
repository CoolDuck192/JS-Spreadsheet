import { describe, expect, it } from "vitest";
import { createLocalHistory, pushLocalHistory, redoLocalHistory, undoLocalHistory } from "./localHistory";

describe("local table history", () => {
  it("bounds entries, clears future after a new edit, and returns exact inverses", () => {
    let history = createLocalHistory<string>(2);
    history = pushLocalHistory(history, "undo-1", "redo-1");
    history = pushLocalHistory(history, "undo-2", "redo-2");
    history = pushLocalHistory(history, "undo-3", "redo-3");
    expect(history.past.map((entry) => entry.undo)).toEqual(["undo-2", "undo-3"]);

    const undone = undoLocalHistory(history);
    expect(undone.operation).toBe("undo-3");
    const redone = redoLocalHistory(undone.history);
    expect(redone.operation).toBe("redo-3");

    const replaced = pushLocalHistory(undone.history, "undo-4", "redo-4");
    expect(replaced.future).toEqual([]);
  });

  it.each([0, -1, 1.5, 1_001])("rejects an invalid history limit: %s", (limit) => {
    expect(() => createLocalHistory(limit)).toThrow("integer from 1 through 1000");
  });
});
