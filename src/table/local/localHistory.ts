export type LocalHistoryEntry<TOperation> = { undo: TOperation; redo: TOperation };

export type LocalHistory<TOperation> = {
  limit: number;
  past: readonly LocalHistoryEntry<TOperation>[];
  future: readonly LocalHistoryEntry<TOperation>[];
};

export function createLocalHistory<TOperation>(limit = 100): LocalHistory<TOperation> {
  assertHistoryLimit(limit);
  return { limit, past: [], future: [] };
}

export function pushLocalHistory<TOperation>(
  history: LocalHistory<TOperation>,
  undo: TOperation,
  redo: TOperation
): LocalHistory<TOperation> {
  const past = [...history.past, { undo, redo }];
  return {
    limit: history.limit,
    past: past.slice(Math.max(0, past.length - history.limit)),
    future: []
  };
}

export function undoLocalHistory<TOperation>(history: LocalHistory<TOperation>): {
  history: LocalHistory<TOperation>;
  operation: TOperation | null;
} {
  const entry = history.past.at(-1);
  if (!entry) return { history, operation: null };
  return {
    history: {
      limit: history.limit,
      past: history.past.slice(0, -1),
      future: [...history.future, entry]
    },
    operation: entry.undo
  };
}

export function redoLocalHistory<TOperation>(history: LocalHistory<TOperation>): {
  history: LocalHistory<TOperation>;
  operation: TOperation | null;
} {
  const entry = history.future.at(-1);
  if (!entry) return { history, operation: null };
  return {
    history: {
      limit: history.limit,
      past: [...history.past, entry].slice(-history.limit),
      future: history.future.slice(0, -1)
    },
    operation: entry.redo
  };
}

function assertHistoryLimit(limit: number): void {
  if (!Number.isInteger(limit) || limit < 1 || limit > 1_000) {
    throw new Error("History limit must be an integer from 1 through 1000");
  }
}
