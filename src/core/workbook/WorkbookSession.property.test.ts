import fc from "fast-check";
import { describe, expect, it } from "vitest";
import type { CellContent, WorkbookModel } from "../../types";
import type { FormulaEngine } from "../../lib/formulaEngine";
import { createBlankWorkbook, getCellContent } from "../../lib/workbook";
import { createWorkbookSession } from "./WorkbookSession";

const PROPERTY_SEED = 20260709;
const PROPERTY_RUNS = 1000;

type PropertyOptions = {
  seed: number;
  path: string | undefined;
  numRuns: number;
  endOnFailure: true;
};

export function createPropertyOptions(
  environment: { readonly FC_REPLAY_PATH?: string } = process.env
): PropertyOptions {
  const replayPath = environment.FC_REPLAY_PATH;
  return {
    seed: PROPERTY_SEED,
    path: replayPath,
    numRuns: PROPERTY_RUNS,
    endOnFailure: true
  };
}

describe("WorkbookSession seeded edit/undo sequences", () => {
  it("locks the deterministic seed, optional replay path, run count, and failure mode", () => {
    expect(createPropertyOptions({ FC_REPLAY_PATH: undefined })).toEqual({
      seed: 20260709,
      path: undefined,
      numRuns: 1000,
      endOnFailure: true
    });
    expect(createPropertyOptions({ FC_REPLAY_PATH: "17:4:2" })).toEqual({
      seed: 20260709,
      path: "17:4:2",
      numRuns: 1000,
      endOnFailure: true
    });
  });

  it("keeps workbook and bounded undo/redo state equivalent to a reference model", () => {
    const operation = fc.oneof(
      fc.record({
        type: fc.constant("edit" as const),
        address: fc.constantFrom("A1", "A2", "B1", "B2"),
        input: fc.oneof(fc.integer({ min: -20, max: 20 }).map(String), fc.constant(""))
      }),
      fc.constant({ type: "undo" as const }),
      fc.constant({ type: "redo" as const })
    );

    const editUndoProperty = fc.property(
      fc.array(operation, { minLength: 0, maxLength: 40 }),
      (operations) => {
        const session = createWorkbookSession({
          workbook: createBlankWorkbook(),
          formulaEngineFactory: createRawFormulaEngine,
          history: { maxEntries: 8, maxWeight: 10_000 }
        });
        const sheetId = session.getSnapshot().workbook.activeSheetId;
        let model: ReferenceHistory = { past: [], present: {}, future: [] };

        for (const operation of operations) {
          if (operation.type === "edit") {
            const next = applyReferenceEdit(model.present, operation.address, operation.input);
            session.dispatch({
              type: "cell.set",
              sheetId,
              address: operation.address,
              input: operation.input
            });
            if (!recordsEqual(next, model.present)) {
              model = {
                past: [...model.past, model.present].slice(-8),
                present: next,
                future: []
              };
            }
          } else if (operation.type === "undo") {
            session.dispatch({ type: "history.undo" });
            const previous = model.past.at(-1);
            if (previous) {
              model = {
                past: model.past.slice(0, -1),
                present: previous,
                future: [model.present, ...model.future]
              };
            }
          } else {
            session.dispatch({ type: "history.redo" });
            const next = model.future[0];
            if (next) {
              model = {
                past: [...model.past, model.present].slice(-8),
                present: next,
                future: model.future.slice(1)
              };
            }
          }

          const snapshot = session.getSnapshot();
          expect(readReferenceCells(snapshot.workbook, sheetId)).toEqual(model.present);
          expect(snapshot.canUndo).toBe(model.past.length > 0);
          expect(snapshot.canRedo).toBe(model.future.length > 0);
        }

        session.destroy();
      }
    );

    fc.assert(editUndoProperty, createPropertyOptions());
  });
});

type ReferenceCells = Record<string, number>;
type ReferenceHistory = {
  past: ReferenceCells[];
  present: ReferenceCells;
  future: ReferenceCells[];
};

function applyReferenceEdit(cells: ReferenceCells, address: string, input: string): ReferenceCells {
  const next = { ...cells };
  if (input === "") {
    delete next[address];
  } else {
    next[address] = Number(input);
  }
  return next;
}

function recordsEqual(left: ReferenceCells, right: ReferenceCells): boolean {
  return JSON.stringify(left) === JSON.stringify(right);
}

function readReferenceCells(workbook: WorkbookModel, sheetId: string): ReferenceCells {
  const result: ReferenceCells = {};
  for (const address of ["A1", "A2", "B1", "B2"]) {
    const value = getCellContent(workbook, sheetId, address);
    if (typeof value === "number") {
      result[address] = value;
    }
  }
  return result;
}

function createRawFormulaEngine(workbook: WorkbookModel): FormulaEngine {
  let projected = workbook;
  const read = (sheetId: string, address: string): CellContent => getCellContent(projected, sheetId, address);
  return {
    getDisplayValue(sheetId, address) {
      const value = read(sheetId, address);
      return value === null ? "" : String(value);
    },
    getComputedValue: read,
    getRawContent: read,
    update(nextWorkbook) {
      projected = nextWorkbook;
    },
    rebuild(nextWorkbook) {
      projected = nextWorkbook;
    },
    destroy() {}
  };
}
