import type { TableCellRef, TableSelection } from "../../table/core/types";
import type { GridEditorState } from "./types";

export type GridInteractionState = {
  selection: TableSelection;
  editing: GridEditorState | null;
  dragging: boolean;
};

export type GridInteractionAction =
  | { type: "select"; cell: TableCellRef; extend: boolean }
  | { type: "move"; rowDelta: number; columnDelta: number; extend: boolean }
  | { type: "pointer-down"; cell: TableCellRef; extend: boolean }
  | { type: "pointer-enter"; cell: TableCellRef }
  | { type: "pointer-up" }
  | { type: "start-edit"; rawText: string }
  | { type: "change-edit"; rawText: string }
  | { type: "commit-edit" }
  | { type: "cancel-edit" };

export type GridInteractionModel = {
  rowIds: readonly string[];
  columnIds: readonly string[];
};

export function createGridInteractionState(initialCell: TableCellRef): GridInteractionState {
  return {
    selection: { anchor: { ...initialCell }, focus: { ...initialCell } },
    editing: null,
    dragging: false
  };
}

export function reduceGridInteraction(
  state: GridInteractionState,
  action: GridInteractionAction,
  model: GridInteractionModel
): GridInteractionState {
  switch (action.type) {
    case "select":
      return {
        ...state,
        selection: selectCell(state.selection, action.cell, action.extend),
        dragging: false
      };
    case "move": {
      const focus = moveCell(state.selection.focus, action.rowDelta, action.columnDelta, model);
      return {
        ...state,
        selection: action.extend
          ? { anchor: state.selection.anchor, focus }
          : { anchor: focus, focus },
        dragging: false
      };
    }
    case "pointer-down":
      return {
        ...state,
        selection: selectCell(state.selection, action.cell, action.extend),
        dragging: true
      };
    case "pointer-enter":
      return state.dragging
        ? { ...state, selection: { anchor: state.selection.anchor, focus: { ...action.cell } } }
        : state;
    case "pointer-up":
      return state.dragging ? { ...state, dragging: false } : state;
    case "start-edit":
      return {
        ...state,
        editing: { ...state.selection.focus, rawText: action.rawText },
        dragging: false
      };
    case "change-edit":
      return state.editing ? { ...state, editing: { ...state.editing, rawText: action.rawText } } : state;
    case "commit-edit":
    case "cancel-edit":
      return state.editing ? { ...state, editing: null } : state;
    default:
      return state;
  }
}

function selectCell(selection: TableSelection, cell: TableCellRef, extend: boolean): TableSelection {
  const focus = { ...cell };
  return extend ? { anchor: selection.anchor, focus } : { anchor: focus, focus };
}

function moveCell(
  cell: TableCellRef,
  rowDelta: number,
  columnDelta: number,
  model: GridInteractionModel
): TableCellRef {
  if (model.rowIds.length === 0 || model.columnIds.length === 0) {
    return cell;
  }
  const rowIndex = clampIndex(indexOfOrFirst(model.rowIds, cell.rowId) + rowDelta, model.rowIds.length);
  const columnIndex = clampIndex(
    indexOfOrFirst(model.columnIds, cell.columnId) + columnDelta,
    model.columnIds.length
  );
  return { rowId: model.rowIds[rowIndex], columnId: model.columnIds[columnIndex] };
}

function indexOfOrFirst(ids: readonly string[], id: string): number {
  const index = ids.indexOf(id);
  return index < 0 ? 0 : index;
}

function clampIndex(index: number, length: number): number {
  return Math.max(0, Math.min(length - 1, index));
}
