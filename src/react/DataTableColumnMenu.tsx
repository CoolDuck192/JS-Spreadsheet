import { useEffect, useLayoutEffect, useMemo, useRef, useState } from "react";
import type { QueryScalar } from "../table/core/query";
import { safeInvokeTableExtension } from "../table/core/safeInvoke";
import type { TableDiagnosticEvent, TableSession, TableViewSnapshot } from "../table/core/types";
import { columnLabel } from "./DataTableCell";
import type { ColumnDef } from "./tableTypes";

export function DataTableColumnMenu<TRow>({
  anchor,
  column,
  session,
  snapshot,
  rows,
  onIssue,
  onDiagnostic,
  onAnnouncement,
  onClose
}: {
  anchor: HTMLButtonElement;
  column: ColumnDef<TRow>;
  session: TableSession<TRow, ColumnDef<TRow>>;
  snapshot: TableViewSnapshot<TRow, ColumnDef<TRow>>;
  rows: readonly TRow[];
  onIssue(message: string): void;
  onDiagnostic?(event: TableDiagnosticEvent): void;
  onAnnouncement(message: string): void;
  onClose(anchor: HTMLButtonElement): void;
}) {
  const label = columnLabel(column);
  const menuRef = useRef<HTMLDivElement>(null);
  const closeRef = useRef<(restoreFocus?: boolean) => void>(() => {});
  const [position, setPosition] = useState({ top: VIEWPORT_MARGIN, left: VIEWPORT_MARGIN });
  const [filterText, setFilterText] = useState("");
  const [width, setWidth] = useState(String(snapshot.state.columnWidths[column.id] ?? column.width ?? 120));
  const [aggregate, setAggregate] = useState<"sum" | "average" | "count" | "min" | "max">(
    column.aggregatable?.[0] ?? "count"
  );
  const orderedIds = completeOrder(snapshot);
  const index = orderedIds.indexOf(column.id);
  const sortState = snapshot.operationStates.sort;
  const filterState = snapshot.operationStates.filter;
  const groupState = snapshot.operationStates.group;
  const aggregateState = snapshot.operationStates.aggregate;
  const headerActionStates = useMemo(() => new Map((column.headerActions ?? []).map((action) => [
    action.id,
    safeInvokeTableExtension("host-callback", () => action.disabled === true
      || (typeof action.disabled === "function" && action.disabled(rows)))
  ])), [column.headerActions, rows]);

  useEffect(() => {
    setWidth(String(snapshot.state.columnWidths[column.id] ?? column.width ?? 120));
  }, [column.id, column.width, snapshot.state.columnWidths]);

  useEffect(() => {
    if ([...headerActionStates.values()].some((result) => !result.ok)) {
      reportExtension(onDiagnostic, column.id, "header-action");
    }
  }, [column.id, headerActionStates, onDiagnostic]);

  useLayoutEffect(() => {
    const menu = menuRef.current;
    const view = anchor.ownerDocument.defaultView;
    if (!menu || !view) return;
    const anchorHeader = anchor.closest('[role="columnheader"]');

    let closed = false;
    let shown = false;

    const close = (restoreFocus = true, hide = true) => {
      if (closed) return;
      closed = true;
      if (hide && shown) {
        shown = false;
        try {
          menu.hidePopover();
        } catch {
          // The browser may already have light-dismissed the popover.
        }
      }
      if (restoreFocus && anchor.isConnected) anchor.focus();
      onClose(anchor);
    };

    const place = () => {
      if (closed) return;
      if (!anchor.isConnected || !menu.isConnected) {
        close(false);
        return;
      }
      const anchorRect = anchor.getBoundingClientRect();
      const menuRect = menu.getBoundingClientRect();
      const maxLeft = Math.max(VIEWPORT_MARGIN, view.innerWidth - menuRect.width - VIEWPORT_MARGIN);
      const maxTop = Math.max(VIEWPORT_MARGIN, view.innerHeight - menuRect.height - VIEWPORT_MARGIN);
      const next = {
        left: clamp(anchorRect.right - menuRect.width, VIEWPORT_MARGIN, maxLeft),
        top: clamp(anchorRect.bottom + MENU_GAP, VIEWPORT_MARGIN, maxTop)
      };
      setPosition((current) => current.left === next.left && current.top === next.top ? current : next);
    };

    const handleToggle = (event: Event) => {
      const state = (event as ToggleEvent).newState;
      if (state === "open") {
        shown = true;
      } else if (state === "closed") {
        shown = false;
        close(true, false);
      }
    };
    const handleScroll = (event: Event) => {
      if (event.target instanceof Node && menu.contains(event.target)) return;
      if (event.target instanceof Node && anchorHeader?.contains(event.target)) {
        place();
        return;
      }
      close();
    };
    const handleResize = () => place();

    closeRef.current = (restoreFocus = true) => close(restoreFocus);
    menu.addEventListener("toggle", handleToggle);
    view.addEventListener("scroll", handleScroll, true);
    view.addEventListener("resize", handleResize);

    const observer = new MutationObserver(() => {
      if (!anchor.isConnected || !menu.isConnected) close(false);
    });
    observer.observe(anchor.ownerDocument, { childList: true, subtree: true });

    if (!anchor.isConnected) {
      close(false, false);
    } else {
      try {
        menu.showPopover();
        shown = true;
        place();
        firstEnabledControl(menu)?.focus();
      } catch {
        close(false, false);
      }
    }

    return () => {
      menu.removeEventListener("toggle", handleToggle);
      view.removeEventListener("scroll", handleScroll, true);
      view.removeEventListener("resize", handleResize);
      observer.disconnect();
      closeRef.current = () => {};
      close(false);
    };
  }, [anchor, onClose]);

  async function run(intent: Parameters<typeof session.dispatch>[0]) {
    const result = await session.dispatch(intent);
    if (result.status === "rejected") onIssue(result.issues?.[0]?.message ?? "Table operation was rejected");
    return result;
  }

  async function runHeaderAction(action: NonNullable<ColumnDef<TRow>["headerActions"]>[number]) {
    const disabled = headerActionStates.get(action.id);
    if (!disabled?.ok) return;
    if (disabled.value) return;
    const invoked = safeInvokeTableExtension("host-callback", () => action.run({ columnId: column.id, rows }));
    if (!invoked.ok) {
      reportExtension(onDiagnostic, column.id, "header-action");
      return;
    }
    try {
      await invoked.value;
    } catch {
      reportExtension(onDiagnostic, column.id, "header-action");
    }
  }

  function applyFilter() {
    const value = filterText.trim();
    if (!value) {
      void run({ type: "set-filter", filter: null });
      return;
    }
    void run({
      type: "set-filter",
      filter: {
        kind: "comparison",
        columnId: column.id,
        operator: column.dataType === "text" || column.dataType === undefined ? "contains" : "eq",
        value: scalarFor(column, value)
      }
    });
  }

  async function move(delta: number) {
    if (index < 0) return;
    const target = Math.max(0, Math.min(orderedIds.length - 1, index + delta));
    if (target === index) return;
    const next = [...orderedIds];
    next.splice(index, 1);
    next.splice(target, 0, column.id);
    const result = await run({ type: "set-column-order", columnIds: next });
    if (result.status === "committed") onAnnouncement(`${label} moved to position ${target + 1}`);
  }

  return (
    <div
      ref={menuRef}
      popover="auto"
      className="js-spreadsheet-data-table__column-menu"
      role="menu"
      aria-label={`${label} column menu`}
      style={{ position: "fixed", top: position.top, left: position.left }}
      onKeyDown={(event) => {
        if (event.key !== "Escape") return;
        event.preventDefault();
        event.stopPropagation();
        closeRef.current();
      }}
    >
      <button
        type="button"
        aria-label={`Sort ${label} ascending`}
        disabled={!column.sortable || !sortState.enabled}
        title={sortState.reason}
        onClick={() => void run({ type: "set-sorting", sorting: [{ columnId: column.id, direction: "asc" }] })}
      >Sort ascending</button>
      <button
        type="button"
        aria-label={`Sort ${label} descending`}
        disabled={!column.sortable || !sortState.enabled}
        title={sortState.reason}
        onClick={() => void run({ type: "set-sorting", sorting: [{ columnId: column.id, direction: "desc" }] })}
      >Sort descending</button>
      <button type="button" aria-label={`Clear ${label} sort`} disabled={!sortState.enabled} onClick={() => void run({ type: "set-sorting", sorting: [] })}>
        Clear sort
      </button>
      {column.filterable ? (
        <div>
          <label>
            Filter {label}
            {column.dataType === "boolean" ? (
              <select aria-label={`Filter ${label}`} value={filterText} onChange={(event) => setFilterText(event.currentTarget.value)}>
                <option value="">Any</option>
                <option value="true">True</option>
                <option value="false">False</option>
              </select>
            ) : (
              <input
                aria-label={`Filter ${label}`}
                type={column.dataType === "number" ? "number" : column.dataType === "date" ? "date" : column.dataType === "datetime" ? "datetime-local" : "text"}
                value={filterText}
                onChange={(event) => setFilterText(event.currentTarget.value)}
              />
            )}
          </label>
          <button type="button" aria-label={`Apply ${label} filter`} disabled={!filterState.enabled} title={filterState.reason} onClick={applyFilter}>Apply filter</button>
        </div>
      ) : null}
      {column.groupable ? (
        <button
          type="button"
          aria-label={snapshot.state.grouping.some((group) => group.columnId === column.id)
            ? `Ungroup ${label}`
            : `Group by ${label}`}
          disabled={!groupState.enabled}
          title={groupState.reason}
          onClick={() => void run({
            type: "set-grouping",
            grouping: snapshot.state.grouping.some((group) => group.columnId === column.id)
              ? snapshot.state.grouping.filter((group) => group.columnId !== column.id)
              : [...snapshot.state.grouping, { columnId: column.id }]
          })}
        >{snapshot.state.grouping.some((group) => group.columnId === column.id) ? "Ungroup" : "Group"}</button>
      ) : null}
      {column.aggregatable?.length ? (
        <div>
          <label>
            Aggregate {label}
            <select aria-label={`Aggregate ${label}`} value={aggregate} onChange={(event) => setAggregate(event.currentTarget.value as typeof aggregate)}>
              {column.aggregatable.map((value) => <option key={value} value={value}>{value}</option>)}
            </select>
          </label>
          <button
            type="button"
            aria-label={`Apply ${label} aggregate`}
            disabled={!aggregateState.enabled}
            title={aggregateState.reason}
            onClick={() => void run({
              type: "set-aggregates",
              aggregates: [{ id: column.id, columnId: column.id, function: aggregate }]
            })}
          >Apply aggregate</button>
        </div>
      ) : null}
      <label>
        Column width
        <input aria-label={`${label} column width`} type="number" min={column.minWidth ?? 40} max={column.maxWidth ?? 1200} value={width} onChange={(event) => setWidth(event.currentTarget.value)} />
      </label>
      <button type="button" aria-label={`Resize ${label}`} onClick={() => void run({ type: "resize-column", columnId: column.id, width: Number(width) })}>
        Resize
      </button>
      <button type="button" aria-label={`Hide ${label}`} onClick={() => void run({ type: "set-column-visibility", columnId: column.id, visible: false })}>Hide</button>
      <button type="button" aria-label={`Pin ${label} left`} onClick={() => void run({ type: "set-column-pinning", columnId: column.id, pin: "left" })}>Pin left</button>
      <button type="button" aria-label={`Pin ${label} right`} onClick={() => void run({ type: "set-column-pinning", columnId: column.id, pin: "right" })}>Pin right</button>
      <button type="button" aria-label={`Unpin ${label}`} onClick={() => void run({ type: "set-column-pinning", columnId: column.id, pin: false })}>Unpin</button>
      <button type="button" aria-label={`Move ${label} left`} disabled={index <= 0} onClick={() => void move(-1)}>Move left</button>
      <button type="button" aria-label={`Move ${label} right`} disabled={index < 0 || index >= orderedIds.length - 1} onClick={() => void move(1)}>Move right</button>
      {column.headerActions?.map((action) => {
        const actionState = headerActionStates.get(action.id);
        return (
          <button key={action.id} type="button" aria-label={action.label} disabled={!actionState?.ok || actionState.value} onClick={() => void runHeaderAction(action)}>
            {action.label}
          </button>
        );
      })}
      {[sortState, filterState, groupState, aggregateState]
        .filter((state) => !state.enabled && state.reason)
        .map((state, reasonIndex) => <span key={`${reasonIndex}:${state.reason}`}>{state.reason}</span>)}
    </div>
  );
}

const VIEWPORT_MARGIN = 8;
const MENU_GAP = 4;

function clamp(value: number, minimum: number, maximum: number): number {
  return Math.min(maximum, Math.max(minimum, value));
}

function firstEnabledControl(container: HTMLElement): HTMLElement | null {
  return container.querySelector<HTMLElement>([
    "button:not(:disabled)",
    "input:not(:disabled)",
    "select:not(:disabled)",
    "textarea:not(:disabled)",
    "[href]",
    '[tabindex]:not([tabindex="-1"])'
  ].join(","));
}

function completeOrder<TRow>(snapshot: TableViewSnapshot<TRow, ColumnDef<TRow>>): string[] {
  const ids = snapshot.columns.map((column) => column.id);
  return [...snapshot.state.columnOrder.filter((id) => ids.includes(id)), ...ids.filter((id) => !snapshot.state.columnOrder.includes(id))];
}

function scalarFor<TRow>(column: ColumnDef<TRow>, value: string): QueryScalar {
  if (column.dataType === "number") return { type: "number", value: Number(value) };
  if (column.dataType === "boolean") return { type: "boolean", value: value === "true" };
  if (column.dataType === "date") return { type: "date", value };
  if (column.dataType === "datetime") return { type: "datetime", value };
  return { type: "string", value };
}

function reportExtension(
  onDiagnostic: ((event: TableDiagnosticEvent) => void) | undefined,
  columnId: string,
  slot: "header-action"
) {
  try {
    onDiagnostic?.({ category: "extension", metadata: { columnId, slot } });
  } catch {
    // Host diagnostics are isolated from the table.
  }
}
