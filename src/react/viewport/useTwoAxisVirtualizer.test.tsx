import { act, fireEvent, render, screen } from "@testing-library/react";
import { useRef } from "react";
import { afterEach, describe, expect, it, vi } from "vitest";
import { useTwoAxisVirtualizer } from "./useTwoAxisVirtualizer";

function LargeHarness() {
  const scrollRef = useRef<HTMLDivElement>(null);
  const virtualizer = useTwoAxisVirtualizer({
    scrollRef,
    rowCount: 100_000,
    columnCount: 10_000,
    getRowKey: (index) => `row-${index}`,
    getColumnKey: (index) => `column-${index}`,
    getRowSize: (index) => (index === 0 ? 40 : 28),
    getColumnSize: (index) => (index === 0 ? 160 : 96),
    rowOverscan: 4,
    columnOverscan: 2
  });
  return (
    <div>
      <div ref={scrollRef} data-testid="scroll" onScroll={virtualizer.onScroll} />
      <output data-testid="window">
        {JSON.stringify({
          rows: virtualizer.visibleRows.map((item) => item.index),
          columns: virtualizer.visibleColumns.map((item) => item.index),
          totalHeight: virtualizer.totalHeight,
          totalWidth: virtualizer.totalWidth
        })}
      </output>
    </div>
  );
}

function GeometryHarness() {
  const scrollRef = useRef<HTMLDivElement>(null);
  const virtualizer = useTwoAxisVirtualizer({
    scrollRef,
    rowCount: 4,
    columnCount: 4,
    getRowKey: (index) => `row-${index}`,
    getColumnKey: (index) => `column-${index}`,
    getRowSize: (index) => [40, 50, 60, 70][index],
    getColumnSize: (index) => [100, 120, 140, 160][index],
    rowOverscan: 0,
    columnOverscan: 0
  });
  return (
    <div>
      <div ref={scrollRef} data-testid="geometry-scroll" onScroll={virtualizer.onScroll} />
      <output data-testid="cell-rect">{JSON.stringify(virtualizer.getCellRect(2, 2))}</output>
      <button type="button" onClick={() => virtualizer.ensureCellVisible(1, 1)}>
        Ensure visible cell
      </button>
      <button type="button" onClick={() => virtualizer.ensureCellVisible(3, 3)}>
        Ensure far cell
      </button>
      <button type="button" onClick={() => virtualizer.ensureCellVisible(0, 0)}>
        Ensure first cell
      </button>
    </div>
  );
}

function ResetHarness({ resetKey }: { resetKey: string }) {
  const scrollRef = useRef<HTMLDivElement>(null);
  const virtualizer = useTwoAxisVirtualizer({
    scrollRef,
    resetKey,
    rowCount: 1_000,
    columnCount: 100,
    getRowKey: (index) => `row-${index}`,
    getColumnKey: (index) => `column-${index}`,
    getRowSize: () => 28,
    getColumnSize: () => 96,
    rowOverscan: 0,
    columnOverscan: 0
  });
  return (
    <div ref={scrollRef} data-testid="reset-scroll" onScroll={virtualizer.onScroll}>
      <output data-testid="reset-window">
        {JSON.stringify({ row: virtualizer.visibleRows[0]?.index, column: virtualizer.visibleColumns[0]?.index })}
      </output>
    </div>
  );
}

afterEach(() => {
  vi.restoreAllMocks();
  vi.unstubAllGlobals();
});

describe("useTwoAxisVirtualizer", () => {
  it("bounds both axes and scrolls to a distant cell", () => {
    render(<LargeHarness />);
    const scroll = screen.getByTestId("scroll");
    Object.defineProperties(scroll, {
      clientHeight: { configurable: true, value: 280 },
      clientWidth: { configurable: true, value: 480 },
      scrollTop: { configurable: true, writable: true, value: 28_000 },
      scrollLeft: { configurable: true, writable: true, value: 9_600 }
    });

    fireEvent.scroll(scroll);

    const windowState = JSON.parse(screen.getByTestId("window").textContent ?? "{}");
    expect(windowState.rows.length).toBeLessThan(24);
    expect(windowState.columns.length).toBeLessThan(12);
    expect(windowState.rows[0]).toBeGreaterThan(900);
    expect(windowState.columns[0]).toBeGreaterThan(90);
    expect(windowState.totalHeight).toBeGreaterThan(2_000_000);
    expect(windowState.totalWidth).toBeGreaterThan(900_000);
  });

  it("returns exact variable-size rectangles and scrolls only axes outside the viewport", () => {
    render(<GeometryHarness />);
    const scroll = screen.getByTestId("geometry-scroll");
    Object.defineProperties(scroll, {
      clientHeight: { configurable: true, value: 100 },
      clientWidth: { configurable: true, value: 200 },
      scrollTop: { configurable: true, writable: true, value: 40 },
      scrollLeft: { configurable: true, writable: true, value: 100 }
    });

    expect(JSON.parse(screen.getByTestId("cell-rect").textContent ?? "{}")).toEqual({
      top: 90,
      left: 220,
      width: 140,
      height: 60
    });

    fireEvent.click(screen.getByRole("button", { name: "Ensure visible cell" }));
    expect({ top: scroll.scrollTop, left: scroll.scrollLeft }).toEqual({ top: 40, left: 100 });

    fireEvent.click(screen.getByRole("button", { name: "Ensure far cell" }));
    expect({ top: scroll.scrollTop, left: scroll.scrollLeft }).toEqual({ top: 120, left: 320 });

    fireEvent.click(screen.getByRole("button", { name: "Ensure first cell" }));
    expect({ top: scroll.scrollTop, left: scroll.scrollLeft }).toEqual({ top: 0, left: 0 });
  });

  it("applies the leading scroll immediately and coalesces a trailing update per animation frame", () => {
    const frames: FrameRequestCallback[] = [];
    vi.spyOn(window, "requestAnimationFrame").mockImplementation((callback) => {
      frames.push(callback);
      return frames.length;
    });
    vi.spyOn(window, "cancelAnimationFrame").mockImplementation(() => undefined);
    render(<LargeHarness />);
    const scroll = screen.getByTestId("scroll");
    Object.defineProperties(scroll, {
      clientHeight: { configurable: true, value: 280 },
      clientWidth: { configurable: true, value: 480 },
      scrollTop: { configurable: true, writable: true, value: 2_800 },
      scrollLeft: { configurable: true, writable: true, value: 9_600 }
    });

    fireEvent.scroll(scroll);
    const leading = JSON.parse(screen.getByTestId("window").textContent ?? "{}");
    expect(leading.rows[0]).toBeGreaterThan(80);
    expect(leading.columns[0]).toBeGreaterThan(90);

    scroll.scrollTop = 5_600;
    scroll.scrollLeft = 19_200;
    fireEvent.scroll(scroll);
    expect(JSON.parse(screen.getByTestId("window").textContent ?? "{}")).toEqual(leading);

    act(() => frames.shift()?.(0));
    const trailing = JSON.parse(screen.getByTestId("window").textContent ?? "{}");
    expect(trailing.rows[0]).toBeGreaterThan(180);
    expect(trailing.columns[0]).toBeGreaterThan(190);
  });

  it("updates both viewport dimensions from ResizeObserver", () => {
    let resize: ResizeObserverCallback | null = null;
    vi.stubGlobal(
      "ResizeObserver",
      class ResizeObserverMock {
        constructor(callback: ResizeObserverCallback) {
          resize = callback;
        }

        observe() {}
        unobserve() {}
        disconnect() {}
      }
    );
    render(<LargeHarness />);
    const scroll = screen.getByTestId("scroll");
    Object.defineProperties(scroll, {
      clientHeight: { configurable: true, value: 140 },
      clientWidth: { configurable: true, value: 240 }
    });

    act(() => resize?.([], {} as ResizeObserver));

    const windowState = JSON.parse(screen.getByTestId("window").textContent ?? "{}");
    expect(windowState.rows.length).toBeLessThan(16);
    expect(windowState.columns.length).toBeLessThan(9);
  });

  it("resets the DOM and measured viewport when the dataset key changes", () => {
    const { rerender } = render(<ResetHarness resetKey="first" />);
    const scroll = screen.getByTestId("reset-scroll");
    Object.defineProperties(scroll, {
      clientHeight: { configurable: true, value: 280 },
      clientWidth: { configurable: true, value: 480 },
      scrollTop: { configurable: true, writable: true, value: 2_800 },
      scrollLeft: { configurable: true, writable: true, value: 1_920 }
    });
    fireEvent.scroll(scroll);
    expect(JSON.parse(screen.getByTestId("reset-window").textContent ?? "{}")).toEqual({ row: 99, column: 19 });

    rerender(<ResetHarness resetKey="second" />);

    expect({ top: scroll.scrollTop, left: scroll.scrollLeft }).toEqual({ top: 0, left: 0 });
    expect(JSON.parse(screen.getByTestId("reset-window").textContent ?? "{}")).toEqual({ row: 0, column: 0 });
  });
});
