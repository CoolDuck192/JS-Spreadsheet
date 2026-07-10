import {
  useCallback,
  useEffect,
  useMemo,
  useRef,
  useState,
  type RefObject,
  type UIEvent as ReactUIEvent
} from "react";
import {
  findVisibleRange,
  measureAxis,
  type AxisMeasurement
} from "../../core/viewport/axis";

const DEFAULT_VIEWPORT_HEIGHT = 560;
const DEFAULT_VIEWPORT_WIDTH = 960;

export type TwoAxisVirtualizerOptions = {
  scrollRef: RefObject<HTMLElement | null>;
  rowCount: number;
  columnCount: number;
  getRowKey(index: number): string;
  getColumnKey(index: number): string;
  getRowSize(index: number): number;
  getColumnSize(index: number): number;
  isRowHidden?(index: number): boolean;
  isColumnHidden?(index: number): boolean;
  rowOverscan?: number;
  columnOverscan?: number;
  rowViewportInset?: number;
  columnViewportInset?: number;
  scale?: number;
  resetKey?: string | number;
};

export type CellRect = { top: number; left: number; width: number; height: number };

export type TwoAxisVirtualizer = {
  rowMeasurements: readonly AxisMeasurement<string>[];
  columnMeasurements: readonly AxisMeasurement<string>[];
  visibleRows: readonly AxisMeasurement<string>[];
  visibleColumns: readonly AxisMeasurement<string>[];
  totalHeight: number;
  totalWidth: number;
  onScroll(event: ReactUIEvent<HTMLElement>): void;
  getCellRect(rowIndex: number, columnIndex: number): CellRect;
  ensureCellVisible(rowIndex: number, columnIndex: number): void;
};

type Viewport = {
  scrollTop: number;
  scrollLeft: number;
  height: number;
  width: number;
};

export function useTwoAxisVirtualizer({
  scrollRef,
  rowCount,
  columnCount,
  getRowKey,
  getColumnKey,
  getRowSize,
  getColumnSize,
  isRowHidden,
  isColumnHidden,
  rowOverscan = 4,
  columnOverscan = 2,
  rowViewportInset = 0,
  columnViewportInset = 0,
  scale = 1,
  resetKey
}: TwoAxisVirtualizerOptions): TwoAxisVirtualizer {
  const [viewport, setViewport] = useState<Viewport>({
    scrollTop: 0,
    scrollLeft: 0,
    height: DEFAULT_VIEWPORT_HEIGHT,
    width: DEFAULT_VIEWPORT_WIDTH
  });
  const scrollFrameRef = useRef<number | null>(null);
  const pendingScrollElementRef = useRef<HTMLElement | null>(null);
  const safeScale = Number.isFinite(scale) && scale > 0 ? scale : 1;
  const safeRowInset = Math.max(0, rowViewportInset);
  const safeColumnInset = Math.max(0, columnViewportInset);

  const rowMeasurements = useMemo(
    () => measureAxis(rowCount, getRowKey, getRowSize, isRowHidden),
    [getRowKey, getRowSize, isRowHidden, rowCount]
  );
  const columnMeasurements = useMemo(
    () => measureAxis(columnCount, getColumnKey, getColumnSize, isColumnHidden),
    [columnCount, getColumnKey, getColumnSize, isColumnHidden]
  );
  const rowsByIndex = useMemo(
    () => new Map(rowMeasurements.map((measurement) => [measurement.index, measurement])),
    [rowMeasurements]
  );
  const columnsByIndex = useMemo(
    () => new Map(columnMeasurements.map((measurement) => [measurement.index, measurement])),
    [columnMeasurements]
  );

  const rowViewportStart = viewport.scrollTop / safeScale;
  const rowViewportEnd = Math.max(
    rowViewportStart,
    (viewport.scrollTop + viewport.height) / safeScale - safeRowInset
  );
  const columnViewportStart = viewport.scrollLeft / safeScale;
  const columnViewportEnd = Math.max(
    columnViewportStart,
    (viewport.scrollLeft + viewport.width) / safeScale - safeColumnInset
  );
  const visibleRowRange = findVisibleRange(rowMeasurements, rowViewportStart, rowViewportEnd, rowOverscan);
  const visibleColumnRange = findVisibleRange(
    columnMeasurements,
    columnViewportStart,
    columnViewportEnd,
    columnOverscan
  );
  const visibleRows = useMemo(
    () => rowMeasurements.slice(visibleRowRange.first, visibleRowRange.last),
    [rowMeasurements, visibleRowRange.first, visibleRowRange.last]
  );
  const visibleColumns = useMemo(
    () => columnMeasurements.slice(visibleColumnRange.first, visibleColumnRange.last),
    [columnMeasurements, visibleColumnRange.first, visibleColumnRange.last]
  );
  const totalHeight = rowMeasurements.at(-1)?.end ?? 0;
  const totalWidth = columnMeasurements.at(-1)?.end ?? 0;

  const applyViewport = useCallback((element: HTMLElement) => {
    const nextViewport = readViewport(element);
    setViewport((current) =>
      current.scrollTop === nextViewport.scrollTop &&
      current.scrollLeft === nextViewport.scrollLeft &&
      current.height === nextViewport.height &&
      current.width === nextViewport.width
        ? current
        : nextViewport
    );
  }, []);

  const onScroll = useCallback(
    (event: ReactUIEvent<HTMLElement>) => {
      const element = event.currentTarget;
      pendingScrollElementRef.current = element;
      if (scrollFrameRef.current !== null) {
        return;
      }
      applyViewport(element);
      scrollFrameRef.current = requestAnimationFrame(() => {
        scrollFrameRef.current = null;
        const pendingElement = pendingScrollElementRef.current;
        if (pendingElement) {
          applyViewport(pendingElement);
        }
      });
    },
    [applyViewport]
  );

  useEffect(() => {
    const element = scrollRef.current;
    if (!element) {
      return undefined;
    }
    applyViewport(element);
    if (typeof ResizeObserver === "undefined") {
      return undefined;
    }
    const observer = new ResizeObserver(() => applyViewport(element));
    observer.observe(element);
    return () => observer.disconnect();
  }, [applyViewport, scrollRef]);

  useEffect(() => {
    const element = scrollRef.current;
    if (!element) {
      return;
    }
    if (scrollFrameRef.current !== null) {
      cancelAnimationFrame(scrollFrameRef.current);
      scrollFrameRef.current = null;
    }
    element.scrollTop = 0;
    element.scrollLeft = 0;
    applyViewport(element);
  }, [applyViewport, resetKey, scrollRef]);

  useEffect(
    () => () => {
      if (scrollFrameRef.current !== null) {
        cancelAnimationFrame(scrollFrameRef.current);
      }
    },
    []
  );

  const getCellRect = useCallback(
    (rowIndex: number, columnIndex: number): CellRect => {
      const row = rowsByIndex.get(rowIndex);
      const column = columnsByIndex.get(columnIndex);
      if (!row || !column) {
        throw new RangeError(`Cell ${rowIndex}:${columnIndex} is outside the measured axes`);
      }
      return {
        top: row.start,
        left: column.start,
        width: column.size,
        height: row.size
      };
    },
    [columnsByIndex, rowsByIndex]
  );

  const ensureCellVisible = useCallback(
    (rowIndex: number, columnIndex: number) => {
      const element = scrollRef.current;
      if (!element) {
        return;
      }
      const row = rowsByIndex.get(rowIndex);
      const column = columnsByIndex.get(columnIndex);
      if (row) {
        const height = element.clientHeight || DEFAULT_VIEWPORT_HEIGHT;
        const viewTop = element.scrollTop / safeScale;
        const viewBottom = (element.scrollTop + height) / safeScale - safeRowInset;
        if (row.start < viewTop) {
          element.scrollTop = Math.max(0, row.start * safeScale);
        } else if (row.end > viewBottom) {
          element.scrollTop = Math.max(0, (row.end + safeRowInset) * safeScale - height);
        }
      }
      if (column) {
        const width = element.clientWidth || DEFAULT_VIEWPORT_WIDTH;
        const viewLeft = element.scrollLeft / safeScale;
        const viewRight = (element.scrollLeft + width) / safeScale - safeColumnInset;
        if (column.start < viewLeft) {
          element.scrollLeft = Math.max(0, column.start * safeScale);
        } else if (column.end > viewRight) {
          element.scrollLeft = Math.max(0, (column.end + safeColumnInset) * safeScale - width);
        }
      }
    },
    [columnsByIndex, rowsByIndex, safeColumnInset, safeRowInset, safeScale, scrollRef]
  );

  return {
    rowMeasurements,
    columnMeasurements,
    visibleRows,
    visibleColumns,
    totalHeight,
    totalWidth,
    onScroll,
    getCellRect,
    ensureCellVisible
  };
}

function readViewport(element: HTMLElement): Viewport {
  return {
    scrollTop: element.scrollTop,
    scrollLeft: element.scrollLeft,
    height: element.clientHeight || DEFAULT_VIEWPORT_HEIGHT,
    width: element.clientWidth || DEFAULT_VIEWPORT_WIDTH
  };
}
