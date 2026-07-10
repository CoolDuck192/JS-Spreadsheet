import { StrictMode, type ReactNode } from "react";
import { act, render, renderHook, screen } from "@testing-library/react";
import { describe, expect, it, vi } from "vitest";
import { createWorkbookSession, type WorkbookSession } from "../core/workbook/WorkbookSession";
import { createBlankWorkbook, getCellContent } from "../lib/workbook";
import {
  WorkbookSessionProvider,
  useWorkbookSessionContext
} from "./WorkbookSessionContext";

function SnapshotCell() {
  const { session, snapshot } = useWorkbookSessionContext();
  return (
    <output data-testid={`snapshot-${snapshot.workbook.activeSheetId}`}>
      {String(getCellContent(snapshot.workbook, snapshot.workbook.activeSheetId, "A1") ?? "")}
      <button
        type="button"
        onClick={() => session.dispatch({
          type: "cell.set",
          sheetId: snapshot.workbook.activeSheetId,
          address: "A1",
          input: "context"
        })}
      >
        update
      </button>
    </output>
  );
}

describe("WorkbookSessionContext", () => {
  it("throws a stable error outside a provider", () => {
    expect(() => renderHook(() => useWorkbookSessionContext())).toThrow(
      "useWorkbookSessionContext must be used inside WorkbookSessionProvider"
    );
  });

  it("isolates nested providers and subscribes to live snapshots", () => {
    const outer = createWorkbookSession({ workbook: createBlankWorkbook() });
    const innerWorkbook = createBlankWorkbook();
    innerWorkbook.sheets[0] = { ...innerWorkbook.sheets[0], id: "inner-sheet" };
    innerWorkbook.activeSheetId = "inner-sheet";
    const inner = createWorkbookSession({ workbook: innerWorkbook });

    render(
      <WorkbookSessionProvider session={outer}>
        <SnapshotCell />
        <WorkbookSessionProvider session={inner}>
          <SnapshotCell />
        </WorkbookSessionProvider>
      </WorkbookSessionProvider>
    );

    const outputs = screen.getAllByText("update").map((button) => button.parentElement!);
    act(() => outputs[0].querySelector("button")!.click());
    expect(outputs[0]).toHaveTextContent("context");
    expect(outputs[1]).not.toHaveTextContent("context");

    outer.destroy();
    inner.destroy();
  });

  it("never destroys a supplied session on unmount or StrictMode replay", async () => {
    const session = createWorkbookSession({ workbook: createBlankWorkbook() });
    const destroy = vi.spyOn(session, "destroy");
    const wrapper = ({ children }: { children: ReactNode }) => (
      <StrictMode>
        <WorkbookSessionProvider session={session}>{children}</WorkbookSessionProvider>
      </StrictMode>
    );

    const { unmount } = renderHook(() => useWorkbookSessionContext(), { wrapper });
    unmount();
    await act(async () => Promise.resolve());
    expect(destroy).not.toHaveBeenCalled();
    session.destroy();
  });
});
