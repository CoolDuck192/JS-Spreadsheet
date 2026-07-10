import { createElement, createRef } from "react";
import { act, render, screen, waitFor } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import { afterEach, describe, expect, it, vi } from "vitest";
import { createLocalRecordTableSession } from "../table/local/RecordTableSession";
import type { ExportArtifact, ExportOptions } from "../table/core/types";
import { DataTable } from "./DataTable";
import { exportArtifactToBlob } from "./exportArtifact";
import type { ColumnDef, DataTableHandle } from "./tableTypes";

type Row = { id: string; name: string };
const rows: readonly Row[] = [{ id: "1", name: "Ada" }];
const columns: readonly ColumnDef<Row>[] = [{
  id: "name",
  header: "Name",
  accessor: (row) => row.name,
  update: (row, value) => ({ ...row, name: String(value) })
}];
const csvArtifact: ExportArtifact = {
  bytes: new Uint8Array([78, 97, 109, 101, 13, 10, 65, 100, 97, 13, 10]),
  mediaType: "text/csv;charset=utf-8",
  fileName: "employees.csv"
};
const xlsxArtifact: ExportArtifact = {
  bytes: new Uint8Array([80, 75, 3, 4]),
  mediaType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  fileName: "employees.xlsx"
};

afterEach(() => {
  vi.restoreAllMocks();
});

describe("browser export artifacts", () => {
  it("copies artifact bytes and media type into a browser Blob", async () => {
    const blob = exportArtifactToBlob(csvArtifact);

    expect(blob.type).toBe(csvArtifact.mediaType);
    expect([...new Uint8Array(await blob.arrayBuffer())]).toEqual([...csvArtifact.bytes]);
  });

  it("does not let later artifact byte mutation alter the returned Blob", async () => {
    const artifact = { ...xlsxArtifact, bytes: new Uint8Array(xlsxArtifact.bytes) };
    const blob = exportArtifactToBlob(artifact);

    artifact.bytes[0] = 0;

    expect([...new Uint8Array(await blob.arrayBuffer())]).toEqual([80, 75, 3, 4]);
  });

  it("delegates exact export options once through the DataTable handle", async () => {
    const session = createSession();
    const exportSession = vi.spyOn(session, "export").mockResolvedValue(csvArtifact);
    const ref = createRef<DataTableHandle>();
    const options: ExportOptions = {
      format: "csv",
      scope: "completeDataset",
      includeHeaders: false,
      fileName: "employees.csv"
    };
    render(createElement(DataTable<Row>, { "aria-label": "Employees", session, ref }));

    let blob!: Blob;
    await act(async () => {
      blob = await ref.current!.export(options);
    });

    expect(exportSession).toHaveBeenCalledTimes(1);
    expect(exportSession).toHaveBeenCalledWith(options);
    expect([...new Uint8Array(await blob.arrayBuffer())]).toEqual([...csvArtifact.bytes]);
    session.destroy();
  });

  it("uses the artifact file name for toolbar download and revokes its object URL", async () => {
    const user = userEvent.setup();
    const session = createSession();
    vi.spyOn(session, "export").mockResolvedValue(csvArtifact);
    const createObjectURL = vi.spyOn(URL, "createObjectURL").mockReturnValue("blob:employees");
    const revokeObjectURL = vi.spyOn(URL, "revokeObjectURL").mockImplementation(() => {});
    const anchorClick = vi.spyOn(HTMLAnchorElement.prototype, "click").mockImplementation(() => {});
    render(createElement(DataTable<Row>, { "aria-label": "Employees", session }));

    await user.click(screen.getByRole("button", { name: "Export CSV" }));

    await waitFor(() => expect(anchorClick).toHaveBeenCalledTimes(1));
    expect(createObjectURL).toHaveBeenCalledTimes(1);
    expect(revokeObjectURL).toHaveBeenCalledWith("blob:employees");
    const clickedAnchor = anchorClick.mock.instances[0] as HTMLAnchorElement;
    expect(clickedAnchor.download).toBe("employees.csv");
    session.destroy();
  });
});

function createSession() {
  return createLocalRecordTableSession<Row, ColumnDef<Row>>({
    source: { kind: "local", rows, getRowId: (row) => row.id },
    columns
  });
}
