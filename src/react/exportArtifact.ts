import type { ExportArtifact, ExportOptions, TableSession } from "../table/core/types";

export function exportArtifactToBlob(artifact: ExportArtifact): Blob {
  if (
    !(artifact.bytes instanceof Uint8Array)
    || artifact.mediaType.trim().length === 0
    || artifact.fileName.trim().length === 0
  ) {
    throw new TypeError("Invalid table export artifact");
  }
  const bytes = Uint8Array.from(artifact.bytes);
  const buffer = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;
  return new Blob([buffer], { type: artifact.mediaType });
}

export async function downloadTableExport(
  session: Pick<TableSession<unknown>, "export">,
  options: ExportOptions
): Promise<void> {
  const artifact = await session.export(options);
  const blob = exportArtifactToBlob(artifact);
  const url = URL.createObjectURL(blob);
  try {
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = artifact.fileName;
    anchor.click();
  } finally {
    URL.revokeObjectURL(url);
  }
}
