import type { TableCellIssue } from "./types";

export type TableExtensionKind =
  | "accessor"
  | "calculate"
  | "format"
  | "parse"
  | "validate"
  | "update"
  | "permission"
  | "formula"
  | "host-callback"
  | "subscriber";

export function safeInvokeTableExtension<T>(
  kind: TableExtensionKind,
  invoke: () => T
): { ok: true; value: T } | { ok: false; issue: TableCellIssue } {
  try {
    return { ok: true, value: invoke() };
  } catch {
    return {
      ok: false,
      issue: {
        code: "TABLE_EXTENSION_ERROR",
        message: `Host ${kind} extension failed`
      }
    };
  }
}
