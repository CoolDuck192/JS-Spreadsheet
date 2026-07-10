import { describe, expect, it } from "vitest";
import { safeInvokeTableExtension, type TableExtensionKind } from "./safeInvoke";

const kinds: readonly TableExtensionKind[] = [
  "accessor",
  "calculate",
  "format",
  "parse",
  "validate",
  "update",
  "permission",
  "formula",
  "host-callback",
  "subscriber"
];

describe("safeInvokeTableExtension", () => {
  it("returns successful extension values unchanged", () => {
    expect(safeInvokeTableExtension("accessor", () => 42)).toEqual({ ok: true, value: 42 });
  });

  it.each(kinds)("sanitizes thrown %s extension failures", (kind) => {
    const result = safeInvokeTableExtension(kind, () => {
      throw new Error("secret row value and credential");
    });

    expect(result).toEqual({
      ok: false,
      issue: {
        code: "TABLE_EXTENSION_ERROR",
        message: `Host ${kind} extension failed`
      }
    });
    expect(JSON.stringify(result)).not.toContain("secret");
    expect(JSON.stringify(result)).not.toContain("credential");
  });
});
