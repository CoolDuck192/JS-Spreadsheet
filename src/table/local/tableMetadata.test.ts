import { describe, expect, it } from "vitest";
import { createTableMetadataKey, parseTableMetadataKey } from "./tableMetadata";

describe("table metadata keys", () => {
  it("round-trips ids containing separators without collisions", () => {
    const first = createTableMetadataKey("row::1", "column/2");
    const second = createTableMetadataKey("row", "1::column/2");
    expect(first).not.toBe(second);
    expect(parseTableMetadataKey(first)).toEqual({ rowId: "row::1", columnId: "column/2" });
    expect(parseTableMetadataKey(second)).toEqual({ rowId: "row", columnId: "1::column/2" });
  });
});
