import { describe, expect, it } from "vitest";
import { createCommandIdFactory } from "./commandId";

describe("createCommandIdFactory", () => {
  it("combines a collision-resistant session id with a monotonic local sequence", () => {
    const next = createCommandIdFactory(() => "session-uuid");
    expect(next()).toBe("session-uuid:1");
    expect(next()).toBe("session-uuid:2");
  });

  it("keeps independent sessions distinct", () => {
    const first = createCommandIdFactory(() => "first-uuid");
    const second = createCommandIdFactory(() => "second-uuid");
    expect(first()).not.toBe(second());
  });
});
