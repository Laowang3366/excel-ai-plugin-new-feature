import { describe, expect, it } from "vitest";

import { shouldArchiveMemory } from "./memoryPruning";

describe("memory pruning", () => {
  it("archives active memories when expiresAt has passed", () => {
    expect(
      shouldArchiveMemory(
        { kind: "file_impression", status: "active", expiresAt: 100 } as any,
        101,
      ),
    ).toBe(true);
  });

  it("keeps active memories without expiresAt", () => {
    expect(
      shouldArchiveMemory(
        { kind: "constraint", status: "active" } as any,
        101,
      ),
    ).toBe(false);
  });
});
