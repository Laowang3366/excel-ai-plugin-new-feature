import * as path from "node:path";

import { describe, expect, it } from "vitest";

import { isAllowedHotPatchPath, isPathInside } from "./hotPatchArchive";

describe("hot patch archive paths", () => {
  it("allows renderer and selected public resources only", () => {
    expect(isAllowedHotPatchPath("dist/index.html")).toBe(true);
    expect(isAllowedHotPatchPath("public/knowledge/builtin-knowledge.json")).toBe(true);
    expect(isAllowedHotPatchPath("electron/main.js")).toBe(false);
    expect(isAllowedHotPatchPath("../outside.txt")).toBe(false);
  });

  it("rejects the root itself and sibling paths with a shared prefix", () => {
    const root = path.resolve("C:/hot-patches/patch-001");

    expect(isPathInside(root, path.join(root, "dist/index.html"))).toBe(true);
    expect(isPathInside(root, root)).toBe(false);
    expect(isPathInside(root, path.resolve("C:/hot-patches/patch-001-escape/file.js"))).toBe(false);
  });
});
