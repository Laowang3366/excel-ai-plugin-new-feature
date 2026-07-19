import { describe, expect, it } from "vitest";
import { normalizeBasePath, resolveViteBase } from "../scripts/basePath.mjs";

describe("normalizeBasePath", () => {
  it("normalizes root and subpaths", () => {
    expect(normalizeBasePath("/")).toBe("/");
    expect(normalizeBasePath("/excel-addin/")).toBe("/excel-addin/");
    expect(normalizeBasePath("excel-addin")).toBe("/excel-addin/");
    expect(normalizeBasePath("excel-addin/")).toBe("/excel-addin/");
    expect(normalizeBasePath("/excel-addin")).toBe("/excel-addin/");
  });

  it("rejects illegal values", () => {
    expect(() => normalizeBasePath("https://example.com/")).toThrow(/Invalid VITE_BASE/);
    expect(() => normalizeBasePath("../x")).toThrow(/Invalid VITE_BASE/);
    expect(() => normalizeBasePath("a?b")).toThrow(/Invalid VITE_BASE/);
    expect(() => normalizeBasePath("a#b")).toThrow(/Invalid VITE_BASE/);
    expect(() => normalizeBasePath("foo\\bar")).toThrow(/Invalid VITE_BASE/);
  });

  it("resolveViteBase defaults to /", () => {
    expect(resolveViteBase({})).toBe("/");
    expect(resolveViteBase({ VITE_BASE: "excel-addin" })).toBe("/excel-addin/");
  });
});
