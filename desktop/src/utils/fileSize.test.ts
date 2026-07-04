import { describe, expect, test } from "vitest";
import { formatFileSize } from "./fileSize";

describe("formatFileSize", () => {
  test("formats exact file sizes with one decimal for legacy folder lists", () => {
    expect(formatFileSize(512)).toBe("512 B");
    expect(formatFileSize(1536)).toBe("1.5 KB");
    expect(formatFileSize(120 * 1024)).toBe("120.0 KB");
    expect(formatFileSize(2.5 * 1024 * 1024)).toBe("2.5 MB");
  });

  test("formats compact attachment sizes and hides invalid values", () => {
    expect(formatFileSize(undefined, { emptyText: "", compact: true })).toBe("");
    expect(formatFileSize(Number.NaN, { emptyText: "", compact: true })).toBe("");
    expect(formatFileSize(-1, { emptyText: "", compact: true })).toBe("");
    expect(formatFileSize(120 * 1024, { emptyText: "", compact: true })).toBe("120 KB");
    expect(formatFileSize(12 * 1024 * 1024, { emptyText: "", compact: true })).toBe("12 MB");
  });
});
