import { describe, expect, it } from "vitest";
import type { ToolExecutor } from "../../shared/types";
import { resolveExecutableToolName } from "./toolNameResolution";

function executorsFor(names: string[]): Map<string, ToolExecutor> {
  return new Map(names.map((name) => [
    name,
    {
      name,
      execute: async () => ({ success: true }),
    },
  ]));
}

describe("resolveExecutableToolName", () => {
  it("prefers the original executor name when it exists", () => {
    expect(resolveExecutableToolName(
      "range.read",
      executorsFor(["range.read", "range_read"])
    )).toBe("range.read");
  });

  it("resolves underscored aliases to canonical dotted executors", () => {
    expect(resolveExecutableToolName(
      "ocr_parseDocument",
      executorsFor(["ocr.parseDocument"])
    )).toBe("ocr.parseDocument");
  });

  it("returns null when no executor matches any candidate", () => {
    expect(resolveExecutableToolName(
      "missing.tool",
      executorsFor(["range.read"])
    )).toBeNull();
  });
});
