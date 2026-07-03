import { describe, expect, it } from "vitest";

import {
  buildToolProfileKey,
  normalizeToolProfileOperation,
  shouldPromoteToolProfile,
  updateToolProfileStats,
} from "./toolSuccessProfile";

describe("tool success profile", () => {
  it("requires multiple samples before promotion", () => {
    expect(shouldPromoteToolProfile({ successCount: 1, failureCount: 0 })).toBe(false);
    expect(shouldPromoteToolProfile({ successCount: 3, failureCount: 1 })).toBe(true);
  });

  it("updates counts without storing user prompt text", () => {
    const updated = updateToolProfileStats(undefined, {
      app: "powerpoint",
      operation: "create",
      toolFamily: "openxml",
      success: true,
    });
    expect(updated.successCount).toBe(1);
    expect(updated.failureCount).toBe(0);
    expect(JSON.stringify(updated)).not.toContain("用户");
  });

  it("increments failure count for failed tool events", () => {
    const updated = updateToolProfileStats(undefined, {
      app: "excel",
      operation: "validate",
      toolFamily: "python",
      success: false,
    });
    expect(updated.successCount).toBe(0);
    expect(updated.failureCount).toBe(1);
  });

  it("normalizes known operation casing and whitespace", () => {
    expect(normalizeToolProfileOperation(" Format ")).toBe("format");
  });

  it("does not preserve arbitrary operation text", () => {
    expect(normalizeToolProfileOperation("客户名单.xlsx")).toBe("unknown");
  });

  it("builds stable app operation family key", () => {
    expect(buildToolProfileKey({
      app: "word",
      operation: "format",
      toolFamily: "com",
    })).toBe("word:format:com");
  });

  it("builds keys without storing raw operation text", () => {
    const key = buildToolProfileKey({
      app: "excel",
      operation: "客户名单.xlsx" as any,
      toolFamily: "openxml",
    });
    expect(key).toBe("excel:unknown:openxml");
    expect(key).not.toContain("客户名单");
  });

  it("updates stats without storing raw operation text", () => {
    const updated = updateToolProfileStats(undefined, {
      app: "powerpoint",
      operation: "请把这份文件美化" as any,
      toolFamily: "office_action",
      success: true,
    });
    expect(updated.operation).toBe("unknown");
    expect(JSON.stringify(updated)).not.toContain("请把这份文件美化");
  });
});
