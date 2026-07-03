import { describe, expect, it } from "vitest";

import { chooseConsolidationAction } from "./memoryConsolidation";

describe("memory consolidation", () => {
  it("ignores single-sample tool success profiles", () => {
    expect(
      chooseConsolidationAction({
        kind: "tool_success_profile",
        visibility: "internal",
        namespace: "global",
        content: "内部工具统计",
        metadata: { source: "telemetry", successCount: 1, failureCount: 0 },
      }),
    ).toBe("ignore");
  });

  it("ignores user-visible tool success profiles even with multiple samples", () => {
    expect(
      chooseConsolidationAction({
        kind: "tool_success_profile",
        visibility: "user",
        namespace: "global",
        content: "内部工具统计",
        metadata: { source: "telemetry", successCount: 3, failureCount: 0 },
      }),
    ).toBe("ignore");
  });

  it("ignores tool success profiles without telemetry source", () => {
    expect(
      chooseConsolidationAction({
        kind: "tool_success_profile",
        visibility: "internal",
        namespace: "global",
        content: "内部工具统计",
        metadata: { successCount: 3, failureCount: 0 },
      }),
    ).toBe("ignore");
  });

  it("ignores telemetry-like tool success profiles in ordinary consolidation", () => {
    expect(
      chooseConsolidationAction({
        kind: "tool_success_profile",
        visibility: "internal",
        namespace: "global",
        content: "内部工具统计",
        metadata: { source: "telemetry", successCount: 3, failureCount: 2 },
      }),
    ).toBe("ignore");
  });

  it("ignores fractional and negative tool profile counts", () => {
    expect(
      chooseConsolidationAction({
        kind: "tool_success_profile",
        visibility: "internal",
        namespace: "global",
        content: "内部工具统计",
        metadata: { source: "telemetry", successCount: 2.5, failureCount: 1 },
      }),
    ).toBe("ignore");

    expect(
      chooseConsolidationAction({
        kind: "tool_success_profile",
        visibility: "internal",
        namespace: "global",
        content: "内部工具统计",
        metadata: { source: "telemetry", successCount: 4, failureCount: -1 },
      }),
    ).toBe("ignore");
  });

  it("adds user-visible corrections with content", () => {
    expect(
      chooseConsolidationAction({
        kind: "correction",
        visibility: "user",
        namespace: "global",
        content: "以后不要把临时文件路径写进长期记忆",
      }),
    ).toBe("add");
  });
});
