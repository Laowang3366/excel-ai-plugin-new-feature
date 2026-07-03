import { describe, expect, it } from "vitest";

import type {
  RuntimeMemoryKind,
  RuntimeMemoryVisibility,
} from "../stateRuntimeTypes";
import {
  getMemoryVisibility,
  isUserVisibleMemoryKind,
  normalizeMemoryWriteInput,
} from "./memoryTypes";

describe("long term memory types", () => {
  it("pins visibility for every memory kind", () => {
    const visibilityByKind: Record<RuntimeMemoryKind, RuntimeMemoryVisibility> =
      {
        preference: "user",
        constraint: "user",
        correction: "user",
        style_preference: "user",
        operation_preference: "user",
        file_impression: "user",
        project_fact: "internal",
        workflow: "internal",
        tool_success_profile: "internal",
      };

    for (const [kind, visibility] of Object.entries(visibilityByKind) as [
      RuntimeMemoryKind,
      RuntimeMemoryVisibility,
    ][]) {
      expect(getMemoryVisibility(kind)).toBe(visibility);
      expect(isUserVisibleMemoryKind(kind)).toBe(visibility === "user");
    }
  });

  it("keeps tool success profiles internal", () => {
    expect(getMemoryVisibility("tool_success_profile")).toBe("internal");
    expect(isUserVisibleMemoryKind("tool_success_profile")).toBe(false);
  });

  it("keeps user preference kinds visible", () => {
    expect(getMemoryVisibility("preference")).toBe("user");
    expect(getMemoryVisibility("operation_preference")).toBe("user");
    expect(isUserVisibleMemoryKind("file_impression")).toBe(true);
    expect(isUserVisibleMemoryKind("project_fact")).toBe(false);
    expect(isUserVisibleMemoryKind("workflow")).toBe(false);
  });

  it("rejects internal tool profiles from ordinary tool writes", () => {
    expect(() =>
      normalizeMemoryWriteInput({
        kind: "tool_success_profile",
        namespace: "global",
        content: "内部工具统计",
        source: "tool",
      }),
    ).toThrow("tool_success_profile 只能由内部遥测写入");
  });

  it("normalizes namespaces without treating explicit blanks as global", () => {
    expect(
      normalizeMemoryWriteInput({
        kind: "preference",
        content: "Prefer concise updates",
      }).namespace,
    ).toBe("global");

    expect(
      normalizeMemoryWriteInput({
        kind: "preference",
        namespace: " team ",
        content: "Prefer concise updates",
      }).namespace,
    ).toBe("team");

    expect(() =>
      normalizeMemoryWriteInput({
        kind: "preference",
        namespace: "   ",
        content: "Prefer concise updates",
      }),
    ).toThrow("记忆命名空间不能为空");
  });

  it("keeps top-level source authoritative over metadata source", () => {
    expect(
      normalizeMemoryWriteInput({
        kind: "preference",
        content: "Prefer concise updates",
        source: "extraction",
        metadata: {
          source: "tool",
        },
      }).metadata?.source,
    ).toBe("extraction");
  });

  it("preserves extraction provenance fields", () => {
    expect(
      normalizeMemoryWriteInput({
        kind: "preference",
        content: "Prefer concise updates",
        source: "extraction",
        sourceThreadId: "thread-1",
        citations: [{ threadId: "thread-1", turnId: "turn-1" }],
      }),
    ).toMatchObject({
      sourceThreadId: "thread-1",
      citations: [{ threadId: "thread-1", turnId: "turn-1" }],
      metadata: { source: "extraction" },
    });
  });

  it("allows telemetry writes for tool success profiles", () => {
    expect(
      normalizeMemoryWriteInput({
        kind: "tool_success_profile",
        content: "内部工具统计",
        source: "telemetry",
      }),
    ).toMatchObject({
      kind: "tool_success_profile",
      visibility: "internal",
      metadata: {
        source: "telemetry",
      },
    });
  });

  it("clamps confidence to the valid range", () => {
    expect(
      normalizeMemoryWriteInput({
        kind: "preference",
        content: "Prefer concise updates",
        confidence: 1.4,
      }).confidence,
    ).toBe(1);

    expect(
      normalizeMemoryWriteInput({
        kind: "preference",
        content: "Prefer concise updates",
        confidence: -0.4,
      }).confidence,
    ).toBe(0);

    expect(
      normalizeMemoryWriteInput({
        kind: "preference",
        content: "Prefer concise updates",
        confidence: Number.NaN,
      }).confidence,
    ).toBeUndefined();

    expect(
      normalizeMemoryWriteInput({
        kind: "preference",
        content: "Prefer concise updates",
        confidence: Number.POSITIVE_INFINITY,
      }).confidence,
    ).toBeUndefined();
  });
});
