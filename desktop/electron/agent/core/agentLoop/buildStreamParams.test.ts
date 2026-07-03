import { describe, expect, it } from "vitest";

import {
  appendRuntimeDateContext,
  appendLongTermMemoryContext,
  appendRuntimeLongTermMemoryContext,
} from "./buildStreamParams";

describe("appendRuntimeDateContext", () => {
  it("injects the current Shanghai date for relative-time tasks", () => {
    const prompt = appendRuntimeDateContext("base", new Date("2026-07-03T04:30:00.000Z"));

    expect(prompt).toContain("## 运行时上下文");
    expect(prompt).toContain("当前日期：2026");
    expect(prompt).toContain("07");
    expect(prompt).toContain("03");
    expect(prompt).toContain("Asia/Shanghai");
    expect(prompt).toContain("近 N 日");
    expect(prompt).toContain("不要自行补入过期年份");
  });
});

describe("appendLongTermMemoryContext", () => {
  it("injects only user-visible memory", () => {
    const prompt = appendLongTermMemoryContext("base", [
      {
        memoryId: "mem-user",
        namespace: "global",
        kind: "preference",
        visibility: "user",
        status: "active",
        content: "回复先给结论",
        createdAt: 1,
        updatedAt: 1,
      },
      {
        memoryId: "mem-internal",
        namespace: "global",
        kind: "tool_success_profile",
        visibility: "internal",
        status: "active",
        content: "内部工具统计",
        createdAt: 1,
        updatedAt: 1,
      },
    ]);

    expect(prompt).toContain("回复先给结论");
    expect(prompt).not.toContain("内部工具统计");
    expect(prompt).not.toContain("tool_success_profile");
  });

  it("injects only the six office memory kinds", () => {
    const prompt = appendLongTermMemoryContext("base", [
      {
        memoryId: "mem-project",
        namespace: "global",
        kind: "project_fact",
        visibility: "user",
        status: "active",
        content: "项目事实不应进普通提示词",
        createdAt: 1,
        updatedAt: 1,
      },
      {
        memoryId: "mem-workflow",
        namespace: "global",
        kind: "workflow",
        visibility: "user",
        status: "active",
        content: "工作流不应进普通提示词",
        createdAt: 1,
        updatedAt: 1,
      },
      {
        memoryId: "mem-file",
        namespace: "global",
        kind: "file_impression",
        visibility: "user",
        status: "active",
        content: "这个文件常用数据透视表",
        createdAt: 1,
        updatedAt: 1,
      },
    ]);

    expect(prompt).toContain("这个文件常用数据透视表");
    expect(prompt).not.toContain("项目事实不应进普通提示词");
    expect(prompt).not.toContain("工作流不应进普通提示词");
    expect(prompt).not.toContain("project_fact");
    expect(prompt).not.toContain("workflow");
  });

  it("returns the original prompt when there are no active user-visible memories", () => {
    const prompt = appendLongTermMemoryContext("base", [
      {
        memoryId: "mem-internal",
        namespace: "global",
        kind: "tool_success_profile",
        visibility: "internal",
        status: "active",
        content: "内部工具统计",
        createdAt: 1,
        updatedAt: 1,
      },
    ]);

    expect(prompt).toBe("base");
  });

  it("does not inject stale user memory", () => {
    const prompt = appendLongTermMemoryContext("base", [
      {
        memoryId: "mem-stale",
        namespace: "global",
        kind: "preference",
        visibility: "user",
        status: "stale",
        content: "过期偏好",
        createdAt: 1,
        updatedAt: 1,
      },
    ]);

    expect(prompt).toBe("base");
  });

  it("prefers summary over content when present", () => {
    const prompt = appendLongTermMemoryContext("base", [
      {
        memoryId: "mem-summary",
        namespace: "global",
        kind: "style_preference",
        visibility: "user",
        status: "active",
        content: "很长的原始记忆内容",
        summary: "简短摘要",
        createdAt: 1,
        updatedAt: 1,
      },
    ]);

    expect(prompt).toContain("- [style_preference] 简短摘要");
    expect(prompt).not.toContain("很长的原始记忆内容");
  });
});

describe("appendRuntimeLongTermMemoryContext", () => {
  it("loads active user memory from runtime and filters internal/non-office kinds", async () => {
    const prompt = await appendRuntimeLongTermMemoryContext("base", {
      listLongTermMemories: async (options) => {
        expect(options).toMatchObject({
          visibility: "user",
          status: "active",
          limit: 8,
        });
        return [
          {
            memoryId: "mem-pref",
            namespace: "global",
            kind: "preference",
            visibility: "user",
            status: "active",
            content: "先给结论",
            createdAt: 1,
            updatedAt: 1,
          },
          {
            memoryId: "mem-tool",
            namespace: "global",
            kind: "tool_success_profile",
            visibility: "internal",
            status: "active",
            content: "内部工具画像",
            createdAt: 1,
            updatedAt: 1,
          },
          {
            memoryId: "mem-project",
            namespace: "global",
            kind: "project_fact",
            visibility: "user",
            status: "active",
            content: "项目事实",
            createdAt: 1,
            updatedAt: 1,
          },
        ];
      },
    });

    expect(prompt).toContain("先给结论");
    expect(prompt).not.toContain("内部工具画像");
    expect(prompt).not.toContain("项目事实");
  });

  it("continues past legacy non-office pages to collect allowed memories", async () => {
    const calls: Array<{ limit?: number; offset?: number }> = [];
    const legacyPage = Array.from({ length: 8 }, (_, index) => ({
      memoryId: `legacy-${index}`,
      namespace: "global",
      kind: index % 2 === 0 ? "project_fact" as const : "workflow" as const,
      visibility: "user" as const,
      status: "active" as const,
      content: `legacy page memory ${index}`,
      createdAt: index,
      updatedAt: index,
    }));

    const prompt = await appendRuntimeLongTermMemoryContext("base", {
      listLongTermMemories: async (options) => {
        const listOptions = options ?? {};
        calls.push({ limit: listOptions.limit, offset: listOptions.offset });
        if (listOptions.offset === 0 || listOptions.offset === undefined) {
          return legacyPage;
        }
        return [
          {
            memoryId: "allowed-pref",
            namespace: "global",
            kind: "preference",
            visibility: "user",
            status: "active",
            content: "合法偏好在第二页",
            createdAt: 10,
            updatedAt: 10,
          },
        ];
      },
    });

    expect(calls).toEqual([
      { limit: 8, offset: 0 },
      { limit: 8, offset: 8 },
    ]);
    expect(prompt).toContain("合法偏好在第二页");
    expect(prompt).not.toContain("legacy page memory");
    expect(prompt).not.toContain("project_fact");
    expect(prompt).not.toContain("workflow");
  });

  it("keeps the base prompt when runtime memory loading fails", async () => {
    const prompt = await appendRuntimeLongTermMemoryContext("base", {
      listLongTermMemories: async () => {
        throw new Error("db unavailable");
      },
    });

    expect(prompt).toBe("base");
  });
});
