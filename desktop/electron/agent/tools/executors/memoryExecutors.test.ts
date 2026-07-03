import { describe, expect, it } from "vitest";

import { StateRuntimeStore } from "../../memory/stateRuntimeStore";
import { LongTermMemoryStore } from "../../memory/longTerm/memoryStore";
import { addMemoryExecutors } from "./memoryExecutors";

describe("memory executors", () => {
  it("returns a clear error when memory store is unavailable", async () => {
    const executors = new Map();
    addMemoryExecutors(executors, {});

    const result = await executors.get("memory.write").execute({
      kind: "preference",
      content: "回复先给结论",
    });

    expect(result).toEqual({
      success: false,
      error: "长期记忆尚未初始化",
    });
  });

  it("writes and searches user-visible memory", async () => {
    const runtime = new StateRuntimeStore(":memory:");
    await runtime.init();
    const memoryStore = new LongTermMemoryStore(runtime);
    const executors = new Map();
    addMemoryExecutors(executors, { memoryStore });

    const write = await executors.get("memory.write").execute({
      kind: "preference",
      content: "回复先给结论",
    });
    expect(write.success).toBe(true);

    const search = await executors.get("memory.search").execute({ query: "结论" });
    expect(search.success).toBe(true);
    expect(search.data[0].content).toBe("回复先给结论");
    await runtime.close();
  });

  it("deletes user-visible memory by archiving it", async () => {
    const runtime = new StateRuntimeStore(":memory:");
    await runtime.init();
    const memoryStore = new LongTermMemoryStore(runtime);
    const executors = new Map();
    addMemoryExecutors(executors, { memoryStore });

    const write = await executors.get("memory.write").execute({
      kind: "preference",
      content: "以后不要保留这条偏好",
    });
    expect(write.success).toBe(true);

    const deleted = await executors.get("memory.delete").execute({
      memoryId: write.data.memoryId,
    });
    const search = await executors.get("memory.search").execute({ query: "偏好" });

    expect(deleted.success).toBe(true);
    expect(deleted.data).toMatchObject({
      memoryId: write.data.memoryId,
      status: "archived",
    });
    expect(search).toEqual({ success: true, data: [] });
    await runtime.close();
  });

  it("returns a clear error when deleting a missing memory", async () => {
    const runtime = new StateRuntimeStore(":memory:");
    await runtime.init();
    const executors = new Map();
    addMemoryExecutors(executors, { memoryStore: new LongTermMemoryStore(runtime) });

    const result = await executors.get("memory.delete").execute({
      memoryId: "missing-memory",
    });

    expect(result).toEqual({
      success: false,
      error: "未找到可删除的长期记忆",
    });
    await runtime.close();
  });

  it("rejects invalid delete arguments", async () => {
    const runtime = new StateRuntimeStore(":memory:");
    await runtime.init();
    const executors = new Map();
    addMemoryExecutors(executors, { memoryStore: new LongTermMemoryStore(runtime) });

    const result = await executors.get("memory.delete").execute({
      memoryId: 123,
    });

    expect(result).toEqual({
      success: false,
      error: "参数 memoryId 应为字符串，实际为 number",
    });
    await runtime.close();
  });

  it("returns a clear delete error when memory store is unavailable", async () => {
    const executors = new Map();
    addMemoryExecutors(executors, {});

    const result = await executors.get("memory.delete").execute({
      memoryId: "memory-1",
    });

    expect(result).toEqual({
      success: false,
      error: "长期记忆尚未初始化",
    });
  });

  it("rejects direct writes to internal tool profiles", async () => {
    const runtime = new StateRuntimeStore(":memory:");
    await runtime.init();
    const executors = new Map();
    addMemoryExecutors(executors, { memoryStore: new LongTermMemoryStore(runtime) });

    const result = await executors.get("memory.write").execute({
      kind: "tool_success_profile",
      content: "内部统计",
    });
    expect(result.success).toBe(false);
    expect(result.error).toContain("tool_success_profile");
    await runtime.close();
  });

  it("rejects writes to user-visible kinds that tools cannot write directly", async () => {
    const runtime = new StateRuntimeStore(":memory:");
    await runtime.init();
    const executors = new Map();
    addMemoryExecutors(executors, { memoryStore: new LongTermMemoryStore(runtime) });

    const projectFact = await executors.get("memory.write").execute({
      kind: "project_fact",
      content: "项目事实",
    });
    const workflow = await executors.get("memory.write").execute({
      kind: "workflow",
      content: "工作流",
    });

    expect(projectFact.success).toBe(false);
    expect(projectFact.error).toContain("project_fact");
    expect(workflow.success).toBe(false);
    expect(workflow.error).toContain("workflow");
    await runtime.close();
  });

  it("rejects invalid search filters", async () => {
    const runtime = new StateRuntimeStore(":memory:");
    await runtime.init();
    const executors = new Map();
    addMemoryExecutors(executors, { memoryStore: new LongTermMemoryStore(runtime) });

    const internalKind = await executors.get("memory.search").execute({
      query: "统计",
      kind: "tool_success_profile",
    });
    const invalidNamespace = await executors.get("memory.search").execute({
      query: "结论",
      namespace: 123,
    });
    const invalidLimit = await executors.get("memory.search").execute({
      query: "结论",
      limit: "10",
    });

    expect(internalKind).toEqual({
      success: false,
      error: expect.stringContaining("参数 kind 必须是"),
    });
    expect(invalidNamespace).toEqual({
      success: false,
      error: "参数 namespace 必须是 string",
    });
    expect(invalidLimit).toEqual({
      success: false,
      error: "参数 limit 必须是 number",
    });
    await runtime.close();
  });

  it("lists only user-visible memory by default", async () => {
    const runtime = new StateRuntimeStore(":memory:");
    await runtime.init();
    const memoryStore = new LongTermMemoryStore(runtime);
    const executors = new Map();
    addMemoryExecutors(executors, { memoryStore });

    await memoryStore.write({
      kind: "preference",
      content: "用户偏好",
      source: "tool",
    });
    await memoryStore.write({
      kind: "tool_success_profile",
      content: "内部统计",
      source: "telemetry",
    });

    const result = await executors.get("memory.list").execute({});

    expect(result.success).toBe(true);
    expect(result.data.map((record: { content: string }) => record.content)).toEqual(["用户偏好"]);
    await runtime.close();
  });

  it("rejects invalid list namespace", async () => {
    const runtime = new StateRuntimeStore(":memory:");
    await runtime.init();
    const executors = new Map();
    addMemoryExecutors(executors, { memoryStore: new LongTermMemoryStore(runtime) });

    const result = await executors.get("memory.list").execute({ namespace: 123 });

    expect(result).toEqual({
      success: false,
      error: "参数 namespace 必须是 string",
    });
    await runtime.close();
  });
});
