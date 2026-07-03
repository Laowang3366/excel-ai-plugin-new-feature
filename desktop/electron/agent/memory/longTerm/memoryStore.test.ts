import { describe, expect, it } from "vitest";

import { StateRuntimeStore } from "../stateRuntimeStore";
import { LongTermMemoryStore } from "./memoryStore";

describe("LongTermMemoryStore", () => {
  it("can switch to a new runtime store after data-path migration", async () => {
    const oldRuntime = new StateRuntimeStore(":memory:");
    const newRuntime = new StateRuntimeStore(":memory:");
    await oldRuntime.init();
    await newRuntime.init();
    const store = new LongTermMemoryStore(oldRuntime);

    store.updateRuntime(newRuntime);
    await store.write({
      kind: "preference",
      namespace: "global",
      content: "write to migrated runtime",
      source: "tool",
    });

    expect(await oldRuntime.listLongTermMemories({ limit: 10 })).toHaveLength(0);
    expect((await newRuntime.listLongTermMemories({ limit: 10 }))[0].content).toBe("write to migrated runtime");

    await oldRuntime.close();
    await newRuntime.close();
  });

  it("writes user-visible memories and hides internal profiles by default", async () => {
    const runtime = new StateRuntimeStore(":memory:");
    await runtime.init();
    const store = new LongTermMemoryStore(runtime);

    await store.write({
      kind: "preference",
      namespace: "global",
      content: "回复先给结论",
      source: "tool",
    });
    await store.write({
      kind: "tool_success_profile",
      namespace: "global",
      content: "PPT 文件级编辑成功率更高",
      source: "telemetry",
      metadata: { successCount: 3, failureCount: 1 },
    });

    const visibleResults = await store.search({ query: "结论" });
    expect(visibleResults).toHaveLength(1);
    expect(visibleResults[0]).toMatchObject({
      kind: "preference",
      visibility: "user",
      content: "回复先给结论",
    });

    const listed = await store.list("global");
    expect(listed.map((m) => m.kind)).toContain("preference");
    expect(listed.map((m) => m.kind)).not.toContain("tool_success_profile");

    expect((await store.search({ query: "PPT" })).map((m) => m.kind)).not.toContain(
      "tool_success_profile",
    );
    expect(
      (await store.search({ query: "PPT", includeInternal: true })).map(
        (m) => m.kind,
      ),
    ).toContain("tool_success_profile");
    await runtime.close();
  });

  it("hides legacy user-visible non-office memories from public list and search", async () => {
    const runtime = new StateRuntimeStore(":memory:");
    await runtime.init();
    const store = new LongTermMemoryStore(runtime);
    const now = Date.now();

    await runtime.upsertLongTermMemory({
      memoryId: "legacy-project",
      kind: "project_fact",
      namespace: "global",
      visibility: "user",
      status: "active",
      content: "legacy leak marker project",
      createdAt: now + 2,
      updatedAt: now + 2,
    });
    await runtime.upsertLongTermMemory({
      memoryId: "legacy-workflow",
      kind: "workflow",
      namespace: "global",
      visibility: "user",
      status: "active",
      content: "legacy leak marker workflow",
      createdAt: now + 1,
      updatedAt: now + 1,
    });
    await runtime.upsertLongTermMemory({
      memoryId: "visible-preference",
      kind: "preference",
      namespace: "global",
      visibility: "user",
      status: "active",
      content: "legacy leak marker preference",
      createdAt: now,
      updatedAt: now,
    });

    const listed = await store.list("global");
    expect(listed.map((memory) => memory.memoryId)).toEqual([
      "visible-preference",
    ]);

    const searched = await store.search({ query: "legacy leak marker" });
    expect(searched.map((memory) => memory.memoryId)).toEqual([
      "visible-preference",
    ]);
    await runtime.close();
  });

  it("applies public search limits after filtering legacy non-office kinds", async () => {
    const runtime = new StateRuntimeStore(":memory:");
    await runtime.init();
    const store = new LongTermMemoryStore(runtime);
    const now = Date.now();

    await runtime.upsertLongTermMemory({
      memoryId: "visible-needle",
      kind: "preference",
      namespace: "global",
      visibility: "user",
      status: "active",
      content: "needle visible",
      createdAt: now,
      updatedAt: now,
    });
    await runtime.upsertLongTermMemory({
      memoryId: "legacy-needle",
      kind: "project_fact",
      namespace: "global",
      visibility: "user",
      status: "active",
      content: "needle legacy",
      createdAt: now + 1,
      updatedAt: now + 1,
    });

    const results = await store.search({ query: "needle", limit: 1 });

    expect(results).toHaveLength(1);
    expect(results[0].memoryId).toBe("visible-needle");
    await runtime.close();
  });

  it("finds older query matches beyond the first candidate page", async () => {
    const runtime = new StateRuntimeStore(":memory:");
    await runtime.init();
    const store = new LongTermMemoryStore(runtime);
    const now = Date.now();

    await runtime.upsertLongTermMemory({
      memoryId: "older-needle",
      kind: "preference",
      namespace: "global",
      visibility: "user",
      status: "active",
      content: "needle lives in an older memory",
      createdAt: now,
      updatedAt: now,
    });

    for (let i = 0; i < 105; i += 1) {
      await runtime.upsertLongTermMemory({
        memoryId: `newer-${i}`,
        kind: "preference",
        namespace: "global",
        visibility: "user",
        status: "active",
        content: `newer memory ${i}`,
        createdAt: now + i + 1,
        updatedAt: now + i + 1,
      });
    }

    const results = await store.search({ query: "needle" });

    expect(results.map((memory) => memory.memoryId)).toContain("older-needle");
    await runtime.close();
  });

  it("applies search limits after filtering query matches", async () => {
    const runtime = new StateRuntimeStore(":memory:");
    await runtime.init();
    const store = new LongTermMemoryStore(runtime);
    const now = Date.now();

    await runtime.upsertLongTermMemory({
      memoryId: "older-needle-1",
      kind: "preference",
      namespace: "global",
      visibility: "user",
      status: "active",
      content: "needle first",
      createdAt: now,
      updatedAt: now,
    });
    await runtime.upsertLongTermMemory({
      memoryId: "older-needle-2",
      kind: "preference",
      namespace: "global",
      visibility: "user",
      status: "active",
      content: "needle second",
      createdAt: now + 1,
      updatedAt: now + 1,
    });
    await runtime.upsertLongTermMemory({
      memoryId: "newer-nonmatch",
      kind: "preference",
      namespace: "global",
      visibility: "user",
      status: "active",
      content: "latest unrelated memory",
      createdAt: now + 2,
      updatedAt: now + 2,
    });

    const results = await store.search({ query: "needle", limit: 1 });

    expect(results).toHaveLength(1);
    expect(results[0].content).toContain("needle");
    await runtime.close();
  });

  it("deletes user-visible memories by archiving them", async () => {
    const runtime = new StateRuntimeStore(":memory:");
    await runtime.init();
    const store = new LongTermMemoryStore(runtime);

    const record = await store.write({
      kind: "preference",
      namespace: "global",
      content: "不要再保留这条偏好",
      source: "tool",
    });

    const deleted = await store.delete(record.memoryId);

    expect(deleted).toMatchObject({
      memoryId: record.memoryId,
      status: "archived",
    });
    expect(await store.search({ query: "偏好" })).toEqual([]);
    expect(await store.delete(record.memoryId)).toBeNull();
    await runtime.close();
  });

  it("does not allow tool deletion of internal memories", async () => {
    const runtime = new StateRuntimeStore(":memory:");
    await runtime.init();
    const store = new LongTermMemoryStore(runtime);

    await store.write({
      kind: "tool_success_profile",
      namespace: "global",
      content: "内部工具统计",
      source: "telemetry",
    });
    const internal = (await store.search({ includeInternal: true, query: "内部工具统计" }))[0];

    await expect(store.delete(internal.memoryId)).rejects.toThrow("只能删除用户可见的长期记忆");
    expect((await store.search({ includeInternal: true, query: "内部工具统计" }))[0].status).toBe("active");
    await runtime.close();
  });
});
