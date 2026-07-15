import { describe, expect, it } from "vitest";

import { parseStageOneOutput, shouldIgnoreCandidateContent } from "./memoryExtraction";

describe("memory extraction", () => {
  it("ignores temporary path candidate content", () => {
    expect(
      shouldIgnoreCandidateContent("临时路径 C:\\Users\\wfq\\AppData\\Local\\Temp\\make.py"),
    ).toBe(true);
  });

  it("parses stage one JSON output with kind-derived visibility", () => {
    const memories = parseStageOneOutput(
      JSON.stringify({
        memories: [
          {
            kind: "operation_preference",
            namespace: "global",
            content: "优先使用稳定的文件级编辑",
            confidence: 0.8,
            citations: [{ threadId: "thread-1", eventId: 1 }],
          },
        ],
      }),
    );

    expect(memories).toHaveLength(1);
    expect(memories[0]).toMatchObject({
      kind: "operation_preference",
      visibility: "user",
      namespace: "global",
      content: "优先使用稳定的文件级编辑",
      confidence: 0.8,
      citations: [{ threadId: "thread-1", eventId: 1 }],
    });
  });

  it("filters stage one output to tool-writable office memory kinds", () => {
    const memories = parseStageOneOutput(
      JSON.stringify({
        memories: [
          {
            kind: "preference",
            content: "先给结论",
          },
          {
            kind: "project_fact",
            content: "项目事实不进入普通提取",
          },
          {
            kind: "workflow",
            content: "工作流不进入普通提取",
          },
          {
            kind: "tool_success_profile",
            content: "模型不能伪造内部工具画像",
            metadata: { source: "telemetry", successCount: 3, failureCount: 0 },
          },
        ],
      }),
    );

    expect(memories).toHaveLength(1);
    expect(memories[0].kind).toBe("preference");
  });

  it("keeps only valid citation fields", () => {
    const memories = parseStageOneOutput(
      JSON.stringify({
        memories: [
          {
            kind: "operation_preference",
            content: "优先使用稳定的文件级编辑",
            citations: [
              {
                threadId: "thread-1",
                eventId: "1",
                turnId: 2,
                extra: "drop",
              },
              {
                threadId: "thread-2",
                eventId: 2,
                turnId: "turn-2",
                extra: "drop",
              },
            ],
          },
        ],
      }),
    );

    expect(memories[0].citations).toEqual([
      { threadId: "thread-1" },
      { threadId: "thread-2", eventId: 2, turnId: "turn-2" },
    ]);
  });

  it("ignores TSV table dumps without blocking ordinary preferences", () => {
    expect(shouldIgnoreCandidateContent("姓名\t部门\t金额\n张三\t销售\t10\n李四\t财务\t20")).toBe(
      true,
    );

    expect(shouldIgnoreCandidateContent("我偏好先用稳定的文件级编辑，再考虑脚本")).toBe(false);
  });

  it("ignores CSV table dumps", () => {
    expect(shouldIgnoreCandidateContent("姓名,部门,金额\n张三,销售,10\n李四,财务,20")).toBe(true);
  });
});
