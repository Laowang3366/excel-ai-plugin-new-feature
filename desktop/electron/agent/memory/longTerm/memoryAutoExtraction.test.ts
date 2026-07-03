import { describe, expect, it, vi } from "vitest";

import type { RuntimeLongTermMemoryRecord } from "../stateRuntimeTypes";
import type { Thread, Turn } from "../../shared/types";
import {
  buildMemoryExtractionTranscript,
  extractAndWriteTurnMemories,
} from "./memoryAutoExtraction";
import type { MemoryWriteInput } from "./memoryTypes";

describe("memory auto extraction", () => {
  it("extracts and writes allowed user-visible memories after a completed turn", async () => {
    const turn = createCompletedTurn("请记住：以后回复先给结论。");
    const aiClient = {
      chat: vi.fn(async () => ({
        content: [
          "```json",
          JSON.stringify({
            memories: [
              {
                kind: "style_preference",
                namespace: "global",
                content: "用户偏好回复先给结论。",
                confidence: 0.9,
              },
              {
                kind: "project_fact",
                content: "不应写入的项目事实。",
              },
            ],
          }),
          "```",
        ].join("\n"),
      })),
    };
    const writtenInputs: MemoryWriteInput[] = [];
    const memoryStore = {
      search: vi.fn(async () => [] as RuntimeLongTermMemoryRecord[]),
      write: vi.fn(async (input: MemoryWriteInput) => {
        writtenInputs.push(input);
        return createMemoryRecord(input);
      }),
    };

    const result = await extractAndWriteTurnMemories({
      aiClient,
      memoryStore,
      thread: createThread(turn),
      turn,
    });

    expect(aiClient.chat).toHaveBeenCalledTimes(1);
    expect(result.written).toHaveLength(1);
    expect(writtenInputs[0]).toMatchObject({
      kind: "style_preference",
      namespace: "global",
      content: "用户偏好回复先给结论。",
      source: "extraction",
      sourceThreadId: "thread-1",
      citations: [{ threadId: "thread-1", turnId: "turn-1" }],
    });
    expect(writtenInputs[0].metadata).toMatchObject({
      extraction: { threadId: "thread-1", turnId: "turn-1" },
    });
  });

  it("skips exact duplicates when the memory store can search existing records", async () => {
    const turn = createCompletedTurn("以后优先使用 range.* 操作 Excel 区域。");
    const content = "用户偏好优先使用 range.* 操作 Excel 区域。";
    const aiClient = {
      chat: vi.fn(async () => ({
        content: JSON.stringify({
          memories: [
            {
              kind: "operation_preference",
              namespace: "global",
              content,
            },
          ],
        }),
      })),
    };
    const memoryStore = {
      search: vi.fn(async () => [createMemoryRecord({
        kind: "operation_preference",
        namespace: "global",
        content,
      })]),
      write: vi.fn(),
    };

    const result = await extractAndWriteTurnMemories({
      aiClient,
      memoryStore,
      thread: createThread(turn),
      turn,
    });

    expect(result.skippedDuplicates).toBe(1);
    expect(memoryStore.write).not.toHaveBeenCalled();
  });

  it("does not include successful tool result bodies in the extraction transcript", () => {
    const turn = createCompletedTurn("分析这个表。");
    turn.items.push({
      type: "tool_result",
      id: "tool-result-1",
      toolCallId: "tool-call-1",
      toolName: "range.read",
      result: { rows: [["secret table body"]] },
      isError: false,
      timestamp: Date.now(),
    });

    const transcript = buildMemoryExtractionTranscript(turn);

    expect(transcript).toContain("TOOL RESULT (range.read): completed");
    expect(transcript).not.toContain("secret table body");
  });
});

function createCompletedTurn(userContent: string): Turn {
  return {
    turnId: "turn-1",
    threadId: "thread-1",
    status: "completed",
    startedAt: Date.now(),
    completedAt: Date.now(),
    items: [
      {
        type: "user_message",
        id: "user-1",
        content: userContent,
        timestamp: Date.now(),
      },
      {
        type: "assistant_message",
        id: "assistant-1",
        content: "好的。",
        phase: "final",
        timestamp: Date.now(),
      },
    ],
  };
}

function createThread(turn: Turn): Thread {
  return {
    metadata: {
      threadId: "thread-1",
      preview: "",
      modelProvider: "openai",
      model: "test-model",
      createdAt: Date.now(),
      updatedAt: Date.now(),
    },
    turns: [turn],
  };
}

function createMemoryRecord(input: MemoryWriteInput): RuntimeLongTermMemoryRecord {
  return {
    memoryId: "memory-1",
    namespace: input.namespace ?? "global",
    kind: input.kind,
    visibility: "user",
    status: "active",
    content: input.content,
    confidence: input.confidence,
    sourceThreadId: input.sourceThreadId,
    metadata: input.metadata,
    citations: input.citations,
    createdAt: Date.now(),
    updatedAt: Date.now(),
  };
}
