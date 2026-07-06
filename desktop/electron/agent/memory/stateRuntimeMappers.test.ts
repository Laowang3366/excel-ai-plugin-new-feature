import { describe, expect, it } from "vitest";
import {
  buildRolloutFtsQuery,
  clampMemoryListOffset,
  mapThreadSnapshot,
  mapToolExecutionLog,
} from "./stateRuntimeMappers";

describe("state runtime row mappers", () => {
  it("maps thread snapshot rows and parses JSON fields", () => {
    expect(mapThreadSnapshot({
      thread_id: "thread-1",
      preview: "Preview",
      name: null,
      model_provider: "openai",
      model: "gpt-5",
      context_window_size: 128000,
      created_at: 1,
      updated_at: 2,
      active_turn_id: null,
      last_turn_status: "completed",
      total_token_usage: "{\"inputTokens\":10,\"outputTokens\":4}",
      archived_at: null,
      folder_id: "C:/Work",
      compacted_history: "[{\"type\":\"assistant_message\",\"id\":\"a1\",\"content\":\"done\",\"timestamp\":3}]",
    })).toMatchObject({
      threadId: "thread-1",
      name: undefined,
      modelProvider: "openai",
      totalTokenUsage: { inputTokens: 10, outputTokens: 4 },
      folderId: "C:/Work",
      compactedHistory: [
        { type: "assistant_message", id: "a1", content: "done", timestamp: 3 },
      ],
    });
  });

  it("maps tool execution log metadata JSON", () => {
    expect(mapToolExecutionLog({
      id: 7,
      thread_id: "thread-1",
      turn_id: "turn-1",
      tool_call_id: "call-1",
      tool_name: "range.read",
      status: "success",
      duration_ms: 12,
      timestamp: 100,
      arguments_summary: "{}",
      result_summary: "{\"ok\":true}",
      error: null,
      metadata_json: "{\"riskLevel\":\"safe\"}",
    })).toMatchObject({
      id: 7,
      toolName: "range.read",
      metadata: { riskLevel: "safe" },
      error: undefined,
    });
  });
});

describe("state runtime search helpers", () => {
  it("quotes FTS terms and escapes embedded quotes", () => {
    expect(buildRolloutFtsQuery("hello \"quoted\" world")).toBe("\"hello\" \"\"\"quoted\"\"\" \"world\"");
  });

  it("clamps memory list offsets to non-negative integers", () => {
    expect(clampMemoryListOffset(undefined)).toBe(0);
    expect(clampMemoryListOffset(Number.NaN)).toBe(0);
    expect(clampMemoryListOffset(-3)).toBe(0);
    expect(clampMemoryListOffset(2.8)).toBe(2);
  });
});
