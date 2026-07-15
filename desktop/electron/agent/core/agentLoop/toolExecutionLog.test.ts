import { describe, expect, it, vi } from "vitest";

import {
  logToolExecutionSafely,
  summarizeForLog,
  type ToolExecutionLogRecord,
} from "./toolExecutionLog";

const CANARY = "sk-1234567890abcdefghijklmnop";

describe("toolExecutionLog", () => {
  it("stores structural summaries instead of argument and result content", () => {
    const summary = summarizeForLog({
      query: `private forecast ${CANARY}`,
      values: [["customer-a", 42]],
    });

    expect(summary).not.toContain(CANARY);
    expect(summary).not.toContain("private forecast");
    expect(summary).not.toContain("customer-a");
    expect(JSON.parse(summary)).toMatchObject({
      type: "object",
      fields: 2,
      stats: { arrays: 2, strings: 2, numbers: 1 },
    });
  });

  it("redacts errors and metadata at the persistence boundary", async () => {
    const append = vi.fn(async (_record: ToolExecutionLogRecord) => undefined);
    await logToolExecutionSafely(append, {
      threadId: "thread-1",
      turnId: "turn-1",
      toolCallId: "call-1",
      toolName: "web.search",
      status: "error",
      durationMs: 1,
      timestamp: 1,
      argumentsSummary: summarizeForLog({ query: CANARY }),
      resultSummary: summarizeForLog({ success: false }),
      error: `provider rejected ${CANARY}`,
      metadata: { authorization: `Bearer ${CANARY}` },
    }, { onEvent: vi.fn() });

    expect(append).toHaveBeenCalledOnce();
    const persisted = append.mock.calls[0]![0];
    expect(JSON.stringify(persisted)).not.toContain(CANARY);
    expect(persisted.error).toContain("[REDACTED:openai-style-key]");
    expect(persisted.metadata).toEqual({ authorization: "[REDACTED]" });
  });
});
