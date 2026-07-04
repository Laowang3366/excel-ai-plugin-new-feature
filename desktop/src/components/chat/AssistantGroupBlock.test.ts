import { describe, expect, it } from "vitest";
import type { TurnItem } from "../../electronApi";
import { sortItemsByRound } from "./AssistantGroupBlock";

describe("sortItemsByRound", () => {
  it("preserves the real event timeline inside assistant details", () => {
    const items: TurnItem[] = [
      {
        type: "assistant_message",
        id: "commentary-1",
        content: "Working...",
        phase: "commentary",
        timestamp: 1000,
      },
      {
        type: "tool_call",
        id: "tool-1",
        toolName: "word.create",
        arguments: {},
        status: "completed",
        timestamp: 2000,
      },
      {
        type: "tool_result",
        id: "result-1",
        toolCallId: "tool-1",
        toolName: "word.create",
        result: { success: true },
        isError: false,
        timestamp: 3000,
      },
      {
        type: "assistant_message",
        id: "final-1",
        content: "Done.",
        phase: "final",
        timestamp: 4000,
      },
    ];

    expect(sortItemsByRound(items).map((item) => item.id)).toEqual([
      "commentary-1",
      "tool-1",
      "result-1",
      "final-1",
    ]);
  });
});
