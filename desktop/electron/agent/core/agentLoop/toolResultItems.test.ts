import { describe, expect, it } from "vitest";
import { createToolResultItem } from "./toolResultItems";

describe("createToolResultItem", () => {
  it("creates a successful tool result item with a stable timestamp", () => {
    expect(createToolResultItem({
      toolCallId: "call-1",
      toolName: "range.read",
      result: { values: [[1]] },
      isError: false,
      timestamp: 1234,
    })).toEqual({
      type: "tool_result",
      id: "result-1234",
      toolCallId: "call-1",
      toolName: "range.read",
      result: { values: [[1]] },
      isError: false,
      timestamp: 1234,
    });
  });

  it("creates an error tool result item", () => {
    expect(createToolResultItem({
      toolCallId: "call-2",
      toolName: "shell.execute",
      result: "blocked",
      isError: true,
      timestamp: 5678,
    })).toMatchObject({
      type: "tool_result",
      id: "result-5678",
      toolCallId: "call-2",
      toolName: "shell.execute",
      result: "blocked",
      isError: true,
      timestamp: 5678,
    });
  });
});
