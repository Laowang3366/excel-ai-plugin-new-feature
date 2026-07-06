import type { ToolResultItem } from "../../shared/types";

export interface ToolResultItemParams {
  toolCallId: string;
  toolName: string;
  result: unknown;
  isError: boolean;
  timestamp?: number;
}

export function createToolResultItem({
  toolCallId,
  toolName,
  result,
  isError,
  timestamp = Date.now(),
}: ToolResultItemParams): ToolResultItem {
  return {
    type: "tool_result",
    id: `result-${timestamp}`,
    toolCallId,
    toolName,
    result,
    isError,
    timestamp,
  };
}
