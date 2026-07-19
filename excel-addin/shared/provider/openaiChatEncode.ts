import type { ToolDefinition } from "../tools/types";
import type { AgentMessage } from "../agent/types";
import type { ToolNameMaps } from "./openaiToolNameMap";

export type EncodeOk = {
  messages: Record<string, unknown>[];
  tools: Record<string, unknown>[];
};

export type EncodeErr = { error: string };

/** Encode agent messages/tools for OpenAI chat.completions (stream). */
export function encodeChatCompletionsBody(
  systemPrompt: string,
  messages: AgentMessage[],
  tools: ToolDefinition[],
  maps: ToolNameMaps,
): EncodeOk | EncodeErr {
  const out: Record<string, unknown>[] = [];
  if (systemPrompt !== "") {
    out.push({ role: "system", content: systemPrompt });
  }
  for (const msg of messages) {
    if (msg.role === "system" || msg.role === "user") {
      out.push({ role: msg.role, content: msg.content });
      continue;
    }
    if (msg.role === "assistant") {
      const encoded: Record<string, unknown> = {
        role: "assistant",
        content: msg.content || null,
      };
      if (msg.toolCalls && msg.toolCalls.length > 0) {
        const tool_calls = [];
        for (const call of msg.toolCalls) {
          if (typeof call.id !== "string" || call.id.trim() === "") {
            return { error: "assistant tool call id is empty" };
          }
          const external = maps.internalToExternal.get(call.name);
          if (external == null) {
            return {
              error: `assistant tool call name not in active tools map: ${call.name}`,
            };
          }
          tool_calls.push({
            id: call.id,
            type: "function",
            function: {
              name: external,
              arguments: call.argumentsJson,
            },
          });
        }
        encoded.tool_calls = tool_calls;
      }
      out.push(encoded);
      continue;
    }
    if (msg.role === "tool") {
      if (typeof msg.toolCallId !== "string" || msg.toolCallId.trim() === "") {
        return { error: "tool message missing tool_call_id" };
      }
      out.push({
        role: "tool",
        tool_call_id: msg.toolCallId,
        content: msg.content,
      });
    }
  }

  const encodedTools = tools.map((tool) => {
    const external = maps.internalToExternal.get(tool.name);
    if (external == null) {
      return null;
    }
    return {
      type: "function",
      function: {
        name: external,
        description: tool.description,
        parameters: tool.parameters,
      },
    };
  });
  if (encodedTools.some((t) => t == null)) {
    return { error: "tool encoding failed for active tools map" };
  }

  return {
    messages: out,
    tools: encodedTools as Record<string, unknown>[],
  };
}
