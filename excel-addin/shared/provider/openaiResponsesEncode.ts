import type { ToolDefinition } from "../tools/types";
import type { AgentMessage } from "../agent/types";
import { encodeOpenAiResponsesContent } from "./messageContentEncode";
import type { ToolNameMaps } from "./openaiToolNameMap";

export type ResponsesEncodeOk = {
  instructions?: string;
  input: Record<string, unknown>[];
  tools: Record<string, unknown>[];
};

export type ResponsesEncodeErr = { error: string };

/** Encode agent history for OpenAI Responses API (stateless input array). */
export function encodeResponsesBody(
  systemPrompt: string,
  messages: AgentMessage[],
  tools: ToolDefinition[],
  maps: ToolNameMaps,
): ResponsesEncodeOk | ResponsesEncodeErr {
  const input: Record<string, unknown>[] = [];

  for (const msg of messages) {
    if (msg.role === "system" || msg.role === "user") {
      input.push({
        type: "message",
        role: msg.role,
        content:
          msg.role === "user" ? encodeOpenAiResponsesContent(msg) : msg.content,
      });
      continue;
    }

    if (msg.role === "assistant") {
      if (msg.content && msg.content.length > 0) {
        input.push({
          type: "message",
          role: "assistant",
          content: msg.content,
        });
      }
      if (msg.toolCalls && msg.toolCalls.length > 0) {
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
          input.push({
            type: "function_call",
            call_id: call.id,
            name: external,
            arguments: call.argumentsJson,
            status: "completed",
          });
        }
      }
      continue;
    }

    if (msg.role === "tool") {
      if (typeof msg.toolCallId !== "string" || msg.toolCallId.trim() === "") {
        return { error: "tool message missing tool_call_id" };
      }
      input.push({
        type: "function_call_output",
        call_id: msg.toolCallId,
        output: msg.content,
      });
    }
  }

  const encodedTools: Record<string, unknown>[] = [];
  for (const tool of tools) {
    const external = maps.internalToExternal.get(tool.name);
    if (external == null) {
      return { error: "tool encoding failed for active tools map" };
    }
    encodedTools.push({
      type: "function",
      name: external,
      description: tool.description,
      parameters: tool.parameters,
    });
  }

  const out: ResponsesEncodeOk = { input, tools: encodedTools };
  if (systemPrompt !== "") out.instructions = systemPrompt;
  return out;
}
