import type { ToolDefinition } from "../tools/types";
import type { AgentMessage } from "../agent/types";
import type { ToolNameMaps } from "./openaiToolNameMap";

export type AnthropicEncodeOk = {
  system?: string;
  messages: Record<string, unknown>[];
  tools: Record<string, unknown>[];
};

export type AnthropicEncodeErr = { error: string };

function isPlainObject(value: unknown): value is Record<string, unknown> {
  return value != null && typeof value === "object" && !Array.isArray(value);
}

/** Encode agent history for Anthropic Messages API (stream). */
export function encodeAnthropicMessagesBody(
  systemPrompt: string,
  messages: AgentMessage[],
  tools: ToolDefinition[],
  maps: ToolNameMaps,
): AnthropicEncodeOk | AnthropicEncodeErr {
  const out: Record<string, unknown>[] = [];
  let i = 0;
  while (i < messages.length) {
    const msg = messages[i]!;

    // History system messages are skipped; systemPrompt is top-level only.
    if (msg.role === "system") {
      i += 1;
      continue;
    }

    if (msg.role === "user") {
      out.push({ role: "user", content: msg.content });
      i += 1;
      continue;
    }

    if (msg.role === "assistant") {
      if (!msg.toolCalls || msg.toolCalls.length === 0) {
        out.push({ role: "assistant", content: msg.content });
        i += 1;
        continue;
      }
      const blocks: Record<string, unknown>[] = [];
      if (msg.content && msg.content.length > 0) {
        blocks.push({ type: "text", text: msg.content });
      }
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
        let input: unknown;
        try {
          input = JSON.parse(call.argumentsJson || "{}");
        } catch {
          return { error: `assistant tool call arguments are not valid JSON: ${call.id}` };
        }
        if (!isPlainObject(input)) {
          return { error: `assistant tool call arguments must be a plain object: ${call.id}` };
        }
        blocks.push({
          type: "tool_use",
          id: call.id,
          name: external,
          input,
        });
      }
      out.push({ role: "assistant", content: blocks });
      i += 1;
      continue;
    }

    if (msg.role === "tool") {
      // Merge consecutive tool messages into one user content array of tool_result.
      const results: Record<string, unknown>[] = [];
      while (i < messages.length && messages[i]!.role === "tool") {
        const toolMsg = messages[i]!;
        if (typeof toolMsg.toolCallId !== "string" || toolMsg.toolCallId.trim() === "") {
          return { error: "tool message missing tool_call_id" };
        }
        results.push({
          type: "tool_result",
          tool_use_id: toolMsg.toolCallId,
          content: toolMsg.content,
        });
        i += 1;
      }
      out.push({ role: "user", content: results });
      continue;
    }

    i += 1;
  }

  const encodedTools: Record<string, unknown>[] = [];
  for (const tool of tools) {
    const external = maps.internalToExternal.get(tool.name);
    if (external == null) {
      return { error: "tool encoding failed for active tools map" };
    }
    encodedTools.push({
      name: external,
      description: tool.description,
      input_schema: tool.parameters,
    });
  }

  const result: AnthropicEncodeOk = { messages: out, tools: encodedTools };
  if (systemPrompt !== "") result.system = systemPrompt;
  return result;
}
