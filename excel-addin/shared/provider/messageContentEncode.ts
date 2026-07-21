import type { AgentMessage } from "../agent/types";

/** OpenAI Chat Completions user/system content (string or multimodal array). */
export function encodeOpenAiChatContent(msg: AgentMessage): string | Record<string, unknown>[] {
  if (!msg.contentParts || msg.contentParts.length === 0) {
    return msg.content;
  }
  const parts: Record<string, unknown>[] = [];
  if (msg.content.trim()) {
    parts.push({ type: "text", text: msg.content });
  }
  for (const part of msg.contentParts) {
    if (part.type === "text") {
      parts.push({ type: "text", text: part.text });
    } else if (part.type === "image") {
      parts.push({
        type: "image_url",
        image_url: {
          url: `data:${part.mimeType};base64,${part.base64}`,
        },
      });
    }
  }
  return parts.length > 0 ? parts : msg.content;
}

/** OpenAI Responses API message content parts. */
export function encodeOpenAiResponsesContent(msg: AgentMessage): string | Record<string, unknown>[] {
  if (!msg.contentParts || msg.contentParts.length === 0) {
    return msg.content;
  }
  const parts: Record<string, unknown>[] = [];
  if (msg.content.trim()) {
    parts.push({ type: "input_text", text: msg.content });
  }
  for (const part of msg.contentParts) {
    if (part.type === "text") {
      parts.push({ type: "input_text", text: part.text });
    } else if (part.type === "image") {
      parts.push({
        type: "input_image",
        image_url: `data:${part.mimeType};base64,${part.base64}`,
      });
    }
  }
  return parts.length > 0 ? parts : msg.content;
}

/** Anthropic Messages API content blocks. */
export function encodeAnthropicContent(msg: AgentMessage): string | Record<string, unknown>[] {
  if (!msg.contentParts || msg.contentParts.length === 0) {
    return msg.content;
  }
  const parts: Record<string, unknown>[] = [];
  if (msg.content.trim()) {
    parts.push({ type: "text", text: msg.content });
  }
  for (const part of msg.contentParts) {
    if (part.type === "text") {
      parts.push({ type: "text", text: part.text });
    } else if (part.type === "image") {
      parts.push({
        type: "image",
        source: {
          type: "base64",
          media_type: part.mimeType,
          data: part.base64,
        },
      });
    }
  }
  return parts.length > 0 ? parts : msg.content;
}
