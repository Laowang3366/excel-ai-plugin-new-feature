import { afterEach, describe, expect, it, vi } from "vitest";
import { desanitizeToolName, OpenAICompatibleClient, sanitizeToolName } from "./openaiCompatibleClient";

const baseConfig = {
  provider: "openai",
  apiKey: "test",
  baseUrl: "https://api.example.com/v1",
  model: "test-model",
};

async function collectStreamEvents(sse: string) {
  vi.stubGlobal("fetch", vi.fn(async () => new Response(sse, { status: 200 })));

  const client = new OpenAICompatibleClient(baseConfig);
  const events = [];
  for await (const event of client.streamChat({
    messages: [{ role: "user", content: "test" }],
  })) {
    events.push(event);
  }
  return events;
}

describe("tool name normalization", () => {
  it("restores Office tool names returned by APIs that reject dots", () => {
    expect(sanitizeToolName("word.open")).toBe("word_open");
    expect(desanitizeToolName("word_open")).toBe("word.open");
    expect(desanitizeToolName("presentation_setShapeText")).toBe("presentation.setShapeText");
    expect(desanitizeToolName("office_action_apply")).toBe("office.action.apply");
    expect(desanitizeToolName("office_script_execute")).toBe("office.script.execute");
    expect(desanitizeToolName("ocr_parseDocument")).toBe("ocr.parseDocument");
    expect(desanitizeToolName("knowledge_write")).toBe("knowledge.write");
    expect(desanitizeToolName("web_search")).toBe("web.search");
    expect(desanitizeToolName("python_execute")).toBe("python.execute");
    expect(desanitizeToolName("memory_search")).toBe("memory.search");
  });
});

describe("OpenAICompatibleClient errors", () => {
  afterEach(() => {
    vi.unstubAllGlobals();
  });

  it("does not expose raw HTML error pages to the chat stream", async () => {
    const html = `<!DOCTYPE html><html><head><title>opencode.ai | 502: Bad gateway</title></head>
      <body><h1>Bad gateway</h1><span>Error code 502</span><div>cloudflare.com</div></body></html>`;
    vi.stubGlobal("fetch", vi.fn(async () => new Response(html, { status: 502 })));

    const client = new OpenAICompatibleClient({
      provider: "openai",
      apiKey: "test",
      baseUrl: "https://inference.opencode.ai/v1",
      model: "deepseek-v4.flash",
    });
    const events = [];
    for await (const event of client.streamChat({
      messages: [{ role: "user", content: "测试" }],
    })) {
      events.push(event);
    }

    expect(events).toEqual([
      {
        type: "error",
        error: "API 请求失败 (502): 模型服务网关暂时不可用（Cloudflare 502 Bad Gateway），请稍后重试或切换模型。",
      },
    ]);
  });
});

describe("OpenAICompatibleClient streaming text", () => {
  afterEach(() => {
    vi.unstubAllGlobals();
  });

  it("emits text from array content parts", async () => {
    const events = await collectStreamEvents([
      "data: {\"choices\":[{\"delta\":{\"content\":[{\"type\":\"text\",\"text\":\"Hello\"},{\"type\":\"text\",\"text\":\" world\"}]}}]}",
      "data: {\"choices\":[{\"delta\":{},\"finish_reason\":\"stop\"}]}",
      "",
    ].join("\n"));

    expect(events.filter((event) => event.type === "text_delta")).toEqual([
      { type: "text_delta", delta: "Hello world" },
    ]);
  });

  it("emits final message.content when no text delta was streamed", async () => {
    const events = await collectStreamEvents([
      "data: {\"choices\":[{\"message\":{\"content\":[{\"type\":\"text\",\"text\":\"Final text\"}]},\"finish_reason\":\"stop\"}],\"usage\":{\"prompt_tokens\":2,\"completion_tokens\":3}}",
      "",
    ].join("\n"));

    expect(events.filter((event) => event.type === "text_delta")).toEqual([
      { type: "text_delta", delta: "Final text" },
    ]);
    expect(events[events.length - 1]).toEqual({ type: "done", finishReason: "stop" });
  });

  it("begins a tool call after both id and name are available", async () => {
    const events = await collectStreamEvents([
      "data: {\"choices\":[{\"delta\":{\"tool_calls\":[{\"index\":0,\"function\":{\"name\":\"range_read\"}}]}}]}",
      "data: {\"choices\":[{\"delta\":{\"tool_calls\":[{\"index\":0,\"id\":\"call_1\",\"function\":{\"arguments\":\"{\\\"range\\\":\\\"A1\\\"}\"}}]}}]}",
      "data: {\"choices\":[{\"delta\":{},\"finish_reason\":\"tool_calls\"}]}",
      "",
    ].join("\n"));

    expect(events.filter((event) => event.type.startsWith("tool_call"))).toEqual([
      { type: "tool_call_begin", toolCallId: "call_1", toolName: "range.read" },
      { type: "tool_call_delta", toolCallId: "call_1", delta: "{\"range\":\"A1\"}" },
      {
        type: "tool_call_end",
        toolCallId: "call_1",
        toolName: "range.read",
        arguments: "{\"range\":\"A1\"}",
      },
    ]);
  });
});
