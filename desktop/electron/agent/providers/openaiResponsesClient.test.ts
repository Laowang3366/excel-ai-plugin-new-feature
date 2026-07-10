import { afterEach, describe, expect, it, vi } from "vitest";
import { createAIClient } from "./aiClientFactory";
import { OpenAIResponsesClient } from "./openaiResponsesClient";

const baseConfig = {
  provider: "openai",
  apiKey: "test-key",
  baseUrl: "https://api.openai.com/v1",
  model: "gpt-test",
};

async function collectStreamEvents(client: OpenAIResponsesClient, sse: string) {
  // @MOCK_INTERFACE: simulates the OpenAI Responses SSE fetch stream consumed by OpenAIResponsesClient.
  const fetchMock = vi.fn(async () => new Response(sse, { status: 200 }));
  vi.stubGlobal("fetch", fetchMock);

  const events = [];
  for await (const event of client.streamChat({
    messages: [{ role: "user", content: "test" }],
  })) {
    events.push(event);
  }
  return { events, fetchMock };
}

describe("OpenAIResponsesClient", () => {
  afterEach(() => {
    vi.unstubAllGlobals();
  });

  it("is selected when apiFormat is responses", () => {
    const client = createAIClient({ ...baseConfig, apiFormat: "responses" });
    expect(client).toBeInstanceOf(OpenAIResponsesClient);
  });

  it("posts PDF and image inputs to the Responses endpoint", async () => {
    const fetchMock = vi.fn(async (_url: string, _init: RequestInit) =>
      new Response(JSON.stringify({
        output_text: "识别完成",
        usage: { input_tokens: 11, output_tokens: 7 },
      }), { status: 200 })
    );
    vi.stubGlobal("fetch", fetchMock);

    const client = new OpenAIResponsesClient({ ...baseConfig, apiFormat: "responses" });
    const result = await client.chat({
      systemPrompt: "只返回 JSON",
      messages: [
        {
          role: "user",
          content: [
            { type: "text", text: "识别这个文件" },
            {
              type: "file",
              file: {
                filename: "invoice.pdf",
                mime_type: "application/pdf",
                file_data: "data:application/pdf;base64,AAAA",
              },
            },
            {
              type: "image_url",
              image_url: {
                url: "data:image/png;base64,BBBB",
                detail: "high",
              },
            },
          ],
        },
      ],
      maxTokens: 3000,
      temperature: 0,
      reasoningMode: "off",
    });

    expect(result.content).toBe("识别完成");
    expect(fetchMock).toHaveBeenCalledTimes(1);
    expect(fetchMock.mock.calls[0][0]).toBe("https://api.openai.com/v1/responses");

    const init = fetchMock.mock.calls[0][1] as RequestInit;
    const body = JSON.parse(init.body as string);
    expect(body).toMatchObject({
      model: "gpt-test",
      stream: false,
      instructions: "只返回 JSON",
      max_output_tokens: 3000,
      temperature: 0,
    });
    expect(body.reasoning).toEqual({ effort: "none" });
    expect(body.input[0].content).toEqual([
      { type: "input_text", text: "识别这个文件" },
      {
        type: "input_file",
        filename: "invoice.pdf",
        file_data: "data:application/pdf;base64,AAAA",
      },
      {
        type: "input_image",
        image_url: "data:image/png;base64,BBBB",
        detail: "high",
      },
    ]);
  });

  it("rejects non-stream chat when the Responses endpoint returns an HTTP failure", async () => {
    vi.stubGlobal("fetch", vi.fn(async () =>
      new Response("service unavailable", { status: 503 })
    ));

    const client = new OpenAIResponsesClient({ ...baseConfig, apiFormat: "responses" });

    await expect(client.chat({
      messages: [{ role: "user", content: "生成摘要" }],
    })).rejects.toThrow("Responses API 请求失败 (503)");
  });

  it("sends OpenAI reasoning effort and visible summary request when enabled", async () => {
    const fetchMock = vi.fn(async (_url: string, _init: RequestInit) =>
      new Response(JSON.stringify({
        output_text: "done",
        usage: { input_tokens: 1, output_tokens: 1 },
      }), { status: 200 })
    );
    vi.stubGlobal("fetch", fetchMock);

    const client = new OpenAIResponsesClient({ ...baseConfig, apiFormat: "responses" });
    await client.chat({
      messages: [{ role: "user", content: "hi" }],
      reasoningMode: "max",
    });

    const init = fetchMock.mock.calls[0][1] as RequestInit;
    const body = JSON.parse(init.body as string);
    expect(body.reasoning).toEqual({ effort: "xhigh", summary: "auto" });
  });

  it("emits missing text from output_text.done without duplicating prior deltas", async () => {
    const sse = [
      "event: response.output_text.delta\ndata: {\"type\":\"response.output_text.delta\",\"item_id\":\"msg_1\",\"content_index\":0,\"delta\":\"Hello\"}",
      "event: response.output_text.done\ndata: {\"type\":\"response.output_text.done\",\"item_id\":\"msg_1\",\"content_index\":0,\"text\":\"Hello world\"}",
      "event: response.completed\ndata: {\"type\":\"response.completed\",\"response\":{\"output_text\":\"Hello world\",\"usage\":{\"input_tokens\":2,\"output_tokens\":3}}}",
      "",
    ].join("\n\n");

    const client = new OpenAIResponsesClient({ ...baseConfig, apiFormat: "responses" });
    const { events } = await collectStreamEvents(client, sse);

    expect(events.filter((event) => event.type === "text_delta")).toEqual([
      { type: "text_delta", delta: "Hello" },
      { type: "text_delta", delta: " world" },
    ]);
  });

  it("emits text from content_part.done when no delta event was streamed", async () => {
    const sse = [
      "event: response.content_part.done\ndata: {\"type\":\"response.content_part.done\",\"item_id\":\"msg_1\",\"content_index\":0,\"part\":{\"type\":\"output_text\",\"text\":\"Part text\"}}",
      "event: response.completed\ndata: {\"type\":\"response.completed\",\"response\":{\"output\":[{\"type\":\"message\",\"id\":\"msg_1\",\"content\":[{\"type\":\"output_text\",\"text\":\"Part text\"}]}]}}",
      "",
    ].join("\n\n");

    const client = new OpenAIResponsesClient({ ...baseConfig, apiFormat: "responses" });
    const { events } = await collectStreamEvents(client, sse);

    expect(events.filter((event) => event.type === "text_delta")).toEqual([
      { type: "text_delta", delta: "Part text" },
    ]);
  });

  it("emits text from message output_item.done when no text delta was streamed", async () => {
    const sse = [
      "event: response.output_item.done\ndata: {\"type\":\"response.output_item.done\",\"item\":{\"type\":\"message\",\"id\":\"msg_1\",\"content\":[{\"type\":\"output_text\",\"text\":\"Item text\"}]}}",
      "event: response.completed\ndata: {\"type\":\"response.completed\",\"response\":{\"output\":[{\"type\":\"message\",\"id\":\"msg_1\",\"content\":[{\"type\":\"output_text\",\"text\":\"Item text\"}]}]}}",
      "",
    ].join("\n\n");

    const client = new OpenAIResponsesClient({ ...baseConfig, apiFormat: "responses" });
    const { events } = await collectStreamEvents(client, sse);

    expect(events.filter((event) => event.type === "text_delta")).toEqual([
      { type: "text_delta", delta: "Item text" },
    ]);
  });

  it("emits final response text from response.completed as a last fallback", async () => {
    const sse = [
      "event: response.completed\ndata: {\"type\":\"response.completed\",\"response\":{\"output\":[{\"type\":\"message\",\"id\":\"msg_1\",\"content\":[{\"type\":\"output_text\",\"text\":\"Completed text\"}]}],\"usage\":{\"input_tokens\":2,\"output_tokens\":2}}}",
      "",
    ].join("\n\n");

    const client = new OpenAIResponsesClient({ ...baseConfig, apiFormat: "responses" });
    const { events } = await collectStreamEvents(client, sse);

    expect(events.slice(0, 1)).toEqual([
      { type: "text_delta", delta: "Completed text" },
    ]);
    expect(events[events.length - 1]).toEqual({ type: "done", finishReason: "stop" });
  });

  it("flushes the final SSE event even without a trailing blank line", async () => {
    const sse = "event: response.completed\ndata: {\"type\":\"response.completed\",\"response\":{\"output_text\":\"No trailing newline\",\"usage\":{\"input_tokens\":2,\"output_tokens\":4}}}";

    const client = new OpenAIResponsesClient({ ...baseConfig, apiFormat: "responses" });
    const { events } = await collectStreamEvents(client, sse);

    expect(events).toEqual([
      { type: "text_delta", delta: "No trailing newline" },
      {
        type: "usage",
        usage: {
          inputTokens: 2,
          outputTokens: 4,
          cachedInputTokens: 0,
          reasoningOutputTokens: 0,
        },
      },
      { type: "done", finishReason: "stop" },
    ]);
  });

  it("parses Responses streaming function calls into internal tool events", async () => {
    const sse = [
      "event: response.output_item.added\ndata: {\"type\":\"response.output_item.added\",\"item\":{\"type\":\"function_call\",\"id\":\"fc_1\",\"call_id\":\"call_1\",\"name\":\"range_read\"}}",
      "event: response.function_call_arguments.delta\ndata: {\"type\":\"response.function_call_arguments.delta\",\"item_id\":\"fc_1\",\"delta\":\"{\\\"range\\\":\\\"A1\\\"}\"}",
      "event: response.output_item.done\ndata: {\"type\":\"response.output_item.done\",\"item\":{\"type\":\"function_call\",\"id\":\"fc_1\",\"call_id\":\"call_1\",\"name\":\"range_read\",\"arguments\":\"{\\\"range\\\":\\\"A1\\\"}\"}}",
      "event: response.completed\ndata: {\"type\":\"response.completed\",\"response\":{\"output\":[{\"type\":\"function_call\"}],\"usage\":{\"input_tokens\":5,\"output_tokens\":3}}}",
      "",
    ].join("\n\n");

    const fetchMock = vi.fn(async () => new Response(sse, { status: 200 }));
    vi.stubGlobal("fetch", fetchMock);

    const client = new OpenAIResponsesClient({ ...baseConfig, apiFormat: "responses" });
    const events = [];
    for await (const event of client.streamChat({
      messages: [{ role: "user", content: "读取 A1" }],
      tools: [
        {
          name: "range.read",
          description: "读取区域",
          parameters: { type: "object", properties: {} },
          riskLevel: "safe",
          requiresApproval: false,
        },
      ],
    })) {
      events.push(event);
    }

    expect(events).toEqual([
      { type: "tool_call_begin", toolCallId: "call_1", toolName: "range.read" },
      { type: "tool_call_delta", toolCallId: "call_1", delta: "{\"range\":\"A1\"}" },
      {
        type: "tool_call_end",
        toolCallId: "call_1",
        toolName: "range.read",
        arguments: "{\"range\":\"A1\"}",
      },
      {
        type: "usage",
        usage: {
          inputTokens: 5,
          outputTokens: 3,
          cachedInputTokens: 0,
          reasoningOutputTokens: 0,
        },
      },
      { type: "done", finishReason: "tool_calls" },
    ]);
  });
});
