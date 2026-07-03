import { afterEach, describe, expect, it, vi } from "vitest";
import { createAIClient } from "./aiClientFactory";
import { OpenAIResponsesClient } from "./openaiResponsesClient";

const baseConfig = {
  provider: "openai",
  apiKey: "test-key",
  baseUrl: "https://api.openai.com/v1",
  model: "gpt-test",
};

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
