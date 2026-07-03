import { afterEach, describe, expect, it, vi } from "vitest";
import { desanitizeToolName, OpenAICompatibleClient, sanitizeToolName } from "./openaiCompatibleClient";

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
