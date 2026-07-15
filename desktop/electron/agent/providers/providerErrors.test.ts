import { describe, expect, it } from "vitest";

import { formatProviderHttpError } from "./providerErrors";

describe("provider error formatting", () => {
  it("summarizes Cloudflare HTML gateway errors", () => {
    const html = `<!DOCTYPE html><html><head><title>opencode.ai | 502: Bad gateway</title></head>
      <body><h1>Bad gateway</h1><span>Error code 502</span><div>cloudflare.com</div></body></html>`;

    expect(formatProviderHttpError("API 请求失败", 502, html)).toBe(
      "API 请求失败 (502): 模型服务网关暂时不可用（Cloudflare 502 Bad Gateway），请稍后重试或切换模型。",
    );
  });

  it("keeps concise JSON API error messages", () => {
    expect(
      formatProviderHttpError(
        "API 请求失败",
        429,
        JSON.stringify({
          error: { message: "rate limit" },
        }),
      ),
    ).toBe("API 请求失败 (429): rate limit");
  });
});
