import { describe, expect, it, vi } from "vitest";

import { MAX_SEARCH_RESPONSE_BYTES, readSearchResponseText } from "./webSearchProviders";

describe("web search provider responses", () => {
  it("reads a response body within the configured limit", async () => {
    const response = new Response("搜索结果");

    await expect(readSearchResponseText(response)).resolves.toBe("搜索结果");
  });

  it("rejects an oversized declared content length before reading the body", async () => {
    const text = vi.fn(async () => "not read");
    const response = {
      headers: new Headers({
        "content-length": String(MAX_SEARCH_RESPONSE_BYTES + 1),
      }),
      body: null,
      text,
    } as unknown as Response;

    await expect(readSearchResponseText(response)).rejects.toThrow(
      `搜索服务响应体超过限制 (${MAX_SEARCH_RESPONSE_BYTES} bytes)`,
    );
    expect(text).not.toHaveBeenCalled();
  });

  it("rejects a streamed body that grows beyond two MiB", async () => {
    const response = new Response(new Uint8Array(MAX_SEARCH_RESPONSE_BYTES + 1));

    await expect(readSearchResponseText(response)).rejects.toThrow(
      `搜索服务响应体超过限制 (${MAX_SEARCH_RESPONSE_BYTES} bytes)`,
    );
  });
});
