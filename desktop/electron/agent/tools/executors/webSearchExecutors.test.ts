import { afterEach, describe, expect, it, vi } from "vitest";

import { addWebSearchExecutors } from "./webSearchExecutors";

describe("web search executors", () => {
  afterEach(() => {
    vi.restoreAllMocks();
    delete process.env.TAVILY_API_KEY;
    delete process.env.BING_SEARCH_API_KEY;
    delete process.env.SERPAPI_API_KEY;
  });

  it("searches Bing HTML when no paid search provider is configured", async () => {
    const html = `
      <li class="b_algo">
        <h2><a href="https://example.com/doc">Example &amp; Docs</a></h2>
        <p>A useful result &amp; summary.</p>
      </li>
    `;
    vi.stubGlobal("fetch", vi.fn(async () => ({
      ok: true,
      text: async () => html,
    })));
    const executors = new Map();
    addWebSearchExecutors(executors, { isRemoteDataProcessingEnabled: () => true });

    const result = await executors.get("web.search").execute({
      query: "excel dynamic array",
      maxResults: 3,
    });

    expect(result).toEqual({
      success: true,
      data: {
        query: "excel dynamic array",
        provider: "bing-html",
        results: [
          {
            title: "Example & Docs",
            url: "https://example.com/doc",
            snippet: "A useful result & summary.",
            source: "example.com",
          },
          ],
          remoteProcessing: {
            operation: "web-search",
            service: "bing-html",
            destination: "www.bing.com",
            dataSummary: "搜索查询，19 个字符",
          },
      },
    });
  });

  it("falls back through global HTML providers to Baidu HTML", async () => {
    const baiduHtml = `
      <div class="result c-container" mu="https://example.cn/public-info">
        <h3><a href="https://www.baidu.com/link?url=abc">Public information</a></h3>
        <div class="c-abstract">A public web result.</div>
      </div>
    `;
    const fetchMock = vi.fn()
      .mockRejectedValueOnce(new Error("fetch failed"))
      .mockRejectedValueOnce(new Error("fetch failed"))
      .mockRejectedValueOnce(new Error("fetch failed"))
      .mockResolvedValueOnce({
        ok: true,
        text: async () => baiduHtml,
      });
    vi.stubGlobal("fetch", fetchMock);
    const executors = new Map();
    addWebSearchExecutors(executors, { isRemoteDataProcessingEnabled: () => true });

    const result = await executors.get("web.search").execute({
      query: "public information",
      maxResults: 5,
    });

    expect(fetchMock).toHaveBeenCalledTimes(4);
    expect(result).toEqual({
      success: true,
      data: {
        query: "public information",
        provider: "baidu-html",
        results: [
          {
            title: "Public information",
            url: "https://example.cn/public-info",
            snippet: "A public web result.",
            source: "example.cn",
          },
          ],
          remoteProcessing: {
            operation: "web-search",
            service: "baidu-html",
            destination: "www.baidu.com",
            dataSummary: "搜索查询，18 个字符",
          },
      },
    });
  });

  it("uses Chinese search providers first for Chinese queries", async () => {
    const html = `
      <li class="res-list">
        <h3 class="res-title">
          <a href="https://www.so.com/link?m=abc" data-mdurl="https://example.cn/public-info">广东佛山公开资料</a>
        </h3>
        <p class="res-desc">广东佛山相关公开网页资料。</p>
      </li>
    `;
    const fetchMock = vi.fn(async (_input: unknown, _init?: unknown) => ({
      ok: true,
      text: async () => html,
    }));
    vi.stubGlobal("fetch", fetchMock);
    const executors = new Map();
    addWebSearchExecutors(executors, { isRemoteDataProcessingEnabled: () => true });

    const result = await executors.get("web.search").execute({
      query: "广东佛山 公开资料",
      maxResults: 5,
    });

    expect(fetchMock.mock.calls).toHaveLength(1);
    expect(String(fetchMock.mock.calls[0]?.[0])).toContain("www.so.com");
    expect(result).toEqual({
      success: true,
      data: {
        query: "广东佛山 公开资料",
        provider: "so-html",
        results: [
          {
            title: "广东佛山公开资料",
            url: "https://example.cn/public-info",
            snippet: "广东佛山相关公开网页资料。",
            source: "example.cn",
          },
          ],
          remoteProcessing: {
            operation: "web-search",
            service: "so-html",
            destination: "www.so.com",
            dataSummary: "搜索查询，9 个字符",
          },
      },
    });
  });

  it("returns a diagnostic error when all search providers fail", async () => {
    vi.stubGlobal("fetch", vi.fn(async () => ({
      ok: false,
      status: 503,
      text: async () => "Service unavailable",
    })));
    const executors = new Map();
    addWebSearchExecutors(executors, { isRemoteDataProcessingEnabled: () => true });

    const result = await executors.get("web.search").execute({
      query: "demo",
    });

    expect(result.success).toBe(false);
    expect(result.error).toContain("联网搜索失败: 所有搜索源均失败");
    expect(result.error).toContain("360 Search HTML: 请求失败 (503)");
    expect(result.error).toContain("Sogou HTML: 请求失败 (503)");
    expect(result.error).toContain("Bing HTML: 请求失败 (503)");
    expect(result.error).toContain("Baidu HTML: 请求失败 (503)");
    expect(result.error).toContain("DuckDuckGo HTML: 请求失败 (503)");
  });

  it("rejects invalid freshness values", async () => {
    const executors = new Map();
    addWebSearchExecutors(executors);

    const result = await executors.get("web.search").execute({
      query: "demo",
      freshness: "hour",
    });

    expect(result).toEqual({
      success: false,
      error: "参数 freshness 必须是 day、week、month、year 或 any",
    });
  });

  it("does not make a request when remote processing is disabled", async () => {
    const fetchMock = vi.fn();
    vi.stubGlobal("fetch", fetchMock);
    const executors = new Map();
    addWebSearchExecutors(executors, { isRemoteDataProcessingEnabled: () => false });

    const result = await executors.get("web.search").execute({ query: "public information" });

    expect(fetchMock).not.toHaveBeenCalled();
    expect(result).toMatchObject({
      success: false,
      data: { code: "remote_data_processing_disabled", operation: "web-search" },
    });
  });

  it("blocks a high-confidence secret before any search request", async () => {
    const fetchMock = vi.fn();
    vi.stubGlobal("fetch", fetchMock);
    const executors = new Map();
    addWebSearchExecutors(executors, { isRemoteDataProcessingEnabled: () => true });

    const result = await executors.get("web.search").execute({
      query: "find sk-1234567890abcdefghijklmnop",
    });

    expect(fetchMock).not.toHaveBeenCalled();
    expect(result).toMatchObject({
      success: false,
      data: { code: "sensitive_data_detected", operation: "web-search" },
    });
  });
});
