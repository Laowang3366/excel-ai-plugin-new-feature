import {
  cleanText,
  parseBaiduHtml,
  parseBingHtml,
  parseDuckDuckGoHtml,
  parseSoHtml,
  parseSogouHtml,
  stripTags,
  type WebSearchResultItem,
} from "./webSearchHtmlParsers";

export interface WebSearchResponse {
  query: string;
  provider: string;
  results: WebSearchResultItem[];
}

export type SearchFreshness = "day" | "week" | "month" | "year" | "any";

export const MAX_SEARCH_RESPONSE_BYTES = 2 * 1024 * 1024;
const SEARCH_TIMEOUT_MS = 10_000;
const HTML_SEARCH_HEADERS = {
  "User-Agent":
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120 Safari/537.36",
  Accept: "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
  "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.7",
};

export function getSearchDestination(provider: string): string {
  switch (provider) {
    case "tavily":
      return "api.tavily.com";
    case "bing":
      return "api.bing.microsoft.com";
    case "serpapi":
      return "serpapi.com";
    case "bing-html":
      return "www.bing.com";
    case "baidu-html":
      return "www.baidu.com";
    case "so-html":
      return "www.so.com";
    case "sogou-html":
      return "www.sogou.com";
    case "duckduckgo":
      return "duckduckgo.com";
    default:
      return provider;
  }
}

export async function searchWeb(
  query: string,
  maxResults: number,
  freshness: SearchFreshness,
): Promise<WebSearchResponse> {
  const providers: Array<{ label: string; run: () => Promise<WebSearchResponse> }> = [];
  if (process.env.TAVILY_API_KEY) {
    providers.push({
      label: "Tavily",
      run: () => searchTavily(query, maxResults, process.env.TAVILY_API_KEY as string),
    });
  }
  if (process.env.BING_SEARCH_API_KEY) {
    providers.push({
      label: "Bing API",
      run: () =>
        searchBing(query, maxResults, freshness, process.env.BING_SEARCH_API_KEY as string),
    });
  }
  if (process.env.SERPAPI_API_KEY) {
    providers.push({
      label: "SerpAPI",
      run: () => searchSerpApi(query, maxResults, process.env.SERPAPI_API_KEY as string),
    });
  }

  providers.push(...buildFreeSearchProviders(query, maxResults));

  const failures: string[] = [];
  let firstEmptyResponse: WebSearchResponse | null = null;
  for (const provider of providers) {
    try {
      const response = await provider.run();
      if (response.results.length > 0) return response;
      firstEmptyResponse ||= response;
      failures.push(`${provider.label}: 未返回搜索结果`);
    } catch (error) {
      failures.push(`${provider.label}: ${formatThrownError(error)}`);
    }
  }

  if (firstEmptyResponse) return firstEmptyResponse;
  throw new Error(`所有搜索源均失败: ${failures.join("; ")}`);
}

export async function readSearchResponseText(
  response: Response,
  maxBytes = MAX_SEARCH_RESPONSE_BYTES,
): Promise<string> {
  const declaredLength = Number(response.headers?.get?.("content-length"));
  if (Number.isFinite(declaredLength) && declaredLength > maxBytes) {
    throw new Error(`搜索服务响应体超过限制 (${maxBytes} bytes)`);
  }

  const reader = response.body?.getReader();
  if (!reader) {
    const text = await response.text();
    if (Buffer.byteLength(text, "utf8") > maxBytes) {
      throw new Error(`搜索服务响应体超过限制 (${maxBytes} bytes)`);
    }
    return text;
  }

  const chunks: Uint8Array[] = [];
  let totalBytes = 0;
  try {
    while (true) {
      const { done, value } = await reader.read();
      if (done) break;
      totalBytes += value.byteLength;
      if (totalBytes > maxBytes) {
        await reader.cancel();
        throw new Error(`搜索服务响应体超过限制 (${maxBytes} bytes)`);
      }
      chunks.push(value);
    }
  } finally {
    reader.releaseLock();
  }
  return Buffer.concat(chunks.map((chunk) => Buffer.from(chunk))).toString("utf8");
}

function buildFreeSearchProviders(
  query: string,
  maxResults: number,
): Array<{ label: string; run: () => Promise<WebSearchResponse> }> {
  const chineseProviders = [
    { label: "360 Search HTML", run: () => searchSoHtml(query, maxResults) },
    { label: "Sogou HTML", run: () => searchSogouHtml(query, maxResults) },
    { label: "Baidu HTML", run: () => searchBaiduHtml(query, maxResults) },
    { label: "Bing HTML", run: () => searchBingHtml(query, maxResults) },
    { label: "DuckDuckGo HTML", run: () => searchDuckDuckGo(query, maxResults) },
  ];
  const globalProviders = [
    { label: "Bing HTML", run: () => searchBingHtml(query, maxResults) },
    { label: "DuckDuckGo HTML", run: () => searchDuckDuckGo(query, maxResults) },
    { label: "Sogou HTML", run: () => searchSogouHtml(query, maxResults) },
    { label: "Baidu HTML", run: () => searchBaiduHtml(query, maxResults) },
    { label: "360 Search HTML", run: () => searchSoHtml(query, maxResults) },
  ];
  return hasCjk(query) ? chineseProviders : globalProviders;
}

async function searchTavily(
  query: string,
  maxResults: number,
  apiKey: string,
): Promise<WebSearchResponse> {
  const response = await fetchWithTimeout("https://api.tavily.com/search", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      api_key: apiKey,
      query,
      max_results: maxResults,
      search_depth: "basic",
      include_answer: false,
      include_raw_content: false,
    }),
  });
  const json = await readJsonResponse(response);
  const results = Array.isArray(json.results) ? json.results : [];
  return {
    query,
    provider: "tavily",
    results: results
      .slice(0, maxResults)
      .map((item: Record<string, unknown>) => ({
        title: cleanText(String(item.title || item.url || "Untitled")),
        url: String(item.url || ""),
        snippet: cleanText(String(item.content || item.snippet || "")),
        source: "Tavily",
      }))
      .filter((item: WebSearchResultItem) => item.url),
  };
}

async function searchBing(
  query: string,
  maxResults: number,
  freshness: SearchFreshness,
  apiKey: string,
): Promise<WebSearchResponse> {
  const url = new URL("https://api.bing.microsoft.com/v7.0/search");
  url.searchParams.set("q", query);
  url.searchParams.set("count", String(maxResults));
  url.searchParams.set("responseFilter", "Webpages");
  const freshnessParam = bingFreshness(freshness);
  if (freshnessParam) url.searchParams.set("freshness", freshnessParam);

  const response = await fetchWithTimeout(url, {
    headers: { "Ocp-Apim-Subscription-Key": apiKey },
  });
  const json = await readJsonResponse(response);
  const webPages =
    json.webPages && typeof json.webPages === "object"
      ? (json.webPages as Record<string, unknown>)
      : {};
  const results = Array.isArray(webPages.value) ? webPages.value : [];
  return {
    query,
    provider: "bing",
    results: results
      .slice(0, maxResults)
      .map((item: Record<string, unknown>) => ({
        title: cleanText(String(item.name || item.url || "Untitled")),
        url: String(item.url || ""),
        snippet: cleanText(String(item.snippet || "")),
        source: "Bing",
      }))
      .filter((item: WebSearchResultItem) => item.url),
  };
}

async function searchSerpApi(
  query: string,
  maxResults: number,
  apiKey: string,
): Promise<WebSearchResponse> {
  const url = new URL("https://serpapi.com/search.json");
  url.searchParams.set("engine", "google");
  url.searchParams.set("q", query);
  url.searchParams.set("api_key", apiKey);
  url.searchParams.set("num", String(maxResults));

  const response = await fetchWithTimeout(url);
  const json = await readJsonResponse(response);
  const results = Array.isArray(json.organic_results) ? json.organic_results : [];
  return {
    query,
    provider: "serpapi",
    results: results
      .slice(0, maxResults)
      .map((item: Record<string, unknown>) => ({
        title: cleanText(String(item.title || item.link || "Untitled")),
        url: String(item.link || ""),
        snippet: cleanText(String(item.snippet || "")),
        source: "SerpAPI",
      }))
      .filter((item: WebSearchResultItem) => item.url),
  };
}

async function searchBingHtml(query: string, maxResults: number): Promise<WebSearchResponse> {
  const url = new URL("https://www.bing.com/search");
  url.searchParams.set("q", query);
  url.searchParams.set("count", String(maxResults));
  url.searchParams.set("ensearch", "0");
  return {
    query,
    provider: "bing-html",
    results: parseBingHtml(await fetchHtml(url)).slice(0, maxResults),
  };
}

async function searchBaiduHtml(query: string, maxResults: number): Promise<WebSearchResponse> {
  const url = new URL("https://www.baidu.com/s");
  url.searchParams.set("wd", query);
  url.searchParams.set("rn", String(maxResults));
  url.searchParams.set("ie", "utf-8");
  return {
    query,
    provider: "baidu-html",
    results: parseBaiduHtml(await fetchHtml(url)).slice(0, maxResults),
  };
}

async function searchSoHtml(query: string, maxResults: number): Promise<WebSearchResponse> {
  const url = new URL("https://www.so.com/s");
  url.searchParams.set("q", query);
  return {
    query,
    provider: "so-html",
    results: parseSoHtml(await fetchHtml(url)).slice(0, maxResults),
  };
}

async function searchSogouHtml(query: string, maxResults: number): Promise<WebSearchResponse> {
  const url = new URL("https://www.sogou.com/web");
  url.searchParams.set("query", query);
  return {
    query,
    provider: "sogou-html",
    results: parseSogouHtml(await fetchHtml(url)).slice(0, maxResults),
  };
}

async function searchDuckDuckGo(query: string, maxResults: number): Promise<WebSearchResponse> {
  const url = new URL("https://duckduckgo.com/html/");
  url.searchParams.set("q", query);
  return {
    query,
    provider: "duckduckgo",
    results: parseDuckDuckGoHtml(await fetchHtml(url)).slice(0, maxResults),
  };
}

async function fetchHtml(url: URL): Promise<string> {
  const response = await fetchWithTimeout(url, { headers: HTML_SEARCH_HEADERS });
  const text = await readSearchResponseText(response).catch((error) => {
    if (response.ok) throw error;
    return "";
  });
  if (!response.ok) {
    throw new Error(`请求失败 (${response.status}): ${formatErrorText(text)}`);
  }
  return text;
}

async function fetchWithTimeout(
  input: Parameters<typeof fetch>[0],
  init: Parameters<typeof fetch>[1] = {},
  timeoutMs = SEARCH_TIMEOUT_MS,
): Promise<Response> {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), timeoutMs);
  try {
    return await fetch(input, { ...init, signal: init?.signal || controller.signal });
  } catch (error: unknown) {
    if (error && typeof error === "object" && "name" in error && error.name === "AbortError") {
      throw new Error(`请求超时 (${timeoutMs}ms)`);
    }
    throw error;
  } finally {
    clearTimeout(timer);
  }
}

async function readJsonResponse(response: Response): Promise<Record<string, unknown>> {
  const text = await readSearchResponseText(response);
  if (!response.ok) {
    throw new Error(`搜索服务请求失败 (${response.status}): ${formatErrorText(text)}`);
  }
  try {
    const parsed: unknown = JSON.parse(text);
    if (!parsed || typeof parsed !== "object" || Array.isArray(parsed)) {
      throw new Error("not an object");
    }
    return parsed as Record<string, unknown>;
  } catch {
    throw new Error(`搜索服务返回非 JSON: ${formatErrorText(text)}`);
  }
}

function bingFreshness(freshness: SearchFreshness): string | undefined {
  switch (freshness) {
    case "day":
      return "Day";
    case "week":
      return "Week";
    case "month":
      return "Month";
    default:
      return undefined;
  }
}

function hasCjk(value: string): boolean {
  return /[\u3400-\u9fff]/.test(value);
}

function formatErrorText(value: string): string {
  return cleanText(stripTags(value)).slice(0, 300) || "空响应";
}

function formatThrownError(error: unknown): string {
  if (error instanceof Error) {
    const cause = (error as Error & { cause?: unknown }).cause;
    const causeText =
      cause && typeof cause === "object" && "code" in cause
        ? String(cause.code)
        : cause instanceof Error
          ? cause.message
          : "";
    return causeText ? `${error.message} (${causeText})` : error.message;
  }
  return String(error);
}
