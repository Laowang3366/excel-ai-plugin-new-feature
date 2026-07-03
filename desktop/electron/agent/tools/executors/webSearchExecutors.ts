import type { ToolExecutor } from "../../shared/types";
import { validateArgs } from "./validation";

export interface WebSearchResultItem {
  title: string;
  url: string;
  snippet: string;
  source?: string;
}

export interface WebSearchResponse {
  query: string;
  provider: string;
  results: WebSearchResultItem[];
}

type Freshness = "day" | "week" | "month" | "year" | "any";

const SEARCH_TIMEOUT_MS = 10_000;
const HTML_SEARCH_HEADERS = {
  "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120 Safari/537.36",
  "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
  "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.7",
};

export function addWebSearchExecutors(target: Map<string, ToolExecutor>): void {
  target.set("web.search", {
    name: "web.search",
    execute: async (args: Record<string, unknown>) => {
      const err = validateArgs(args, { query: "string" });
      if (err) return { success: false, error: err };
      const optionalErr = validateOptionalSearchArgs(args);
      if (optionalErr) return { success: false, error: optionalErr };

      const query = (args.query as string).trim();
      if (!query) return { success: false, error: "参数 query 不能为空" };

      const maxResults = clampMaxResults(typeof args.maxResults === "number" ? args.maxResults : 5);
      const freshness = normalizeFreshness(args.freshness);

      try {
        const data = await searchWeb(query, maxResults, freshness);
        return { success: true, data };
      } catch (e: any) {
        return { success: false, error: `联网搜索失败: ${e.message}` };
      }
    },
  });
}

async function searchWeb(
  query: string,
  maxResults: number,
  freshness: Freshness
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
      run: () => searchBing(query, maxResults, freshness, process.env.BING_SEARCH_API_KEY as string),
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

  if (firstEmptyResponse) {
    return firstEmptyResponse;
  }
  throw new Error(`所有搜索源均失败: ${failures.join("; ")}`);
}

function buildFreeSearchProviders(
  query: string,
  maxResults: number
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
  apiKey: string
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
    results: results.slice(0, maxResults).map((item: any) => ({
      title: cleanText(item.title || item.url || "Untitled"),
      url: String(item.url || ""),
      snippet: cleanText(item.content || item.snippet || ""),
      source: "Tavily",
    })).filter((item: WebSearchResultItem) => item.url),
  };
}

async function searchBing(
  query: string,
  maxResults: number,
  freshness: Freshness,
  apiKey: string
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
  const results = Array.isArray(json.webPages?.value) ? json.webPages.value : [];
  return {
    query,
    provider: "bing",
    results: results.slice(0, maxResults).map((item: any) => ({
      title: cleanText(item.name || item.url || "Untitled"),
      url: String(item.url || ""),
      snippet: cleanText(item.snippet || ""),
      source: "Bing",
    })).filter((item: WebSearchResultItem) => item.url),
  };
}

async function searchSerpApi(
  query: string,
  maxResults: number,
  apiKey: string
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
    results: results.slice(0, maxResults).map((item: any) => ({
      title: cleanText(item.title || item.link || "Untitled"),
      url: String(item.link || ""),
      snippet: cleanText(item.snippet || ""),
      source: "SerpAPI",
    })).filter((item: WebSearchResultItem) => item.url),
  };
}

async function searchBingHtml(
  query: string,
  maxResults: number
): Promise<WebSearchResponse> {
  const url = new URL("https://www.bing.com/search");
  url.searchParams.set("q", query);
  url.searchParams.set("count", String(maxResults));
  url.searchParams.set("ensearch", "0");

  const html = await fetchHtml(url);
  return {
    query,
    provider: "bing-html",
    results: parseBingHtml(html).slice(0, maxResults),
  };
}

async function searchBaiduHtml(
  query: string,
  maxResults: number
): Promise<WebSearchResponse> {
  const url = new URL("https://www.baidu.com/s");
  url.searchParams.set("wd", query);
  url.searchParams.set("rn", String(maxResults));
  url.searchParams.set("ie", "utf-8");

  const html = await fetchHtml(url);
  return {
    query,
    provider: "baidu-html",
    results: parseBaiduHtml(html).slice(0, maxResults),
  };
}

async function searchSoHtml(
  query: string,
  maxResults: number
): Promise<WebSearchResponse> {
  const url = new URL("https://www.so.com/s");
  url.searchParams.set("q", query);

  const html = await fetchHtml(url);
  return {
    query,
    provider: "so-html",
    results: parseSoHtml(html).slice(0, maxResults),
  };
}

async function searchSogouHtml(
  query: string,
  maxResults: number
): Promise<WebSearchResponse> {
  const url = new URL("https://www.sogou.com/web");
  url.searchParams.set("query", query);

  const html = await fetchHtml(url);
  return {
    query,
    provider: "sogou-html",
    results: parseSogouHtml(html).slice(0, maxResults),
  };
}

async function searchDuckDuckGo(
  query: string,
  maxResults: number
): Promise<WebSearchResponse> {
  const url = new URL("https://duckduckgo.com/html/");
  url.searchParams.set("q", query);

  const html = await fetchHtml(url);
  const results = parseDuckDuckGoHtml(html).slice(0, maxResults);
  return {
    query,
    provider: "duckduckgo",
    results,
  };
}

function parseDuckDuckGoHtml(html: string): WebSearchResultItem[] {
  const results: WebSearchResultItem[] = [];
  const resultRe = /<div[^>]+class="[^"]*\bresult\b[^"]*"[\s\S]*?<\/div>\s*<\/div>/gi;
  let match: RegExpExecArray | null;
  while ((match = resultRe.exec(html))) {
    const block = match[0];
    const link = /<a[^>]+class="[^"]*\bresult__a\b[^"]*"[^>]+href="([^"]+)"[^>]*>([\s\S]*?)<\/a>/i.exec(block);
    if (!link) continue;
    const snippet = /<a[^>]+class="[^"]*\bresult__snippet\b[^"]*"[^>]*>([\s\S]*?)<\/a>/i.exec(block)
      || /<div[^>]+class="[^"]*\bresult__snippet\b[^"]*"[^>]*>([\s\S]*?)<\/div>/i.exec(block);
    const url = normalizeDuckDuckGoUrl(decodeHtml(link[1]));
    if (!url || !/^https?:\/\//i.test(url)) continue;
    results.push({
      title: cleanText(decodeHtml(stripTags(link[2]))),
      url,
      snippet: cleanText(decodeHtml(stripTags(snippet?.[1] || ""))),
      source: hostnameFromUrl(url),
    });
  }
  return dedupeResults(results);
}

function parseBingHtml(html: string): WebSearchResultItem[] {
  const results: WebSearchResultItem[] = [];
  const resultRe = /<li[^>]+class=["'][^"']*\bb_algo\b[^"']*["'][^>]*>[\s\S]*?<\/li>/gi;
  let match: RegExpExecArray | null;
  while ((match = resultRe.exec(html))) {
    const block = match[0];
    const link = /<h2[^>]*>[\s\S]*?<a[^>]+href=["']([^"']+)["'][^>]*>([\s\S]*?)<\/a>[\s\S]*?<\/h2>/i.exec(block)
      || /<a[^>]+href=["']([^"']+)["'][^>]*>([\s\S]*?)<\/a>/i.exec(block);
    if (!link) continue;
    const url = normalizeBingUrl(decodeHtml(link[1]));
    if (!url || !/^https?:\/\//i.test(url)) continue;
    const snippet = /<p[^>]*>([\s\S]*?)<\/p>/i.exec(block)
      || /<div[^>]+class=["'][^"']*\bb_caption\b[^"']*["'][^>]*>([\s\S]*?)<\/div>/i.exec(block);
    results.push({
      title: cleanText(decodeHtml(stripTags(link[2]))),
      url,
      snippet: cleanText(decodeHtml(stripTags(snippet?.[1] || ""))),
      source: hostnameFromUrl(url),
    });
  }
  return dedupeResults(results);
}

function parseBaiduHtml(html: string): WebSearchResultItem[] {
  const results: WebSearchResultItem[] = [];
  const resultRe = /<div[^>]+class=["'][^"']*(?:result|c-container)[^"']*["'][^>]*>[\s\S]*?(?=<div[^>]+class=["'][^"']*(?:result|c-container)|<\/body>|$)/gi;
  let match: RegExpExecArray | null;
  while ((match = resultRe.exec(html))) {
    const block = match[0];
    const link = /<h3[^>]*>[\s\S]*?<a[^>]+href=["']([^"']+)["'][^>]*>([\s\S]*?)<\/a>[\s\S]*?<\/h3>/i.exec(block)
      || /<a[^>]+href=["']([^"']+)["'][^>]*>([\s\S]*?)<\/a>/i.exec(block);
    if (!link) continue;
    const directUrl = /\bmu=["']([^"']+)["']/i.exec(block)?.[1];
    const url = decodeHtml(directUrl || link[1]);
    if (!url || !/^https?:\/\//i.test(url)) continue;
    const snippet = /<div[^>]+class=["'][^"']*(?:c-abstract|content-right|c-span-last)[^"']*["'][^>]*>([\s\S]*?)<\/div>/i.exec(block)
      || /<span[^>]+class=["'][^"']*(?:content-right|c-color-text)[^"']*["'][^>]*>([\s\S]*?)<\/span>/i.exec(block);
    results.push({
      title: cleanText(decodeHtml(stripTags(link[2]))),
      url,
      snippet: cleanText(decodeHtml(stripTags(snippet?.[1] || ""))),
      source: hostnameFromUrl(url),
    });
  }
  return dedupeResults(results);
}

function parseSoHtml(html: string): WebSearchResultItem[] {
  const results: WebSearchResultItem[] = [];
  const resultRe = /<li[^>]+class=["'][^"']*\bres-list\b[^"']*["'][^>]*>[\s\S]*?(?=<li[^>]+class=["'][^"']*\bres-list\b|<\/ul>|$)/gi;
  let match: RegExpExecArray | null;
  while ((match = resultRe.exec(html))) {
    const block = match[0];
    const link = /<h3[^>]+class=["'][^"']*\bres-title\b[^"']*["'][^>]*>[\s\S]*?<a[^>]+href=["']([^"']+)["'][^>]*>([\s\S]*?)<\/a>[\s\S]*?<\/h3>/i.exec(block)
      || /<h3[^>]*>[\s\S]*?<a[^>]+href=["']([^"']+)["'][^>]*>([\s\S]*?)<\/a>[\s\S]*?<\/h3>/i.exec(block);
    if (!link) continue;
    const mdUrl = /\bdata-mdurl=["']([^"']+)["']/i.exec(link[0] || block)?.[1];
    const url = decodeHtml(mdUrl || link[1]);
    if (!url || !/^https?:\/\//i.test(url)) continue;
    const snippet = /<p[^>]+class=["'][^"']*\bres-desc\b[^"']*["'][^>]*>([\s\S]*?)<\/p>/i.exec(block)
      || /<span[^>]+class=["'][^"']*\bres-list-summary\b[^"']*["'][^>]*>([\s\S]*?)<\/span>/i.exec(block)
      || /<div[^>]+class=["'][^"']*\bres-rich\b[^"']*["'][^>]*>([\s\S]*?)<\/div>/i.exec(block);
    results.push({
      title: cleanText(decodeHtml(stripTags(link[2]))),
      url,
      snippet: cleanText(decodeHtml(stripTags(snippet?.[1] || ""))),
      source: hostnameFromUrl(url),
    });
  }
  return dedupeResults(results);
}

function parseSogouHtml(html: string): WebSearchResultItem[] {
  const results: WebSearchResultItem[] = [];
  const resultRe = /<div[^>]+class=["'][^"']*\bvrwrap\b[^"']*["'][^>]*>[\s\S]*?(?=<div[^>]+class=["'][^"']*\bvrwrap\b|<\/body>|$)/gi;
  let match: RegExpExecArray | null;
  while ((match = resultRe.exec(html))) {
    const block = match[0];
    const link = /<h3[^>]+class=["'][^"']*(?:vr-title|vrTitle)[^"']*["'][^>]*>[\s\S]*?<a[^>]+href=["']([^"']+)["'][^>]*>([\s\S]*?)<\/a>[\s\S]*?<\/h3>/i.exec(block)
      || /<h3[^>]*>[\s\S]*?<a[^>]+href=["']([^"']+)["'][^>]*>([\s\S]*?)<\/a>[\s\S]*?<\/h3>/i.exec(block);
    if (!link) continue;
    const directUrl = /\bdata-url=["']([^"']+)["']/i.exec(block)?.[1];
    const url = normalizeSearchResultUrl(decodeHtml(directUrl || link[1]), "https://www.sogou.com");
    if (!url || !/^https?:\/\//i.test(url)) continue;
    const snippet = /<div[^>]+class=["'][^"']*(?:fz-mid|space-txt|base-ellipsis|text-layout)[^"']*["'][^>]*>([\s\S]*?)<\/div>/i.exec(block)
      || /<p[^>]+class=["'][^"']*(?:str_info|txt-info)[^"']*["'][^>]*>([\s\S]*?)<\/p>/i.exec(block);
    results.push({
      title: cleanText(decodeHtml(stripTags(link[2]))),
      url,
      snippet: cleanText(decodeHtml(stripTags(snippet?.[1] || ""))),
      source: hostnameFromUrl(url),
    });
  }
  return dedupeResults(results);
}

async function fetchHtml(url: URL): Promise<string> {
  const response = await fetchWithTimeout(url, { headers: HTML_SEARCH_HEADERS });
  const text = await response.text().catch(() => "");
  if (!response.ok) {
    throw new Error(`请求失败 (${response.status}): ${formatErrorText(text)}`);
  }
  return text;
}

async function fetchWithTimeout(
  input: Parameters<typeof fetch>[0],
  init: Parameters<typeof fetch>[1] = {},
  timeoutMs = SEARCH_TIMEOUT_MS
): Promise<Response> {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), timeoutMs);
  try {
    return await fetch(input, {
      ...init,
      signal: init?.signal || controller.signal,
    });
  } catch (error) {
    if ((error as any)?.name === "AbortError") {
      throw new Error(`请求超时 (${timeoutMs}ms)`);
    }
    throw error;
  } finally {
    clearTimeout(timer);
  }
}

async function readJsonResponse(response: Response): Promise<any> {
  const text = await response.text();
  if (!response.ok) {
    throw new Error(`搜索服务请求失败 (${response.status}): ${formatErrorText(text)}`);
  }
  try {
    return JSON.parse(text);
  } catch {
    throw new Error(`搜索服务返回非 JSON: ${formatErrorText(text)}`);
  }
}

function validateOptionalSearchArgs(args: Record<string, unknown>): string | null {
  if (args.maxResults !== undefined && typeof args.maxResults !== "number") {
    return "参数 maxResults 必须是 number";
  }
  if (args.freshness !== undefined && typeof args.freshness !== "string") {
    return "参数 freshness 必须是 string";
  }
  const freshness = normalizeFreshness(args.freshness);
  if (args.freshness !== undefined && freshness === "any" && args.freshness !== "any") {
    return "参数 freshness 必须是 day、week、month、year 或 any";
  }
  return null;
}

function clampMaxResults(value: number): number {
  if (!Number.isFinite(value)) return 5;
  return Math.max(1, Math.min(10, Math.floor(value)));
}

function normalizeFreshness(value: unknown): Freshness {
  return value === "day" || value === "week" || value === "month" || value === "year" || value === "any"
    ? value
    : "any";
}

function bingFreshness(freshness: Freshness): string | undefined {
  switch (freshness) {
    case "day": return "Day";
    case "week": return "Week";
    case "month": return "Month";
    default: return undefined;
  }
}

function normalizeDuckDuckGoUrl(url: string): string {
  const normalized = url.startsWith("//") ? `https:${url}` : url;
  if (!normalized.includes("duckduckgo.com/l/")) return normalized;
  try {
    const parsed = new URL(normalized);
    const target = parsed.searchParams.get("uddg");
    return target ? decodeURIComponent(target) : normalized;
  } catch {
    return normalized;
  }
}

function normalizeBingUrl(url: string): string {
  const normalized = url.startsWith("//") ? `https:${url}` : url;
  try {
    const parsed = new URL(normalized);
    if (parsed.hostname.endsWith("bing.com") && parsed.pathname === "/ck/a") {
      const target = parsed.searchParams.get("u");
      if (target) return decodeURIComponent(target.replace(/^a1/i, ""));
    }
  } catch {
    return normalized;
  }
  return normalized;
}

function normalizeSearchResultUrl(url: string, baseUrl: string): string {
  try {
    return new URL(url, baseUrl).toString();
  } catch {
    return url;
  }
}

function dedupeResults(results: WebSearchResultItem[]): WebSearchResultItem[] {
  const seen = new Set<string>();
  const deduped: WebSearchResultItem[] = [];
  for (const item of results) {
    const key = item.url.replace(/\/+$/, "");
    if (seen.has(key)) continue;
    seen.add(key);
    deduped.push(item);
  }
  return deduped;
}

function hostnameFromUrl(url: string): string | undefined {
  try {
    return new URL(url).hostname.replace(/^www\./i, "");
  } catch {
    return undefined;
  }
}

function stripTags(value: string): string {
  return value.replace(/<[^>]+>/g, " ");
}

function decodeHtml(value: string): string {
  return value
    .replace(/&quot;/g, "\"")
    .replace(/&#x27;/g, "'")
    .replace(/&#39;/g, "'")
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&nbsp;/g, " ");
}

function cleanText(value: string): string {
  return value.replace(/\s+/g, " ").trim();
}

function hasCjk(value: string): boolean {
  return /[\u3400-\u9fff]/.test(value);
}

function formatErrorText(value: string): string {
  return cleanText(stripTags(value)).slice(0, 300) || "空响应";
}

function formatThrownError(error: unknown): string {
  if (error instanceof Error) {
    const cause = (error as any).cause;
    const causeText = cause?.code || cause?.message;
    return causeText ? `${error.message} (${causeText})` : error.message;
  }
  return String(error);
}
