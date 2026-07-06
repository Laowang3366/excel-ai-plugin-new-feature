import { decodeHtmlText as decodeHtml } from "../../shared/xmlEntities";

export interface WebSearchResultItem {
  title: string;
  url: string;
  snippet: string;
  source?: string;
}

export function parseDuckDuckGoHtml(html: string): WebSearchResultItem[] {
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

export function parseBingHtml(html: string): WebSearchResultItem[] {
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

export function parseBaiduHtml(html: string): WebSearchResultItem[] {
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

export function parseSoHtml(html: string): WebSearchResultItem[] {
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

export function parseSogouHtml(html: string): WebSearchResultItem[] {
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

export function cleanText(value: string): string {
  return value.replace(/\s+/g, " ").trim();
}

export function stripTags(value: string): string {
  return value.replace(/<[^>]+>/g, " ");
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
