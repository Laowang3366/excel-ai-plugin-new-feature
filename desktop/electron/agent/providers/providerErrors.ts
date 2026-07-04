import { decodeHtmlText } from "../shared/xmlEntities";

const HTML_TAG_RE = /<[^>]+>/g;

export function formatProviderHttpError(prefix: string, status: number, body: string): string {
  return `${prefix} (${status}): ${formatProviderErrorBody(status, body)}`;
}

export function formatProviderErrorBody(status: number, body: string): string {
  const trimmed = body.trim();
  if (!trimmed) return statusHint(status);

  const jsonMessage = extractJsonErrorMessage(trimmed);
  if (jsonMessage) return truncate(jsonMessage);

  if (looksLikeHtml(trimmed)) {
    const title = extractHtmlTitle(trimmed);
    const cloudflare = /cloudflare/i.test(trimmed);
    if (status === 502 || /bad gateway/i.test(trimmed)) {
      return cloudflare
        ? "模型服务网关暂时不可用（Cloudflare 502 Bad Gateway），请稍后重试或切换模型。"
        : "模型服务网关暂时不可用（502 Bad Gateway），请稍后重试或切换模型。";
    }
    if (title) return truncate(title);
    return statusHint(status);
  }

  return truncate(stripControlChars(trimmed));
}

function extractJsonErrorMessage(text: string): string | null {
  try {
    const data = JSON.parse(text);
    const message =
      data?.error?.message ||
      data?.error ||
      data?.message ||
      data?.msg ||
      data?.detail;
    return typeof message === "string" && message.trim() ? message.trim() : null;
  } catch {
    return null;
  }
}

function looksLikeHtml(text: string): boolean {
  return /^<!doctype html/i.test(text) || /^<html[\s>]/i.test(text) || /<body[\s>]/i.test(text);
}

function extractHtmlTitle(html: string): string | null {
  const match = html.match(/<title[^>]*>([\s\S]*?)<\/title>/i);
  if (!match) return null;
  const title = decodeHtmlText(match[1].replace(HTML_TAG_RE, " ").replace(/\s+/g, " ").trim());
  return title || null;
}

function statusHint(status: number): string {
  if (status === 429) return "请求过于频繁或额度受限，请稍后重试。";
  if (status === 502) return "模型服务网关暂时不可用（502 Bad Gateway），请稍后重试或切换模型。";
  if (status === 503) return "模型服务暂时不可用（503 Service Unavailable），请稍后重试或切换模型。";
  if (status === 504) return "模型服务响应超时（504 Gateway Timeout），请稍后重试或切换模型。";
  if (status >= 500) return "模型服务暂时异常，请稍后重试或切换模型。";
  return "请求失败，请检查模型配置或稍后重试。";
}

function stripControlChars(text: string): string {
  return text.replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f]+/g, " ");
}

function truncate(text: string, maxLength = 500): string {
  const compact = stripControlChars(text).replace(/\s+/g, " ").trim();
  return compact.length > maxLength ? `${compact.slice(0, maxLength)}...` : compact;
}
