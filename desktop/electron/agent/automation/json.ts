/**
 * 自动化脚本 JSON 解析
 *
 * 被 COM 桥接层用于解析 Python、JScript、PowerShell 返回的 JSON 输出。
 */

import type { ScriptEngine } from "./scriptEngine";

/**
 * 安全解析脚本引擎返回的 JSON 字符串。
 */
export function safeJsonParse<T = unknown>(
  raw: string,
  engine: ScriptEngine,
  context: string
): T {
  if (!raw || raw.trim() === "") {
    throw new Error(
      `${context}失败: ${engine} 脚本未返回任何输出（可能超时或 Excel 正忙）`
    );
  }

  const trimmed = raw.trim();

  try {
    return JSON.parse(trimmed) as T;
  } catch {
    const startsWithBrace = trimmed.startsWith("{") || trimmed.startsWith("[");
    const endsWithBrace = trimmed.endsWith("}") || trimmed.endsWith("]");
    if (startsWithBrace && !endsWithBrace) {
      throw new Error(
        `${context}失败: ${engine} 返回的 JSON 被截断（输出过长或脚本中途异常），前 200 字符: ${trimmed.slice(0, 200)}`
      );
    }
    throw new Error(
      `${context}失败: ${engine} 返回了非 JSON 输出，前 200 字符: ${trimmed.slice(0, 200)}`
    );
  }
}
