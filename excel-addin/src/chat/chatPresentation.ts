import type { AgentToolOutcome } from "../../shared/agent/types";
import type {
  ChatPublicError,
  ChatTraceEvent,
  ChatTurnStatus,
} from "@shared/agentChat";

export const MAX_TRACE_TEXT = 160;

/** Skip JSON.parse for oversized tool args to avoid memory spikes / secret leaks. */
export const MAX_PARSEABLE_TOOL_ARGS_CHARS = 4096;

export type DisplayRole = "user" | "assistant" | "system";

export interface DisplayMessage {
  id: string;
  role: DisplayRole;
  content: string;
  pending?: boolean;
}

export interface DisplayTraceItem {
  id: string;
  kind: "round" | "tool_parsed" | "tool_outcome" | "approval";
  text: string;
  tone?: "ok" | "fail" | "info" | "warn";
}

export interface DisplayTurn {
  id: string;
  userText: string;
  assistantText: string;
  pending: boolean;
  turnStatus?: ChatTurnStatus;
  errorText?: string;
  traces: DisplayTraceItem[];
}

export function truncateDisplay(text: string, max = MAX_TRACE_TEXT): string {
  const normalized = text.replace(/\s+/g, " ").trim();
  if (normalized.length <= max) return normalized;
  return `${normalized.slice(0, Math.max(0, max - 1))}…`;
}

/** Collapse huge base64/json blobs for UI-only display. */
export function summarizePayload(raw: string, max = MAX_TRACE_TEXT): string {
  if (!raw) return "";
  // Long base64-looking payload
  if (raw.length > 80 && /^[A-Za-z0-9+/=\s]+$/.test(raw.slice(0, 200)) && raw.length > max) {
    return `[binary/base64 ${raw.length} chars]`;
  }
  if (raw.length > max * 2) {
    return truncateDisplay(raw, max);
  }
  return truncateDisplay(raw, max);
}

export function formatToolArgs(argsJson: string | undefined): string {
  if (argsJson == null || argsJson === "") return "{}";
  if (argsJson.length > MAX_PARSEABLE_TOOL_ARGS_CHARS) {
    return `[arguments JSON omitted: ${argsJson.length} chars]`;
  }
  try {
    return safeJson(JSON.parse(argsJson) as unknown);
  } catch {
    // Never echo raw invalid JSON prefixes (may contain password/token).
    return `[invalid arguments JSON: ${argsJson.length} chars]`;
  }
}

export function formatToolOutcome(outcome: AgentToolOutcome): {
  text: string;
  tone: "ok" | "fail" | "info";
} {
  if (outcome.kind === "host") {
    if (outcome.result.ok) {
      const dataPreview = summarizePayload(safeJson(outcome.result.data));
      return {
        text: `✓ ${outcome.toolName}${dataPreview ? ` · ${dataPreview}` : ""}`,
        tone: "ok",
      };
    }
    const err = outcome.result.error || "failed";
    return {
      text: `✗ ${outcome.toolName} · ${truncateDisplay(err)}`,
      tone: "fail",
    };
  }
  if (outcome.kind === "unknown_tool") {
    return {
      text: `? 未知工具 ${outcome.toolName} · ${truncateDisplay(outcome.error)}`,
      tone: "fail",
    };
  }
  return {
    text: `! 参数无效 ${outcome.toolName} · ${truncateDisplay(outcome.error)}`,
    tone: "fail",
  };
}

export function mapChatError(
  error: ChatPublicError | undefined,
  turnStatus?: ChatTurnStatus,
): string | undefined {
  if (turnStatus === "busy") return "当前对话进行中，请稍候。";
  if (turnStatus === "empty") return "请输入内容后再发送。";
  if (turnStatus === "max_rounds") {
    return "已达到本轮最大工具调用轮数，请精简问题后重试。";
  }
  if (turnStatus === "aborted") return "已停止生成。进行中的表格操作可能仍会完成。";
  if (!error) {
    if (turnStatus === "failed") return "请求失败，请稍后重试。";
    return undefined;
  }
  const kind = error.kind ?? "";
  const msg = error.message || "";
  if (kind === "missing_key" || /API key|密钥|未设置/.test(msg)) {
    return "未配置 API 密钥。请到「模型供应商」页添加并选择可用供应商。";
  }
  if (/no active provider/i.test(msg) || /active provider/i.test(msg)) {
    return "未选择活动模型供应商。请到「模型供应商」页配置并设为当前。";
  }
  if (kind === "cors") {
    return "浏览器 CORS/网络拦截：任务窗格直连第三方 API 常被拒绝。请检查供应商地址或网络环境。";
  }
  if (kind === "network") {
    return `网络错误：${truncateDisplay(msg, 120)}`;
  }
  if (kind === "http") {
    const status = error.status != null ? `HTTP ${error.status}` : "HTTP 错误";
    return `${status}：${truncateDisplay(msg, 120)}`;
  }
  if (kind === "parse") {
    return `响应解析失败：${truncateDisplay(msg, 120)}`;
  }
  if (kind === "provider") {
    return `模型服务错误：${truncateDisplay(msg, 120)}`;
  }
  if (kind === "aborted") {
    return "已停止生成。进行中的表格操作可能仍会完成。";
  }
  return truncateDisplay(msg || "请求失败", 160);
}

export function projectTraceEvent(
  event: ChatTraceEvent,
  seq: number,
): DisplayTraceItem | null {
  switch (event.type) {
    case "approval_needed": {
      const risk =
        event.request.riskLevel === "dangerous" ? "高风险" : "需批准";
      const dest = event.request.destructive ? " · 破坏性" : "";
      return {
        id: `tr-${seq}`,
        kind: "approval",
        text: `待审批 ${event.request.name} · ${risk}${dest} · ${truncateDisplay(event.request.impactHint, 80)}`,
        tone: "warn",
      };
    }
    case "approval_resolved": {
      if (event.decision === "approved") {
        return {
          id: `tr-${seq}`,
          kind: "approval",
          text: `已批准 ${event.request.name}`,
          tone: "ok",
        };
      }
      if (event.decision === "rejected") {
        return {
          id: `tr-${seq}`,
          kind: "approval",
          text: `已拒绝 ${event.request.name}`,
          tone: "fail",
        };
      }
      return {
        id: `tr-${seq}`,
        kind: "approval",
        text: `已取消审批 ${event.request.name}`,
        tone: "info",
      };
    }
    case "round_start":
      return {
        id: `tr-${seq}`,
        kind: "round",
        text: `回合 ${event.round} 开始`,
        tone: "info",
      };
    case "round_end":
      return {
        id: `tr-${seq}`,
        kind: "round",
        text: `回合 ${event.round} 结束 · ${event.finishReason} · 工具 ${event.toolCallCount}`,
        tone: "info",
      };
    case "tool_call_parsed":
      return {
        id: `tr-${seq}`,
        kind: "tool_parsed",
        text: `调用 ${event.call.name}(${formatToolArgs(event.call.argumentsJson)})`,
        tone: "info",
      };
    case "tool_outcome": {
      const formatted = formatToolOutcome(event.outcome);
      return {
        id: `tr-${seq}`,
        kind: "tool_outcome",
        text: formatted.text,
        tone: formatted.tone,
      };
    }
    default:
      return null;
  }
}

const SECRET_KEY_RE =
  /^(password|apiKey|api_key|token|secret|authorization|accessToken|refreshToken|bearer)$/i;
const BINARY_KEY_RE =
  /^(imageBase64|base64|dataUrl|dataURL|thumbnailBase64|contentBase64|payloadBase64)$/i;
const MAX_JSON_DEPTH = 4;
const MAX_OBJECT_KEYS = 12;
const MAX_ARRAY_ITEMS = 8;
const MAX_STRING_LEN = 80;

/** Budgeted, key-aware sanitizer — never walks/stringifies unbounded trees. */
export function safeJson(value: unknown): string {
  try {
    const sanitized = sanitizeForTrace(value, 0, new Set());
    const text = JSON.stringify(sanitized) ?? "";
    return truncateDisplay(text, MAX_TRACE_TEXT);
  } catch {
    return "[unserializable]";
  }
}

function sanitizeForTrace(
  value: unknown,
  depth: number,
  seen: Set<object>,
): unknown {
  if (value == null) return value;
  if (typeof value === "string") {
    if (value.length > MAX_STRING_LEN) {
      if (/^[A-Za-z0-9+/=\s]+$/.test(value.slice(0, 120))) {
        return `[binary/base64 ${value.length} chars]`;
      }
      return truncateDisplay(value, MAX_STRING_LEN);
    }
    return value;
  }
  if (typeof value === "number" || typeof value === "boolean") return value;
  if (typeof value === "bigint") return String(value);
  if (typeof value === "function") return "[function]";
  if (typeof value !== "object") return String(value);

  if (depth >= MAX_JSON_DEPTH) return "[max depth]";
  if (seen.has(value as object)) return "[circular]";
  seen.add(value as object);

  if (Array.isArray(value)) {
    const out: unknown[] = [];
    const n = Math.min(value.length, MAX_ARRAY_ITEMS);
    for (let i = 0; i < n; i += 1) {
      out.push(sanitizeForTrace(value[i], depth + 1, seen));
    }
    if (value.length > MAX_ARRAY_ITEMS) {
      out.push(`[+${value.length - MAX_ARRAY_ITEMS} more]`);
    }
    return out;
  }

  const obj = value as Record<string, unknown>;
  const keys = Object.keys(obj);
  const out: Record<string, unknown> = {};
  const n = Math.min(keys.length, MAX_OBJECT_KEYS);
  for (let i = 0; i < n; i += 1) {
    const key = keys[i]!;
    if (SECRET_KEY_RE.test(key)) {
      out[key] = "[REDACTED]";
      continue;
    }
    if (BINARY_KEY_RE.test(key)) {
      const raw = obj[key];
      const len = typeof raw === "string" ? raw.length : undefined;
      out[key] = len != null ? `[omitted binary ${len} chars]` : "[omitted binary]";
      continue;
    }
    // Only access selected key after budget checks above.
    out[key] = sanitizeForTrace(obj[key], depth + 1, seen);
  }
  if (keys.length > MAX_OBJECT_KEYS) {
    out["…"] = `+${keys.length - MAX_OBJECT_KEYS} keys`;
  }
  return out;
}
