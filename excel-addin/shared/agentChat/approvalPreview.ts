import type { ToolCall } from "../tools/types";

const SECRET_KEY_RE =
  /^(password|apiKey|api_key|token|secret|authorization|accessToken|refreshToken|bearer)$/i;
const BINARY_KEY_RE =
  /^(imageBase64|base64|dataUrl|dataURL|thumbnailBase64|contentBase64|payloadBase64)$/i;

const MAX_DEPTH = 4;
const MAX_KEYS = 12;
const MAX_ARRAY = 8;
const MAX_STRING = 80;

export type ArgsPreview = unknown;

/** Budgeted structured preview for approval UI — never stores raw args. */
export function buildArgsPreview(args: Record<string, unknown>): ArgsPreview {
  return sanitize(args, 0, new Set());
}

export function isDestructiveTool(
  name: string,
  args: Record<string, unknown>,
): boolean {
  if (
    /\.delete$/.test(name) ||
    /\.clear$/.test(name) ||
    name === "table.unlist" ||
    name.endsWith(".unlist")
  ) {
    return true;
  }
  if (name === "sheet.operation") {
    const op = args.operation;
    if (typeof op === "string" && op.trim().toLowerCase() === "delete") {
      return true;
    }
  }
  return false;
}

export function buildImpactHint(
  name: string,
  args: Record<string, unknown>,
  destructive: boolean,
): string {
  if (destructive) {
    if (name === "sheet.delete" || (name === "sheet.operation" && isDestructiveTool(name, args))) {
      return "将删除工作表，可能造成不可恢复的数据丢失。";
    }
    if (name.includes("delete") || name.includes("clear") || name.includes("unlist")) {
      return "将删除或清除现有内容，请确认范围正确。";
    }
    return "此操作可能破坏性修改工作簿，需你确认后才会执行。";
  }
  if (name.includes("write") || name.includes("set") || name.includes("create") || name.includes("add") || name.includes("update") || name.includes("rename")) {
    return "将修改工作簿内容，需你确认后才会执行。";
  }
  if (name.includes("protection") || name.includes("password")) {
    return "将更改保护/密码相关设置，需你确认。";
  }
  return "将执行变更操作，需你确认后才会写入宿主。";
}

function sanitize(value: unknown, depth: number, seen: Set<object>): unknown {
  if (value == null) return value;
  if (typeof value === "string") {
    if (value.length > MAX_STRING) {
      if (/^[A-Za-z0-9+/=\s]+$/.test(value.slice(0, 120))) {
        return `[binary/base64 ${value.length} chars]`;
      }
      return `${value.slice(0, MAX_STRING - 1)}…`;
    }
    return value;
  }
  if (typeof value === "number" || typeof value === "boolean") return value;
  if (typeof value === "bigint") return String(value);
  if (typeof value === "function") return "[function]";
  if (typeof value !== "object") return String(value);
  if (depth >= MAX_DEPTH) return "[max depth]";
  if (seen.has(value as object)) return "[circular]";
  seen.add(value as object);

  if (Array.isArray(value)) {
    const out: unknown[] = [];
    const n = Math.min(value.length, MAX_ARRAY);
    for (let i = 0; i < n; i += 1) {
      out.push(sanitize(value[i], depth + 1, seen));
    }
    if (value.length > MAX_ARRAY) out.push(`[+${value.length - MAX_ARRAY} more]`);
    return out;
  }

  const obj = value as Record<string, unknown>;
  const keys = Object.keys(obj);
  const out: Record<string, unknown> = {};
  const n = Math.min(keys.length, MAX_KEYS);
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
    // Cell grids can be large / sensitive — summarize shape only.
    if (key === "values" && Array.isArray(obj[key])) {
      const grid = obj[key] as unknown[];
      const rows = grid.length;
      const cols = Array.isArray(grid[0]) ? (grid[0] as unknown[]).length : 0;
      out[key] = `[grid ${rows}x${cols}]`;
      continue;
    }
    out[key] = sanitize(obj[key], depth + 1, seen);
  }
  if (keys.length > MAX_KEYS) out["…"] = `+${keys.length - MAX_KEYS} keys`;
  return out;
}

export function previewFromToolCall(call: ToolCall): {
  argsPreview: ArgsPreview;
  destructive: boolean;
  impactHint: string;
} {
  const args = call.arguments ?? {};
  const destructive = isDestructiveTool(call.name, args);
  return {
    argsPreview: buildArgsPreview(args),
    destructive,
    impactHint: buildImpactHint(call.name, args, destructive),
  };
}
