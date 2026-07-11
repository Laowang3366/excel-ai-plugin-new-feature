/**
 * 聊天工具函数 — 从 ChatPage.tsx 提取的纯函数与小组件
 *
 * 包含：消息分组、时长计算、标题摘要、Excel 选区获取、
 *       文件大小格式化、权限图标
 */

import React from "react";
import type { AppLanguage, PermissionMode } from "../store/settingsStore";
import type { TurnItem } from "../electronApi";
import { ipcApi } from "../services/ipcApi";
import { getAppText } from "../i18n";
import {
  Shield,
  ShieldCheck,
  ShieldX,
  MessageSquare,
} from "../components/common/IconMap";

// ============================================================
// 消息分组
// ============================================================

/**
 * 将消息列表中连续的助手相关条目（reasoning, assistant_message, tool_call, tool_result）
 * 分组到一个容器中，以便它们共享宽度、实现对齐。
 */
export function groupAssistantItems(items: TurnItem[]): Array<
  | { kind: "single"; items: [TurnItem] }
  | { kind: "assistant"; items: TurnItem[]; previousUserTimestamp?: number }
> {
  const groups: Array<
    | { kind: "single"; items: [TurnItem] }
    | { kind: "assistant"; items: TurnItem[]; previousUserTimestamp?: number }
  > = [];
  let currentAssistantItems: TurnItem[] = [];
  let currentAssistantStartTimestamp: number | undefined;
  let lastUserTimestamp: number | undefined;

  const ASSISTANT_TYPES = new Set(["reasoning", "assistant_message", "tool_call", "tool_result"]);

  const flushAssistant = () => {
    if (currentAssistantItems.length > 0) {
      groups.push({
        kind: "assistant",
        items: currentAssistantItems,
        previousUserTimestamp: currentAssistantStartTimestamp,
      });
      currentAssistantItems = [];
      currentAssistantStartTimestamp = undefined;
    }
  };

  for (const item of items) {
    if (ASSISTANT_TYPES.has(item.type)) {
      if (currentAssistantItems.length === 0) {
        currentAssistantStartTimestamp = lastUserTimestamp;
      }
      currentAssistantItems.push(item);
    } else {
      flushAssistant();
      groups.push({ kind: "single", items: [item] });
      if (item.type === "user_message") {
        lastUserTimestamp = item.timestamp;
      }
    }
  }
  flushAssistant();

  return groups;
}

// ============================================================
// 时长计算
// ============================================================

export function getElapsedSeconds(startTimestamp: number | undefined, endTimestamp: number): number | undefined {
  if (!startTimestamp || endTimestamp <= startTimestamp) return undefined;
  return Math.max(1, Math.round((endTimestamp - startTimestamp) / 1000));
}

export function getItemDurationSeconds(items: TurnItem[], startTimestamp: number | undefined): Map<string, number> {
  const durations = new Map<string, number>();
  let cursorTimestamp = startTimestamp;

  for (const item of items) {
    if (!item.timestamp) continue;
    const elapsedSeconds = getElapsedSeconds(cursorTimestamp, item.timestamp);
    if (elapsedSeconds !== undefined) {
      durations.set(item.id, elapsedSeconds);
    }
    cursorTimestamp = item.timestamp;
  }

  return durations;
}

export function sumDurations(durations: Map<string, number>): number | undefined {
  let totalSeconds = 0;
  for (const duration of durations.values()) {
    totalSeconds += duration;
  }
  return totalSeconds > 0 ? totalSeconds : undefined;
}

export function getLiveTurnDurationSeconds(
  items: TurnItem[],
  startTimestamp: number | undefined,
  nowTimestamp: number
): number | undefined {
  const completedDuration = sumDurations(getItemDurationSeconds(items, startTimestamp)) ?? 0;
  const lastItemTimestamp = [...items].reverse().find((item) => item.timestamp)?.timestamp;
  const liveDuration = getElapsedSeconds(lastItemTimestamp ?? startTimestamp, nowTimestamp) ?? 0;
  const total = completedDuration + liveDuration;
  return total > 0 ? total : undefined;
}

export function formatDuration(totalSeconds: number | undefined, language: AppLanguage): string | undefined {
  if (totalSeconds === undefined) return undefined;
  const text = getAppText(language);
  if (totalSeconds < 60) return text.time.seconds(totalSeconds);
  const minutes = Math.floor(totalSeconds / 60);
  const seconds = totalSeconds % 60;
  if (minutes < 60) return text.time.minuteSecond(minutes, seconds);
  const hours = Math.floor(minutes / 60);
  const remainingMinutes = minutes % 60;
  return text.time.hourParts(hours, remainingMinutes, seconds);
}

// ============================================================
// 标题与选区
// ============================================================

const MODULE_INTERNAL_LINE_PREFIXES = [
  "模块指令：",
  "交付要求：",
  "交付方式：",
];

const MODULE_INTERNAL_EXACT_LINES = new Set([
  "数据源选区：未指定，请读取工作簿快照后自主判断。",
  "答案参考样例：未指定。",
  "答案填入锚点/选区：由 Agent 选择空白区域",
  "输出/操作锚点：未指定。",
]);

function isTaskModulePayload(lines: string[]): boolean {
  return lines.some((line) => line.trim().length > 0) &&
    (lines.find((line) => line.trim().length > 0)?.trim().startsWith("【功能模块：") ?? false);
}

function isHiddenModuleLine(line: string, taskModulePayload: boolean): boolean {
  const trimmed = line.trim();
  if (trimmed.startsWith("模块指令：")) return true;
  if (!taskModulePayload) return false;
  if (MODULE_INTERNAL_LINE_PREFIXES.some((prefix) => trimmed.startsWith(prefix))) return true;
  return MODULE_INTERNAL_EXACT_LINES.has(trimmed);
}

export function getUserFacingMessageContent(content: string): string {
  const lines = content.split(/\r?\n/);
  const taskModulePayload = isTaskModulePayload(lines);
  return lines
    .filter((line) => !isHiddenModuleLine(line, taskModulePayload))
    .join("\n")
    .trim();
}

export function getChatTitleSummary(messages: TurnItem[], fallback: string) {
  const firstUserMessage = messages.find((item) => item.type === "user_message");
  const content = firstUserMessage
    ? getUserFacingMessageContent(firstUserMessage.content).replace(/\s+/g, " ").trim()
    : "";
  if (!content) return fallback;
  return content.length > 42 ? `${content.slice(0, 42)}...` : content;
}

/** 从 Excel 获取当前选区 */
export async function pickExcelRange(): Promise<string> {
  try {
    const sel = await ipcApi.excel.getSelectionAddress();
    if (sel.address) {
      return sel.sheetName ? `${sel.sheetName}!${sel.address}` : sel.address;
    }
    return "";
  } catch {
    return "";
  }
}

// ============================================================
// 任务面板元数据
// ============================================================

/** 权限模式图标 */
export function PermissionIcon({ mode }: { mode: PermissionMode }) {
  switch (mode) {
    case "auto_approve_safe":
      return <ShieldCheck size={15} />;
    case "confirm_all":
      return <ShieldX size={15} />;
    default:
      return <Shield size={15} />;
  }
}

/** 聊天头部小图标 */
export function MessageBubbleIcon() {
  return <MessageSquare size={16} style={{ verticalAlign: "middle" }} />;
}
