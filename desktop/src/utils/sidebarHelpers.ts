/**
 * 侧边栏工具函数 — 从 Sidebar.tsx 提取的纯函数与常量
 *
 * 包含：意图快捷定义、时间格式化、会话状态判断
 */

import type { AppLanguage } from "../store/settingsStore";
import { getAppText } from "../i18n";
import {
  Hash,
  Code,
  FileScan,
  Eraser,
  FileBarChart,
  LineChart,
} from "../components/common/IconMap";

/** 6 个核心意图（侧边栏快捷入口） */
export const INTENT_SHORTCUTS = [
  { key: "formula" as const, icon: Hash },
  { key: "code" as const, icon: Code },
  { key: "ocr" as const, icon: FileScan },
  { key: "clean" as const, icon: Eraser },
  { key: "report" as const, icon: FileBarChart },
  { key: "chart" as const, icon: LineChart },
] as const;

export type IntentKind = typeof INTENT_SHORTCUTS[number]["key"] | null;

/** Excel 连接状态 */
export interface ExcelStatus {
  connected: boolean;
  host: string;
  version?: string;
  workbookName?: string;
}

/** 格式化时间 */
export function formatTime(timestamp: number, language: AppLanguage): string {
  const text = getAppText(language);
  const now = Date.now();
  const diff = now - timestamp;
  if (diff < 60000) return text.sidebar.justNow;
  if (diff < 3600000) return text.sidebar.minutesAgo(Math.floor(diff / 60000));
  if (diff < 86400000) return text.sidebar.hoursAgo(Math.floor(diff / 3600000));
  const date = new Date(timestamp);
  return `${date.getMonth() + 1}/${date.getDate()}`;
}

export function getThreadDisplayStatus(
  metadataStatus?: string,
  activeStatus?: "idle" | "in_progress" | "completed" | "interrupted" | "failed",
  statusViewed = false
): "running" | "completed" | "failed" | null {
  if (activeStatus === "in_progress") return "running";
  const status = activeStatus && activeStatus !== "idle" ? activeStatus : metadataStatus;
  if (!status || statusViewed) return null;
  if (status === "in_progress") return "running";
  if (status === "completed") return "completed";
  if (status === "failed" || status === "interrupted") return "failed";
  return null;
}

export function getThreadStatusLabel(status: ReturnType<typeof getThreadDisplayStatus>, language: AppLanguage): string {
  if (language === "en-US") {
    if (status === "running") return "In conversation";
    if (status === "completed") return "Completed";
    if (status === "failed") return "Error";
    return "";
  }
  if (status === "running") return "对话中";
  if (status === "completed") return "对话完成";
  if (status === "failed") return "异常";
  return "";
}

export function matchesSidebarSearch(values: Array<string | undefined | null>, query: string): boolean {
  const normalizedQuery = query.trim().toLowerCase();
  if (!normalizedQuery) return true;
  return values.some((value) => value?.toLowerCase().includes(normalizedQuery));
}
