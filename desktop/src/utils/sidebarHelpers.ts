/**
 * 侧边栏工具函数 — 从 Sidebar.tsx 提取的纯函数与常量
 *
 * 包含：意图快捷定义、时间格式化、会话状态判断
 */

import type { FolderFileInfo, ThreadMetadata } from "../electronApi";
import type { AppLanguage, PinnedFolder } from "../store/settingsStore";
import { getAppText } from "../i18n";
import {
  Hash,
  Code,
  FileScan,
  Eraser,
  FileBarChart,
  LineChart,
  Workflow,
} from "../components/common/IconMap";

/** 6 个核心意图（侧边栏快捷入口） */
export const INTENT_SHORTCUTS = [
  { key: "formula" as const, icon: Hash },
  { key: "code" as const, icon: Code },
  { key: "ocr" as const, icon: FileScan },
  { key: "clean" as const, icon: Eraser },
  { key: "report" as const, icon: FileBarChart },
  { key: "chart" as const, icon: LineChart },
  { key: "office" as const, icon: Workflow },
] as const;

export type IntentKind = (typeof INTENT_SHORTCUTS)[number]["key"] | null;

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
  statusViewed = false,
): "running" | "completed" | "failed" | null {
  if (activeStatus === "in_progress") return "running";
  const status = activeStatus && activeStatus !== "idle" ? activeStatus : metadataStatus;
  if (!status || statusViewed) return null;
  if (status === "in_progress") return "running";
  if (status === "completed") return "completed";
  if (status === "failed" || status === "interrupted") return "failed";
  return null;
}

export function getThreadStatusLabel(
  status: ReturnType<typeof getThreadDisplayStatus>,
  language: AppLanguage,
): string {
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

export function matchesSidebarSearch(
  values: Array<string | undefined | null>,
  query: string,
): boolean {
  const normalizedQuery = query.trim().toLowerCase();
  if (!normalizedQuery) return true;
  return values.some((value) => value?.toLowerCase().includes(normalizedQuery));
}

export type SidebarSortMode = "recentDesc" | "recentAsc" | "nameAsc" | "nameDesc";

export interface SidebarGroupedFolder {
  folder: PinnedFolder;
  threads: ThreadMetadata[];
  files: FolderFileInfo[];
}

export function compareSidebarText(left: string, right: string, language: AppLanguage): number {
  return left.localeCompare(right, language === "zh-CN" ? "zh-CN" : "en", {
    numeric: true,
    sensitivity: "base",
  });
}

export function sortSidebarItems<T extends { preview?: string; updatedAt: number }>(
  items: T[],
  mode: SidebarSortMode,
  language: AppLanguage,
): T[] {
  return [...items].sort((a, b) => {
    if (mode === "recentAsc") return a.updatedAt - b.updatedAt;
    if (mode === "nameAsc") return compareSidebarText(a.preview || "", b.preview || "", language);
    if (mode === "nameDesc") return compareSidebarText(b.preview || "", a.preview || "", language);
    return b.updatedAt - a.updatedAt;
  });
}

export interface SidebarDerivedListsParams {
  threads: ThreadMetadata[];
  pinnedFolders: PinnedFolder[];
  folderFiles: Record<string, FolderFileInfo[]>;
  projectSortMode: SidebarSortMode;
  conversationSortMode: SidebarSortMode;
  language: AppLanguage;
  hasSearchQuery?: boolean;
}

export interface SidebarDerivedLists {
  ungroupedThreads: ThreadMetadata[];
  groupedByFolder: SidebarGroupedFolder[];
  hasProjectItems: boolean;
  hasConversationItems: boolean;
  showNoSearchResults: boolean;
}

export function buildSidebarDerivedLists({
  threads,
  pinnedFolders,
  folderFiles,
  projectSortMode,
  conversationSortMode,
  language,
  hasSearchQuery = false,
}: SidebarDerivedListsParams): SidebarDerivedLists {
  const ungroupedThreads = sortSidebarItems(
    threads.filter((thread) => !thread.folderId),
    conversationSortMode,
    language,
  );
  const groupedByFolder = pinnedFolders
    .map((folder) => ({
      folder,
      folderMatches: true,
      threads: sortSidebarItems(
        threads.filter((thread) => thread.folderId === folder.path),
        projectSortMode,
        language,
      ),
      files: folderFiles[folder.path] || [],
    }))
    .filter(
      ({ folderMatches, threads: folderThreads, files }) =>
        !hasSearchQuery || folderMatches || folderThreads.length > 0 || files.length > 0,
    )
    .sort((a, b) => {
      if (projectSortMode === "recentAsc") return a.folder.addedAt - b.folder.addedAt;
      if (projectSortMode === "nameAsc")
        return compareSidebarText(a.folder.name, b.folder.name, language);
      if (projectSortMode === "nameDesc")
        return compareSidebarText(b.folder.name, a.folder.name, language);
      return b.folder.addedAt - a.folder.addedAt;
    })
    .map(({ folder, threads: folderThreads, files }) => ({
      folder,
      threads: folderThreads,
      files,
    }));
  const hasProjectItems = groupedByFolder.length > 0;
  const hasConversationItems = ungroupedThreads.length > 0;

  return {
    ungroupedThreads,
    groupedByFolder,
    hasProjectItems,
    hasConversationItems,
    showNoSearchResults: hasSearchQuery && !hasProjectItems && !hasConversationItems,
  };
}
