/**
 * 侧边栏 — 会话管理 + 意图快捷 + 导航 + 连接状态
 *
 * 已拆分模块：
 * - utils/sidebarHelpers.ts: 意图常量、时间格式化、状态判断
 * - hooks/useExcelConnection.ts: Excel/WPS 连接状态管理
 * - components/sidebar/FolderSection.tsx: 文件夹分组渲染
 * - components/sidebar/ThreadContextMenu.tsx: 右键菜单
 */

import React, { useEffect, useState, useCallback, useRef } from "react";
import { useChatStore } from "../store/chatStore";
import { useSettingsStore } from "../store/settingsStore";
import type { AppPage } from "../App";
import type { SettingsSection } from "./SettingsPage";
import type { FolderFileInfo } from "../electronApi";
import { getAppText } from "../i18n";
import { ipcApi } from "../services/ipcApi";
import {
  type IntentKind,
  type ExcelStatus,
} from "../utils/sidebarHelpers";
import { useExcelConnection } from "../hooks/useExcelConnection";
import { useOfficeConnection } from "../hooks/useOfficeConnection";
import { HostSelectionDialog } from "./excel/HostSelectionDialog";
import { FolderSection, UngroupedThreadList } from "./sidebar/FolderSection";
import { ThreadContextMenu, type ContextMenuState } from "./sidebar/ThreadContextMenu";
import { FileContextMenu, type FileContextMenuState } from "./sidebar/FileContextMenu";
import { SidebarSearchPalette } from "./sidebar/SidebarSearchPalette";
import {
  Settings,
  User,
  Package,
  Plus,
  Search,
  RefreshCw,
  LogOut,
  FolderOpen,
  MessageSquare,
  PenLine,
  ChevronDown,
  ChevronRight,
  ClipboardList,
  Clock,
  Check,
} from "./common/IconMap";

// Re-export IntentKind for ChatPage
export type { IntentKind } from "../utils/sidebarHelpers";

interface SidebarProps {
  collapsed: boolean;
  currentPage: AppPage;
  onNavigate: (page: AppPage) => void;
  onOpenSettingsSection?: (section: SettingsSection) => void;
  activeIntent: IntentKind;
  onIntentClick: (intent: IntentKind) => void;
}

type SidebarSortSection = "projects" | "conversations";
type SidebarSortMode = "recentDesc" | "recentAsc" | "nameAsc" | "nameDesc";

export const Sidebar: React.FC<SidebarProps> = ({
  collapsed,
  currentPage,
  onNavigate,
  onOpenSettingsSection,
  activeIntent,
  onIntentClick,
}) => {
  const { threads, activeThreadId, runningThreadIds, turnStatus, loadThreads, switchThread, createNewThread, deleteThread, moveThreadToFolder, addFilesToComposer } =
    useChatStore();
  const { language, pinnedFolders, addPinnedFolder, removePinnedFolder, updatePinnedFolder } = useSettingsStore();
  const text = getAppText(language);

  // ---- 文件夹状态 ----
  const [folderFiles, setFolderFiles] = useState<Record<string, FolderFileInfo[]>>({});
  const [expandedFolders, setExpandedFolders] = useState<Record<string, boolean>>({});
  const [creatingFolderThread, setCreatingFolderThread] = useState<string | null>(null);
  const [creatingNewThread, setCreatingNewThread] = useState(false);

  // 添加文件夹
  const handleAddFolder = useCallback(async () => {
    try {
      const result = await ipcApi.dialog.openFolder();
      if (!result.canceled && result.filePaths.length > 0) {
        const folderPath = result.filePaths[0];
        const folderName = folderPath.split(/[\\/]/).pop() || folderPath;
        addPinnedFolder({ path: folderPath, name: folderName, addedAt: Date.now() });
        setExpandedFolders((prev) => ({ ...prev, [folderPath]: true }));
        const files = await ipcApi.folder.listFiles(folderPath);
        setFolderFiles((prev) => ({ ...prev, [folderPath]: files }));
      }
    } catch { /* ignore */ }
  }, [addPinnedFolder]);

  // 展开/折叠文件夹（懒加载文件列表）
  const handleToggleFolder = useCallback(async (folderPath: string) => {
    setExpandedFolders((prev) => {
      const next = !prev[folderPath];
      if (next && !folderFiles[folderPath]) {
        ipcApi.folder.listFiles(folderPath).then((files) => {
          setFolderFiles((prev2) => ({ ...prev2, [folderPath]: files }));
        });
      }
      return { ...prev, [folderPath]: next };
    });
  }, [folderFiles]);

  const handleAddFile = useCallback((file: FolderFileInfo) => {
    addFilesToComposer([{
      filePath: file.filePath,
      fileName: file.fileName,
      fileType: "document",
      size: file.size,
    }]);
  }, [addFilesToComposer]);

  const handleCreateFolderThread = useCallback(async (folderPath: string) => {
    setCreatingFolderThread(folderPath);
    try {
      await createNewThread(folderPath);
    } finally {
      setTimeout(() => setCreatingFolderThread(null), 300);
    }
  }, [createNewThread]);

  const handleCreateNewThread = useCallback(async () => {
    setCreatingNewThread(true);
    try {
      await createNewThread();
    } finally {
      setTimeout(() => setCreatingNewThread(false), 300);
    }
  }, [createNewThread]);

  const handleToggleSearch = useCallback(() => {
    setSearchOpen((open) => !open);
  }, []);

  // Excel 连接 hook
  const {
    excelStatus,
    connecting,
    connectFailed,
    pulseDot,
    pendingHosts,
    handleConnect,
    handleSelectHost,
    setPendingHosts,
  } = useExcelConnection();
  const {
    wordStatus,
    presentationStatus,
  } = useOfficeConnection();

  // 已查看状态追踪
  const [viewedThreadStatusAt, setViewedThreadStatusAt] = useState<Record<string, number>>({});
  const initializedViewedStatuses = useRef(false);
  const [settingsMenuOpen, setSettingsMenuOpen] = useState(false);
  const [searchOpen, setSearchOpen] = useState(false);
  const [projectsExpanded, setProjectsExpanded] = useState(true);
  const [conversationsExpanded, setConversationsExpanded] = useState(true);
  const [sortMenu, setSortMenu] = useState<{
    section: SidebarSortSection;
    x: number;
    y: number;
  } | null>(null);
  const [projectSortMode, setProjectSortMode] = useState<SidebarSortMode>("recentDesc");
  const [conversationSortMode, setConversationSortMode] = useState<SidebarSortMode>("recentDesc");

  // 右键菜单状态
  const [contextMenu, setContextMenu] = useState<ContextMenuState | null>(null);
  const [fileContextMenu, setFileContextMenu] = useState<FileContextMenuState | null>(null);

  // 拖拽调整宽度
  const [sidebarWidth, setSidebarWidth] = useState(260);
  const resizingRef = useRef(false);
  const sidebarRef = useRef<HTMLElement>(null);
  const [isResizing, setIsResizing] = useState(false);

  const handleResizeStart = useCallback((e: React.MouseEvent) => {
    e.preventDefault();
    resizingRef.current = true;
    setIsResizing(true);
    const startX = e.clientX;
    const startWidth = sidebarWidth;

    const handleMouseMove = (e: MouseEvent) => {
      if (!resizingRef.current) return;
      const delta = e.clientX - startX;
      const newWidth = Math.min(400, Math.max(180, startWidth + delta));
      setSidebarWidth(newWidth);
    };

    const handleMouseUp = () => {
      resizingRef.current = false;
      setIsResizing(false);
      document.removeEventListener("mousemove", handleMouseMove);
      document.removeEventListener("mouseup", handleMouseUp);
    };

    document.addEventListener("mousemove", handleMouseMove);
    document.addEventListener("mouseup", handleMouseUp);
  }, [sidebarWidth]);

  useEffect(() => { loadThreads(); }, [loadThreads]);

  useEffect(() => {
    if (!searchOpen) return;
    pinnedFolders.forEach((folder) => {
      if (folderFiles[folder.path]) return;
      ipcApi.folder.listFiles(folder.path).then((files) => {
        setFolderFiles((prev) => ({ ...prev, [folder.path]: files }));
      }).catch(() => {
        setFolderFiles((prev) => ({ ...prev, [folder.path]: [] }));
      });
    });
  }, [folderFiles, pinnedFolders, searchOpen]);

  useEffect(() => {
    if (initializedViewedStatuses.current || threads.length === 0) return;
    initializedViewedStatuses.current = true;
    const viewed: Record<string, number> = {};
    threads.forEach((thread) => { viewed[thread.threadId] = thread.updatedAt; });
    setViewedThreadStatusAt(viewed);
  }, [threads]);

  const markThreadViewed = useCallback((threadId: string) => {
    const thread = threads.find((item) => item.threadId === threadId);
    if (!thread) return;
    setViewedThreadStatusAt((prev) => {
      if (prev[threadId] === thread.updatedAt) return prev;
      return { ...prev, [threadId]: thread.updatedAt };
    });
  }, [threads]);

  useEffect(() => {
    if (!activeThreadId) return;
    markThreadViewed(activeThreadId);
  }, [activeThreadId, markThreadViewed]);

  const handleSwitchThread = useCallback((threadId: string) => {
    markThreadViewed(threadId);
    switchThread(threadId);
  }, [markThreadViewed, switchThread]);

  const handleOpenSortMenu = useCallback((e: React.MouseEvent, section: SidebarSortSection) => {
    e.stopPropagation();
    setContextMenu(null);
    setFileContextMenu(null);
    const rect = e.currentTarget.getBoundingClientRect();
    const menuWidth = 168;
    setSortMenu({
      section,
      x: Math.max(8, Math.min(window.innerWidth - menuWidth - 8, rect.right - menuWidth)),
      y: Math.max(8, Math.min(window.innerHeight - 164, rect.bottom + 6)),
    });
  }, []);

  // 右键会话菜单
  const handleThreadContextMenu = useCallback((e: React.MouseEvent, threadId: string, inFolder?: boolean) => {
    e.preventDefault();
    e.stopPropagation();
    setFileContextMenu(null);
    setContextMenu({ threadId, x: e.clientX, y: e.clientY, confirming: false, inFolder });
  }, []);

  const handleConfirmDelete = useCallback(async (threadId: string) => {
    setContextMenu(null);
    await deleteThread(threadId);
  }, [deleteThread]);

  const handleMoveToFolder = useCallback(async (threadId: string, folderId?: string) => {
    setContextMenu(null);
    await moveThreadToFolder(threadId, folderId);
  }, [moveThreadToFolder]);

  const handlePinThread = useCallback(async (threadId: string) => {
    setContextMenu(null);
    // TODO: implement pin via ipcApi.thread.updateMetadata
  }, []);

  const handleRenameThread = useCallback(async (threadId: string) => {
    setContextMenu(null);
    // TODO: implement rename via prompt + ipcApi.thread.updateMetadata
  }, []);

  // ---- 文件右键菜单 ----
  const handleFileContextMenu = useCallback((e: React.MouseEvent, file: FolderFileInfo, isPinned: boolean) => {
    e.preventDefault();
    e.stopPropagation();
    setContextMenu(null);
    setFileContextMenu({ file, x: e.clientX, y: e.clientY, isPinned });
  }, []);

  const handleTrashFile = useCallback(async (filePath: string) => {
    setFileContextMenu(null);
    await ipcApi.file.trashFile(filePath);
    // 刷新文件列表
    const folderPath = Object.keys(folderFiles).find(fp =>
      folderFiles[fp].some(f => f.filePath === filePath)
    );
    if (folderPath) {
      const files = await ipcApi.folder.listFiles(folderPath);
      setFolderFiles(prev => ({ ...prev, [folderPath]: files }));
    }
  }, [folderFiles]);

  const handleOpenFile = useCallback(async (filePath: string) => {
    setFileContextMenu(null);
    await ipcApi.file.openFile(filePath);
  }, []);

  const handleCopyPath = useCallback(async (filePath: string) => {
    setFileContextMenu(null);
    await ipcApi.file.copyPath(filePath);
  }, []);

  const handleRevealInExplorer = useCallback(async (filePath: string) => {
    setFileContextMenu(null);
    await ipcApi.file.revealInExplorer(filePath);
  }, []);

  const handlePinFile = useCallback(async (filePath: string) => {
    setFileContextMenu(null);
    // 在当前 folder 的 pinnedFiles 中切换
    const folderPath = Object.keys(folderFiles).find(fp =>
      folderFiles[fp].some(f => f.filePath === filePath)
    );
    if (!folderPath) return;
    const folder = pinnedFolders.find(f => f.path === folderPath);
    if (!folder) return;
    const pinned = folder.pinnedFiles || [];
    const idx = pinned.indexOf(filePath);
    const newPinned = idx >= 0 ? pinned.filter(p => p !== filePath) : [...pinned, filePath];
    updatePinnedFolder(folderPath, { pinnedFiles: newPinned });
  }, [folderFiles, pinnedFolders]);

  useEffect(() => {
    if (!contextMenu) return;
    const close = () => setContextMenu(null);
    const closeOnEscape = (e: KeyboardEvent) => { if (e.key === "Escape") close(); };
    document.addEventListener("click", close);
    document.addEventListener("keydown", closeOnEscape);
    return () => {
      document.removeEventListener("click", close);
      document.removeEventListener("keydown", closeOnEscape);
    };
  }, [contextMenu]);

  useEffect(() => {
    if (!fileContextMenu) return;
    const close = () => setFileContextMenu(null);
    const closeOnEscape = (e: KeyboardEvent) => { if (e.key === "Escape") close(); };
    document.addEventListener("click", close);
    document.addEventListener("keydown", closeOnEscape);
    return () => {
      document.removeEventListener("click", close);
      document.removeEventListener("keydown", closeOnEscape);
    };
  }, [fileContextMenu]);

  useEffect(() => {
    if (!settingsMenuOpen) return;
    const close = () => setSettingsMenuOpen(false);
    const closeOnEscape = (e: KeyboardEvent) => { if (e.key === "Escape") close(); };
    document.addEventListener("click", close);
    document.addEventListener("keydown", closeOnEscape);
    return () => {
      document.removeEventListener("click", close);
      document.removeEventListener("keydown", closeOnEscape);
    };
  }, [settingsMenuOpen]);

  useEffect(() => {
    if (!sortMenu) return;
    const close = () => setSortMenu(null);
    const closeOnEscape = (e: KeyboardEvent) => { if (e.key === "Escape") close(); };
    document.addEventListener("click", close);
    document.addEventListener("keydown", closeOnEscape);
    return () => {
      document.removeEventListener("click", close);
      document.removeEventListener("keydown", closeOnEscape);
    };
  }, [sortMenu]);

  const handleSelectSortMode = useCallback((section: SidebarSortSection, mode: SidebarSortMode) => {
    if (section === "projects") {
      setProjectSortMode(mode);
    } else {
      setConversationSortMode(mode);
    }
    setSortMenu(null);
  }, []);

  const openSettingsSection = useCallback((section: SettingsSection) => {
    setSettingsMenuOpen(false);
    if (onOpenSettingsSection) {
      onOpenSettingsSection(section);
    } else {
      onNavigate("settings");
    }
  }, [onNavigate, onOpenSettingsSection]);

  // ---- 折叠态 ---- 展开态 ----
  let sidebarContent: React.ReactElement | null = null;
  let expandedContent: React.ReactElement | null = null;

  if (collapsed) {
    sidebarContent = (
    <aside className="sidebar sidebar-collapsed">
      <button
        className={`sidebar-icon-btn${creatingNewThread ? " creating" : ""}`}
        onClick={handleCreateNewThread}
        disabled={creatingNewThread}
        title={text.sidebar.newThread}
      >
        {creatingNewThread ? <RefreshCw size={18} className="spin" /> : <Plus size={18} />}
      </button>
      <div className="sidebar-spacer" />
      <button
        className="sidebar-icon-btn"
        onClick={handleToggleSearch}
        title={text.sidebar.search}
      >
        <Search size={16} />
      </button>
      <button
        className="sidebar-icon-btn"
        onClick={handleAddFolder}
        title={text.sidebar.addFolder}
      >
        <FolderOpen size={16} />
      </button>
      <div
        className={`sidebar-status-dot ${excelStatus.connected ? "connected" : "disconnected"}`}
        title={`Excel: ${excelStatus.connected ? "已连接" : "未连接"} | Word: ${wordStatus.connected ? "已连接" : "未连接"} | PPT: ${presentationStatus.connected ? "已连接" : "未连接"}`}
      />
      <button
        className={`sidebar-icon-btn ${currentPage === "settings" ? "active" : ""}`}
        onClick={() => openSettingsSection("general")}
        title={text.sidebar.settings}
      >
        <Settings size={16} />
      </button>
    </aside>
    );
  } else {

  // ---- 展开态 ----
  const hasSearchQuery = false;
  const compareText = (left: string, right: string) =>
    left.localeCompare(right, language === "zh-CN" ? "zh-CN" : "en", {
      numeric: true,
      sensitivity: "base",
    });
  const sortThreads = <T extends { preview?: string; updatedAt: number }>(items: T[], mode: SidebarSortMode) =>
    [...items].sort((a, b) => {
      if (mode === "recentAsc") return a.updatedAt - b.updatedAt;
      if (mode === "nameAsc") return compareText(a.preview || "", b.preview || "");
      if (mode === "nameDesc") return compareText(b.preview || "", a.preview || "");
      return b.updatedAt - a.updatedAt;
    });
  const ungroupedThreads = sortThreads(
    threads.filter((t) => !t.folderId),
    conversationSortMode
  );
  const groupedByFolder = pinnedFolders.map((folder) => ({
    folder,
    folderMatches: true,
    threads: sortThreads(
      threads.filter((t) => t.folderId === folder.path),
      projectSortMode
    ),
    files: folderFiles[folder.path] || [],
  })).filter(({ folderMatches, threads: folderThreads, files }) =>
    !hasSearchQuery || folderMatches || folderThreads.length > 0 || files.length > 0
  ).sort((a, b) => {
    if (projectSortMode === "recentAsc") return a.folder.addedAt - b.folder.addedAt;
    if (projectSortMode === "nameAsc") return compareText(a.folder.name, b.folder.name);
    if (projectSortMode === "nameDesc") return compareText(b.folder.name, a.folder.name);
    return b.folder.addedAt - a.folder.addedAt;
  });
  const hasProjectItems = groupedByFolder.length > 0;
  const hasConversationItems = ungroupedThreads.length > 0;
  const showNoSearchResults = hasSearchQuery && !hasProjectItems && !hasConversationItems;

  sidebarContent = (
    <aside className={`sidebar${isResizing ? " no-transition" : ""}`} ref={sidebarRef} style={{ width: isResizing || sidebarWidth !== 260 ? sidebarWidth : undefined }}>
      {/* 拖拽手柄 */}
      <div
        className={`sidebar-resize-handle${isResizing ? " resizing" : ""}`}
        onMouseDown={handleResizeStart}
      />
      {/* 顶部操作按钮 */}
      <div className="sidebar-primary-nav">
        <button
          className={`sidebar-primary-action${creatingNewThread ? " creating" : ""}`}
          onClick={handleCreateNewThread}
          disabled={creatingNewThread}
          title={text.sidebar.newThread}
        >
          {creatingNewThread ? <RefreshCw size={16} className="spin" /> : <PenLine size={16} />}
          <span>{text.sidebar.newThread}</span>
        </button>
        <button
          className={`sidebar-primary-action${searchOpen ? " active" : ""}`}
          onClick={handleToggleSearch}
          title={text.sidebar.search}
        >
          <Search size={16} />
          <span>{text.sidebar.search}</span>
        </button>
      </div>

      {/* 合并内容区：项目分组在上 + 普通对话在下 */}
      <div className="sidebar-content">
        <div className="sidebar-section-group">
          <div className="sidebar-section-header">
            <button
              className="sidebar-section-toggle"
              onClick={() => setProjectsExpanded((expanded) => !expanded)}
            >
              <span>{text.sidebar.projects}</span>
              {projectsExpanded ? <ChevronDown size={14} /> : <ChevronRight size={14} />}
            </button>
            <button
              className="sidebar-section-add"
              onClick={(e) => handleOpenSortMenu(e, "projects")}
              title={text.sidebar.sort}
            >
              <ClipboardList size={14} />
            </button>
            <button
              className="sidebar-section-add"
              onClick={handleAddFolder}
              title={text.sidebar.addFolder}
            >
              <Plus size={14} />
            </button>
          </div>
          {projectsExpanded && (
            <div className="sidebar-section-content">
              {groupedByFolder.map(({ folder, threads: folderThreads, files }) => (
                <FolderSection
                  key={folder.path}
                  folder={folder}
                  folderThreads={folderThreads}
                  files={files}
                  isExpanded={!!expandedFolders[folder.path]}
                  activeThreadId={activeThreadId}
                  runningThreadIds={runningThreadIds}
                  turnStatus={turnStatus}
                  creatingFolderThread={creatingFolderThread}
                  viewedThreadStatusAt={viewedThreadStatusAt}
                  language={language}
                  onToggleFolder={handleToggleFolder}
                  onCreateFolderThread={handleCreateFolderThread}
                  onRemoveFolder={removePinnedFolder}
                  onAddFile={handleAddFile}
                  onSwitchThread={handleSwitchThread}
                  onThreadContextMenu={handleThreadContextMenu}
                  fileContextMenu={fileContextMenu}
                  onFileContextMenu={handleFileContextMenu}
                  onFileContextMenuClose={() => setFileContextMenu(null)}
                  onTrashFile={handleTrashFile}
                  onOpenFile={handleOpenFile}
                  onCopyPath={handleCopyPath}
                  onRevealInExplorer={handleRevealInExplorer}
                  onPinFile={handlePinFile}
                />
              ))}
              {groupedByFolder.length === 0 && !hasSearchQuery && (
                <div className="sidebar-section-empty">{text.sidebar.noProjects}</div>
              )}
            </div>
          )}
        </div>

        <div className="sidebar-section-group">
          <div className="sidebar-section-header">
            <button
              className="sidebar-section-toggle"
              onClick={() => setConversationsExpanded((expanded) => !expanded)}
            >
              <span>{text.sidebar.conversations}</span>
              {conversationsExpanded ? <ChevronDown size={14} /> : <ChevronRight size={14} />}
            </button>
            <button
              className="sidebar-section-add"
              onClick={(e) => handleOpenSortMenu(e, "conversations")}
              title={text.sidebar.sort}
            >
              <ClipboardList size={14} />
            </button>
            <button
              className={`sidebar-section-add${creatingNewThread ? " creating" : ""}`}
              onClick={handleCreateNewThread}
              disabled={creatingNewThread}
              title={text.sidebar.newThread}
            >
              {creatingNewThread ? <RefreshCw size={14} className="spin" /> : <Plus size={14} />}
            </button>
          </div>
          {conversationsExpanded && (
            <div className="sidebar-section-content">
              <UngroupedThreadList
                threads={ungroupedThreads}
                activeThreadId={activeThreadId}
                runningThreadIds={runningThreadIds}
                turnStatus={turnStatus}
                viewedThreadStatusAt={viewedThreadStatusAt}
                language={language}
                onSwitchThread={handleSwitchThread}
                onThreadContextMenu={handleThreadContextMenu}
              />
            </div>
          )}
        </div>

        {/* 无内容占位 */}
        {showNoSearchResults ? (
          <div className="sidebar-empty">{text.sidebar.noSearchResults}</div>
        ) : (!hasSearchQuery && !hasProjectItems && !hasConversationItems && (
          <div className="sidebar-empty">{text.sidebar.noThreads}</div>
        ))}
      </div>

      {/* 右键菜单 */}
      {sortMenu && (
        <div
          className="sidebar-sort-menu"
          style={{ left: sortMenu.x, top: sortMenu.y }}
          data-section={sortMenu.section}
          onClick={(e) => e.stopPropagation()}
        >
          {([
            ["recentDesc", text.sidebar.sortRecentDesc],
            ["recentAsc", text.sidebar.sortRecentAsc],
            ["nameAsc", text.sidebar.sortNameAsc],
            ["nameDesc", text.sidebar.sortNameDesc],
          ] as const).map(([mode, label]) => {
            const activeMode = sortMenu.section === "projects" ? projectSortMode : conversationSortMode;
            return (
              <button
                key={mode}
                className={`sidebar-sort-menu-item${activeMode === mode ? " active" : ""}`}
                onClick={() => handleSelectSortMode(sortMenu.section, mode)}
              >
                <Clock size={14} />
                <span>{label}</span>
                {activeMode === mode && <Check size={14} />}
              </button>
            );
          })}
        </div>
      )}
      {contextMenu && (
        <ThreadContextMenu
          contextMenu={contextMenu}
          pinnedFolders={pinnedFolders}
          language={language}
          onConfirmDelete={handleConfirmDelete}
          onMoveToFolder={handleMoveToFolder}
          onPinThread={handlePinThread}
          onRenameThread={handleRenameThread}
          onClose={() => setContextMenu(null)}
          setShowMoveMenu={(show) => setContextMenu((m) => m ? { ...m, showMoveMenu: show } : m)}
          setConfirming={(confirming) => setContextMenu((m) => m ? { ...m, confirming } : m)}
        />
      )}

      {/* 底部：连接状态 + 设置 */}
      <div className="sidebar-footer">
        <div className={`sidebar-connection ${connectFailed ? "connection-failed" : ""}`}>
          <div className="sidebar-connection-apps">
            {/* Excel 连接 */}
            <div className={`connection-indicator ${excelStatus.connected ? "connected" : "disconnected"}`}
                 title={excelStatus.connected ? `Excel ${excelStatus.version || ""}` : "Excel 未连接"}>
              <span className={`connection-dot ${pulseDot ? "pulse" : ""}`} />
              <span className="connection-text">
                {excelStatus.connected
                  ? `${text.sidebar.connectedExcelPrefix}${excelStatus.version ? ` (${excelStatus.version})` : ""}`
                  : text.sidebar.excelDisconnected}
              </span>
            </div>
            {/* Word 连接 */}
            <div className={`connection-indicator ${wordStatus.connected ? "connected" : "disconnected"}`}
                 title={wordStatus.connected ? `Word ${wordStatus.version || ""}` : "Word 未连接"}>
              <span className="connection-dot" />
              <span className="connection-text">
                {wordStatus.connected
                  ? `${text.sidebar.connectedWordPrefix}${wordStatus.version ? ` (${wordStatus.version})` : ""}`
                  : text.sidebar.wordDisconnected}
              </span>
            </div>
            {/* PowerPoint 连接 */}
            <div className={`connection-indicator ${presentationStatus.connected ? "connected" : "disconnected"}`}
                 title={presentationStatus.connected ? `PowerPoint ${presentationStatus.version || ""}` : "PowerPoint 未连接"}>
              <span className="connection-dot" />
              <span className="connection-text">
                {presentationStatus.connected
                  ? `${text.sidebar.connectedPresentationPrefix}${presentationStatus.version ? ` (${presentationStatus.version})` : ""}`
                  : text.sidebar.presentationDisconnected}
              </span>
            </div>
          </div>
          {excelStatus.connected ? (
            <button
              className="btn-connect btn-connect-icon"
              onClick={handleConnect}
              title={text.sidebar.reconnect}
              disabled={connecting}
            >
              <RefreshCw size={10} className={connecting ? "spin" : ""} />
            </button>
          ) : (
            <button className="btn-connect" onClick={handleConnect} disabled={connecting}>
              {connecting ? text.sidebar.connecting : text.sidebar.connect}
            </button>
          )}
        </div>

        {/* 设置按钮 */}
        <div className="sidebar-settings-menu-wrap">
          <button
            className={`sidebar-nav-btn ${currentPage === "settings" ? "active" : ""}`}
            onClick={(e) => {
              e.stopPropagation();
              setSettingsMenuOpen((open) => !open);
            }}
          >
            <Settings size={16} />
            <span>{text.sidebar.settings}</span>
          </button>

          {settingsMenuOpen && (
            <div className="sidebar-settings-menu" onClick={(e) => e.stopPropagation()}>
              <div className="sidebar-settings-account">
                <User size={16} />
                <span>{text.sidebar.localAccount}</span>
              </div>
              <div className="sidebar-settings-divider" />
              <button className="sidebar-settings-menu-item" onClick={() => openSettingsSection("profile")}>
                <User size={16} />
                <span>{text.sidebar.profile}</span>
              </button>
              <button className="sidebar-settings-menu-item" onClick={() => openSettingsSection("general")}>
                <Settings size={16} />
                <span>{text.sidebar.settings}</span>
              </button>
              <button className="sidebar-settings-menu-item" onClick={() => openSettingsSection("opensource")}>
                <Package size={16} />
                <span>{text.sidebar.openSource}</span>
              </button>
              <button className="sidebar-settings-menu-item" onClick={() => setSettingsMenuOpen(false)}>
                <LogOut size={16} />
                <span>{text.sidebar.logout}</span>
              </button>
            </div>
          )}
        </div>
      </div>
    </aside>
  );
  }

  // ---- 宿主选择弹窗 ----
  const hostSelectionDialog = pendingHosts && pendingHosts.length > 1 ? (
    <HostSelectionDialog
      availableHosts={pendingHosts}
      onSelect={handleSelectHost}
      onDismiss={() => setPendingHosts(null)}
    />
  ) : null;

  return (
    <>
      {sidebarContent}
      <SidebarSearchPalette
        open={searchOpen}
        threads={threads}
        folders={pinnedFolders}
        folderFiles={folderFiles}
        language={language}
        activeThreadId={activeThreadId}
        onClose={() => setSearchOpen(false)}
        onSwitchThread={handleSwitchThread}
        onAddFile={handleAddFile}
        onCreateNewThread={handleCreateNewThread}
        onAddFolder={handleAddFolder}
        onOpenSettings={() => openSettingsSection("general")}
      />
      {hostSelectionDialog}
    </>
  );
};
