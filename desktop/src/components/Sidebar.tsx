/**
 * 侧边栏 — 会话管理 + 意图快捷 + 导航 + 连接状态
 *
 * 已拆分模块：
 * - utils/sidebarHelpers.ts: 意图常量、时间格式化、状态判断
 * - hooks/useExcelConnection.ts: Excel/WPS 连接状态管理
 * - components/sidebar/FolderSection.tsx: 文件夹分组渲染
 * - components/sidebar/ThreadContextMenu.tsx: 右键菜单
 * - components/sidebar/SidebarCollapsed.tsx: 折叠态渲染
 * - components/sidebar/SidebarExpanded.tsx: 展开态渲染
 */

import React, { useEffect, useState, useCallback, useMemo, useRef } from "react";
import { useChatStore } from "../store/chatStore";
import { useSettingsStore } from "../store/settingsStore";
import type { AppPage } from "../App";
import type { SettingsSection } from "./SettingsPage";
import type { FolderFileInfo } from "../electronApi";
import { getAppText } from "../i18n";
import { ipcApi } from "../services/ipcApi";
import type { IntentKind } from "../utils/sidebarHelpers";
import { useExcelConnection } from "../hooks/useExcelConnection";
import { useOfficeConnection } from "../hooks/useOfficeConnection";
import { useDocumentDismiss } from "../hooks/useDocumentDismiss";
import { HostSelectionDialog } from "./excel/HostSelectionDialog";
import type { ContextMenuState } from "./sidebar/ThreadContextMenu";
import type { FileContextMenuState } from "./sidebar/FileContextMenu";
import { SidebarSearchPalette } from "./sidebar/SidebarSearchPalette";
import { SidebarCollapsed } from "./sidebar/SidebarCollapsed";
import {
  SidebarExpanded,
  type SidebarSortMode,
  type SidebarSortSection,
} from "./sidebar/SidebarExpanded";

export type { IntentKind } from "../utils/sidebarHelpers";

interface SidebarProps {
  collapsed: boolean;
  currentPage: AppPage;
  onNavigate: (page: AppPage) => void;
  onOpenSettingsSection?: (section: SettingsSection) => void;
  activeIntent: IntentKind;
  onIntentClick: (intent: IntentKind) => void;
}

function compareSidebarText(left: string, right: string, language: string): number {
  return left.localeCompare(right, language === "zh-CN" ? "zh-CN" : "en", {
    numeric: true,
    sensitivity: "base",
  });
}

function sortSidebarItems<T extends { preview?: string; updatedAt: number }>(
  items: T[],
  mode: SidebarSortMode,
  language: string
): T[] {
  return [...items].sort((a, b) => {
    if (mode === "recentAsc") return a.updatedAt - b.updatedAt;
    if (mode === "nameAsc") return compareSidebarText(a.preview || "", b.preview || "", language);
    if (mode === "nameDesc") return compareSidebarText(b.preview || "", a.preview || "", language);
    return b.updatedAt - a.updatedAt;
  });
}

export const Sidebar: React.FC<SidebarProps> = ({
  collapsed,
  currentPage,
  onNavigate,
  onOpenSettingsSection,
}) => {
  const { threads, activeThreadId, runningThreadIds, turnStatus, loadThreads, switchThread, createNewThread, deleteThread, moveThreadToFolder, addFilesToComposer } =
    useChatStore();
  const { language, pinnedFolders, addPinnedFolder, removePinnedFolder, updatePinnedFolder } = useSettingsStore();
  const text = getAppText(language);

  const [folderFiles, setFolderFiles] = useState<Record<string, FolderFileInfo[]>>({});
  const [expandedFolders, setExpandedFolders] = useState<Record<string, boolean>>({});
  const [creatingFolderThread, setCreatingFolderThread] = useState<string | null>(null);
  const [creatingNewThread, setCreatingNewThread] = useState(false);
  const [viewedThreadStatusAt, setViewedThreadStatusAt] = useState<Record<string, number>>({});
  const initializedViewedStatuses = useRef(false);
  const [settingsMenuOpen, setSettingsMenuOpen] = useState(false);
  const [searchOpen, setSearchOpen] = useState(false);
  const [projectsExpanded, setProjectsExpanded] = useState(true);
  const [conversationsExpanded, setConversationsExpanded] = useState(true);
  const [sortMenu, setSortMenu] = useState<{ section: SidebarSortSection; x: number; y: number } | null>(null);
  const [projectSortMode, setProjectSortMode] = useState<SidebarSortMode>("recentDesc");
  const [conversationSortMode, setConversationSortMode] = useState<SidebarSortMode>("recentDesc");
  const [contextMenu, setContextMenu] = useState<ContextMenuState | null>(null);
  const [fileContextMenu, setFileContextMenu] = useState<FileContextMenuState | null>(null);
  const [sidebarWidth, setSidebarWidth] = useState(260);
  const [isResizing, setIsResizing] = useState(false);
  const resizingRef = useRef(false);
  const sidebarRef = useRef<HTMLElement>(null);

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
  const { wordStatus, presentationStatus } = useOfficeConnection();

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
    } catch {
      // ignore
    }
  }, [addPinnedFolder]);

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

  const handleResizeStart = useCallback((e: React.MouseEvent) => {
    e.preventDefault();
    resizingRef.current = true;
    setIsResizing(true);
    const startX = e.clientX;
    const startWidth = sidebarWidth;
    const handleMouseMove = (event: MouseEvent) => {
      if (!resizingRef.current) return;
      setSidebarWidth(Math.min(400, Math.max(180, startWidth + event.clientX - startX)));
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
    if (activeThreadId) markThreadViewed(activeThreadId);
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

  const handleFileContextMenu = useCallback((e: React.MouseEvent, file: FolderFileInfo, isPinned: boolean) => {
    e.preventDefault();
    e.stopPropagation();
    setContextMenu(null);
    setFileContextMenu({ file, x: e.clientX, y: e.clientY, isPinned });
  }, []);

  const refreshFolderContainingFile = useCallback(async (filePath: string) => {
    const folderPath = Object.keys(folderFiles).find((fp) =>
      folderFiles[fp].some((file) => file.filePath === filePath)
    );
    if (!folderPath) return;
    const files = await ipcApi.folder.listFiles(folderPath);
    setFolderFiles((prev) => ({ ...prev, [folderPath]: files }));
  }, [folderFiles]);

  const handleTrashFile = useCallback(async (filePath: string) => {
    setFileContextMenu(null);
    await ipcApi.file.trashFile(filePath);
    await refreshFolderContainingFile(filePath);
  }, [refreshFolderContainingFile]);

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
    const folderPath = Object.keys(folderFiles).find((fp) =>
      folderFiles[fp].some((file) => file.filePath === filePath)
    );
    const folder = folderPath ? pinnedFolders.find((item) => item.path === folderPath) : undefined;
    if (!folder || !folderPath) return;
    const pinned = folder.pinnedFiles || [];
    const nextPinned = pinned.includes(filePath)
      ? pinned.filter((item) => item !== filePath)
      : [...pinned, filePath];
    updatePinnedFolder(folderPath, { pinnedFiles: nextPinned });
  }, [folderFiles, pinnedFolders, updatePinnedFolder]);

  useDocumentDismiss({ active: contextMenu !== null, onDismiss: () => setContextMenu(null) });
  useDocumentDismiss({ active: fileContextMenu !== null, onDismiss: () => setFileContextMenu(null) });
  useDocumentDismiss({ active: settingsMenuOpen, onDismiss: () => setSettingsMenuOpen(false) });
  useDocumentDismiss({ active: sortMenu !== null, onDismiss: () => setSortMenu(null) });

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

  const hasSearchQuery = false;
  const ungroupedThreads = useMemo(() => sortSidebarItems(
    threads.filter((thread) => !thread.folderId),
    conversationSortMode,
    language
  ), [conversationSortMode, language, threads]);
  const groupedByFolder = useMemo(() => pinnedFolders.map((folder) => ({
    folder,
    folderMatches: true,
    threads: sortSidebarItems(
      threads.filter((thread) => thread.folderId === folder.path),
      projectSortMode,
      language
    ),
    files: folderFiles[folder.path] || [],
  })).filter(({ folderMatches, threads: folderThreads, files }) =>
    !hasSearchQuery || folderMatches || folderThreads.length > 0 || files.length > 0
  ).sort((a, b) => {
    if (projectSortMode === "recentAsc") return a.folder.addedAt - b.folder.addedAt;
    if (projectSortMode === "nameAsc") return compareSidebarText(a.folder.name, b.folder.name, language);
    if (projectSortMode === "nameDesc") return compareSidebarText(b.folder.name, a.folder.name, language);
    return b.folder.addedAt - a.folder.addedAt;
  }), [folderFiles, hasSearchQuery, language, pinnedFolders, projectSortMode, threads]);
  const hasProjectItems = groupedByFolder.length > 0;
  const hasConversationItems = ungroupedThreads.length > 0;
  const showNoSearchResults = hasSearchQuery && !hasProjectItems && !hasConversationItems;

  const hostSelectionDialog = pendingHosts && pendingHosts.length > 1 ? (
    <HostSelectionDialog
      availableHosts={pendingHosts}
      onSelect={handleSelectHost}
      onDismiss={() => setPendingHosts(null)}
    />
  ) : null;

  const sidebarContent = collapsed ? (
    <SidebarCollapsed
      currentPage={currentPage}
      text={text}
      creatingNewThread={creatingNewThread}
      excelStatus={excelStatus}
      wordStatus={wordStatus}
      presentationStatus={presentationStatus}
      onCreateNewThread={handleCreateNewThread}
      onToggleSearch={() => setSearchOpen((open) => !open)}
      onAddFolder={handleAddFolder}
      onOpenSettings={() => openSettingsSection("general")}
    />
  ) : (
    <SidebarExpanded
      currentPage={currentPage}
      text={text}
      language={language}
      sidebarRef={sidebarRef}
      isResizing={isResizing}
      sidebarWidth={sidebarWidth}
      projectsExpanded={projectsExpanded}
      conversationsExpanded={conversationsExpanded}
      groupedByFolder={groupedByFolder}
      ungroupedThreads={ungroupedThreads}
      hasSearchQuery={hasSearchQuery}
      hasProjectItems={hasProjectItems}
      hasConversationItems={hasConversationItems}
      showNoSearchResults={showNoSearchResults}
      activeThreadId={activeThreadId}
      runningThreadIds={runningThreadIds}
      turnStatus={turnStatus}
      creatingNewThread={creatingNewThread}
      creatingFolderThread={creatingFolderThread}
      viewedThreadStatusAt={viewedThreadStatusAt}
      expandedFolders={expandedFolders}
      fileContextMenu={fileContextMenu}
      contextMenu={contextMenu}
      pinnedFolders={pinnedFolders}
      sortMenu={sortMenu}
      projectSortMode={projectSortMode}
      conversationSortMode={conversationSortMode}
      settingsMenuOpen={settingsMenuOpen}
      searchOpen={searchOpen}
      excelStatus={excelStatus}
      wordStatus={wordStatus}
      presentationStatus={presentationStatus}
      connectFailed={connectFailed}
      connecting={connecting}
      pulseDot={pulseDot}
      onResizeStart={handleResizeStart}
      onCreateNewThread={handleCreateNewThread}
      onToggleSearch={() => setSearchOpen((open) => !open)}
      onToggleProjectsExpanded={() => setProjectsExpanded((expanded) => !expanded)}
      onToggleConversationsExpanded={() => setConversationsExpanded((expanded) => !expanded)}
      onOpenSortMenu={handleOpenSortMenu}
      onAddFolder={handleAddFolder}
      onToggleFolder={handleToggleFolder}
      onCreateFolderThread={handleCreateFolderThread}
      onRemoveFolder={removePinnedFolder}
      onAddFile={handleAddFile}
      onSwitchThread={handleSwitchThread}
      onThreadContextMenu={handleThreadContextMenu}
      onFileContextMenu={handleFileContextMenu}
      onFileContextMenuClose={() => setFileContextMenu(null)}
      onTrashFile={handleTrashFile}
      onOpenFile={handleOpenFile}
      onCopyPath={handleCopyPath}
      onRevealInExplorer={handleRevealInExplorer}
      onPinFile={handlePinFile}
      onSelectSortMode={handleSelectSortMode}
      onConfirmDelete={handleConfirmDelete}
      onMoveToFolder={handleMoveToFolder}
      onPinThread={() => setContextMenu(null)}
      onRenameThread={() => setContextMenu(null)}
      onCloseContextMenu={() => setContextMenu(null)}
      onSetContextMoveMenu={(show) => setContextMenu((menu) => menu ? { ...menu, showMoveMenu: show } : menu)}
      onSetContextConfirming={(confirming) => setContextMenu((menu) => menu ? { ...menu, confirming } : menu)}
      onConnect={handleConnect}
      onToggleSettingsMenu={(event) => {
        event.stopPropagation();
        setSettingsMenuOpen((open) => !open);
      }}
      onOpenSettingsSection={openSettingsSection}
      onCloseSettingsMenu={() => setSettingsMenuOpen(false)}
    />
  );

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
