/**
 * FolderSection — 侧边栏文件夹分组渲染 + 未分组会话列表
 *
 * 从 Sidebar.tsx 提取，包含：
 * - FolderSection: 文件夹头（展开/折叠、新建会话、查看文件、删除文件夹）
 * - UngroupedThreadList: 未分组会话列表
 */

import React, { useEffect, useState } from "react";
import { useSettingsStore, type AppLanguage, type PinnedFolder } from "../../store/settingsStore";
import type { FolderFileInfo } from "../../electronApi";
import { getAppText } from "../../i18n";
import { formatTime, getThreadDisplayStatus, getThreadStatusLabel } from "../../utils/sidebarHelpers";
import {
  Plus,
  RefreshCw,
  ListPlus,
  Trash2,
  ChevronDown,
  ChevronRight,
  FolderOpen,
  FileSpreadsheet,
  MessageSquare,
} from "../common/IconMap";
import { FileContextMenu, type FileContextMenuState } from "./FileContextMenu";

interface FolderThread {
  threadId: string;
  preview?: string;
  updatedAt: number;
  lastTurnStatus?: string;
  folderId?: string;
}

interface FolderSectionProps {
  folder: PinnedFolder;
  folderThreads: FolderThread[];
  files: FolderFileInfo[];
  isExpanded: boolean;
  activeThreadId: string | null;
  runningThreadIds: Record<string, boolean>;
  turnStatus: string;
  creatingFolderThread: string | null;
  viewedThreadStatusAt: Record<string, number>;
  language: AppLanguage;
  onToggleFolder: (folderPath: string) => void;
  onCreateFolderThread: (folderPath: string) => void;
  onRemoveFolder: (folderPath: string) => void;
  onAddFile: (file: FolderFileInfo) => void;
  onSwitchThread: (threadId: string) => void;
  onThreadContextMenu: (e: React.MouseEvent, threadId: string, inFolder?: boolean) => void;
  /** 文件右键菜单状态 */
  fileContextMenu: FileContextMenuState | null;
  onFileContextMenu: (e: React.MouseEvent, file: FolderFileInfo, isPinned: boolean) => void;
  onFileContextMenuClose: () => void;
  onTrashFile: (filePath: string) => void;
  onOpenFile: (filePath: string) => void;
  onCopyPath: (filePath: string) => void;
  onRevealInExplorer: (filePath: string) => void;
  onPinFile: (filePath: string) => void;
}

export function FolderSection({
  folder,
  folderThreads,
  files,
  isExpanded,
  activeThreadId,
  runningThreadIds,
  turnStatus,
  creatingFolderThread,
  viewedThreadStatusAt,
  language,
  onToggleFolder,
  onCreateFolderThread,
  onRemoveFolder,
  onAddFile,
  onSwitchThread,
  onThreadContextMenu,
  fileContextMenu,
  onFileContextMenu,
  onFileContextMenuClose,
  onTrashFile,
  onOpenFile,
  onCopyPath,
  onRevealInExplorer,
  onPinFile,
}: FolderSectionProps) {
  const text = getAppText(language);
  const pinnedFiles = folder.pinnedFiles || [];
  const [showFileList, setShowFileList] = useState(false);

  useEffect(() => {
    if (!isExpanded) setShowFileList(false);
  }, [isExpanded]);

  // 排序：置顶文件在前，其余按名称排序
  const sortedFiles = [...files].sort((a, b) => {
    const aPinned = pinnedFiles.indexOf(a.filePath);
    const bPinned = pinnedFiles.indexOf(b.filePath);
    if (aPinned >= 0 && bPinned >= 0) return aPinned - bPinned;
    if (aPinned >= 0) return -1;
    if (bPinned >= 0) return 1;
    return a.fileName.localeCompare(b.fileName);
  });

  return (
    <div className="sidebar-folder-section">
      <div className="sidebar-folder-header">
        <button
          className="sidebar-folder-toggle"
          onClick={() => onToggleFolder(folder.path)}
        >
          {isExpanded ? <ChevronDown size={14} /> : <ChevronRight size={14} />}
          <FolderOpen size={14} />
          <span className="sidebar-folder-name">{folder.name}</span>
          {folderThreads.length > 0 && (
            <span className="sidebar-folder-count">{folderThreads.length}</span>
          )}
        </button>
        <div className="sidebar-folder-actions">
          <button
            className={`sidebar-folder-action-btn${creatingFolderThread === folder.path ? " creating" : ""}`}
            onClick={() => onCreateFolderThread(folder.path)}
            title={text.sidebar.newThreadInFolder}
          >
            {creatingFolderThread === folder.path ? (
              <RefreshCw size={12} className="spin" />
            ) : (
              <Plus size={12} />
            )}
          </button>
          {isExpanded && (
            <button
              className={`sidebar-folder-action-btn${showFileList ? " active" : ""}`}
              onClick={() => setShowFileList((visible) => !visible)}
              title={showFileList ? text.sidebar.viewFolderChats : text.sidebar.viewFolderFiles}
            >
              {showFileList ? <MessageSquare size={12} /> : <ListPlus size={12} />}
            </button>
          )}
          <button
            className="sidebar-folder-action-btn danger"
            onClick={() => onRemoveFolder(folder.path)}
            title={text.sidebar.removeFolder}
          >
            <Trash2 size={12} />
          </button>
        </div>
      </div>
      {isExpanded && (
        <div className="sidebar-folder-content">
          {showFileList ? (
            sortedFiles.length === 0 ? (
              <div className="sidebar-folder-empty">{text.sidebar.noExcelFiles}</div>
            ) : (
              sortedFiles.map((file) => (
                <button
                  key={file.filePath}
                  className="sidebar-file-item"
                  onClick={() => onAddFile(file)}
                  onContextMenu={(e) => onFileContextMenu(e, file, pinnedFiles.includes(file.filePath))}
                  title={text.sidebar.addFileToChat}
                >
                  <FileSpreadsheet size={13} />
                  <span className="sidebar-file-name">{file.fileName}</span>
                </button>
              ))
            )
          ) : (
            folderThreads.length === 0 ? (
              <div className="sidebar-folder-empty">{text.sidebar.noFolderThreads}</div>
            ) : (
              folderThreads.map((thread) => {
                const isActiveThread = activeThreadId === thread.threadId;
                const isRunningThread = Boolean(runningThreadIds[thread.threadId]);
                const statusViewed = isActiveThread || viewedThreadStatusAt[thread.threadId] === thread.updatedAt;
                const status = getThreadDisplayStatus(
                  thread.lastTurnStatus,
                  isRunningThread ? "in_progress" : isActiveThread ? turnStatus as any : undefined,
                  statusViewed
                );
                const statusLabel = getThreadStatusLabel(status, language);
                return (
                  <div
                    key={thread.threadId}
                    className={`sidebar-thread-item sidebar-thread-in-folder ${activeThreadId === thread.threadId ? "active" : ""}`}
                    onClick={() => onSwitchThread(thread.threadId)}
                    onContextMenu={(e) => onThreadContextMenu(e, thread.threadId, true)}
                  >
                    <div className="thread-item-main">
                      <MessageSquare size={12} className="thread-item-icon" />
                      <div className="thread-item-preview">
                        {thread.preview || text.sidebar.newChat}
                      </div>
                      <span className="thread-item-time">
                        {formatTime(thread.updatedAt, language)}
                      </span>
                      {status && (
                        <span
                          className={`thread-status-indicator ${status}`}
                          title={statusLabel}
                          aria-label={statusLabel}
                        >
                          {status === "running" ? <RefreshCw size={12} className="spin" /> : null}
                        </span>
                      )}
                    </div>
                  </div>
                );
              })
            )
          )}
        </div>
      )}

      {/* 文件右键菜单 */}
      {fileContextMenu && (
        <FileContextMenu
          state={fileContextMenu}
          language={language}
          onClose={onFileContextMenuClose}
          onDelete={onTrashFile}
          onOpen={onOpenFile}
          onCopyPath={onCopyPath}
          onRevealInExplorer={onRevealInExplorer}
          onPinFile={onPinFile}
        />
      )}
    </div>
  );
}

// ============================================================
// UngroupedThreadList — 未分组会话列表
// ============================================================

interface ThreadItem {
  threadId: string;
  preview?: string;
  updatedAt: number;
  lastTurnStatus?: string;
  folderId?: string;
}

interface UngroupedThreadListProps {
  threads: ThreadItem[];
  activeThreadId: string | null;
  runningThreadIds: Record<string, boolean>;
  turnStatus: string;
  viewedThreadStatusAt: Record<string, number>;
  language: AppLanguage;
  onSwitchThread: (threadId: string) => void;
  onThreadContextMenu: (e: React.MouseEvent, threadId: string, inFolder?: boolean) => void;
}

export function UngroupedThreadList({
  threads,
  activeThreadId,
  runningThreadIds,
  turnStatus,
  viewedThreadStatusAt,
  language,
  onSwitchThread,
  onThreadContextMenu,
}: UngroupedThreadListProps) {
  const text = getAppText(language);

  return (
    <>
      {threads.map((thread) => {
        const isActiveThread = activeThreadId === thread.threadId;
        const isRunningThread = Boolean(runningThreadIds[thread.threadId]);
        const statusViewed = isActiveThread || viewedThreadStatusAt[thread.threadId] === thread.updatedAt;
        const status = getThreadDisplayStatus(
          thread.lastTurnStatus,
          isRunningThread ? "in_progress" : isActiveThread ? turnStatus as any : undefined,
          statusViewed
        );
        const statusLabel = getThreadStatusLabel(status, language);
        return (
          <div
            key={thread.threadId}
            className={`sidebar-thread-item ${activeThreadId === thread.threadId ? "active" : ""}`}
            onClick={() => onSwitchThread(thread.threadId)}
            onContextMenu={(e) => onThreadContextMenu(e, thread.threadId)}
          >
            <div className="thread-item-main">
              <div className="thread-item-preview">
                {thread.preview || text.sidebar.newChat}
              </div>
              <span className="thread-item-time">
                {formatTime(thread.updatedAt, language)}
              </span>
              {status && (
                <span
                  className={`thread-status-indicator ${status}`}
                  title={statusLabel}
                  aria-label={statusLabel}
                >
                  {status === "running" ? <RefreshCw size={12} className="spin" /> : null}
                </span>
              )}
            </div>
          </div>
        );
      })}
    </>
  );
}
