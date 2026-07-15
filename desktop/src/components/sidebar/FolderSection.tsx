/**
 * FolderSection — 侧边栏文件夹分组渲染 + 未分组会话列表
 *
 * 从 Sidebar.tsx 提取，包含：
 * - FolderSection: 文件夹头（展开/折叠、新建会话、查看文件、删除文件夹）
 * - UngroupedThreadList: 未分组会话列表
 */

import React, { useEffect, useState } from "react";
import type { AppLanguage, PinnedFolder } from "../../store/settingsStore";
import type { FolderFileInfo } from "../../electronApi";
import { getAppText } from "../../i18n";
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
import { SidebarThreadItem } from "./SidebarThreadItem";

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
  folderActions: FolderSectionActions;
  threadActions: FolderSectionThreadActions;
  fileMenuApi: FolderSectionFileMenuApi;
}

export interface FolderSectionActions {
  toggle: (folderPath: string) => void;
  createThread: (folderPath: string) => void;
  remove: (folderPath: string) => void;
}

export interface FolderSectionThreadActions {
  switchThread: (threadId: string) => void;
  openContextMenu: (e: React.MouseEvent, threadId: string, inFolder?: boolean) => void;
}

export interface FolderSectionFileMenuApi {
  state: FileContextMenuState | null;
  addFile: (file: FolderFileInfo) => void;
  openContextMenu: (e: React.MouseEvent, file: FolderFileInfo, isPinned: boolean) => void;
  close: () => void;
  trashFile: (filePath: string) => void;
  openFile: (filePath: string) => void;
  copyPath: (filePath: string) => void;
  revealInExplorer: (filePath: string) => void;
  pinFile: (filePath: string) => void;
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
  folderActions,
  threadActions,
  fileMenuApi,
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
        <button className="sidebar-folder-toggle" onClick={() => folderActions.toggle(folder.path)}>
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
            onClick={() => folderActions.createThread(folder.path)}
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
            onClick={() => folderActions.remove(folder.path)}
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
                  onClick={() => fileMenuApi.addFile(file)}
                  onContextMenu={(e) =>
                    fileMenuApi.openContextMenu(e, file, pinnedFiles.includes(file.filePath))
                  }
                  title={text.sidebar.addFileToChat}
                >
                  <FileSpreadsheet size={13} />
                  <span className="sidebar-file-name">{file.fileName}</span>
                </button>
              ))
            )
          ) : folderThreads.length === 0 ? (
            <div className="sidebar-folder-empty">{text.sidebar.noFolderThreads}</div>
          ) : (
            folderThreads.map((thread) => (
              <SidebarThreadItem
                key={thread.threadId}
                thread={thread}
                activeThreadId={activeThreadId}
                runningThreadIds={runningThreadIds}
                turnStatus={turnStatus}
                viewedThreadStatusAt={viewedThreadStatusAt}
                language={language}
                fallbackTitle={text.sidebar.newChat}
                inFolder
                onSwitchThread={threadActions.switchThread}
                onThreadContextMenu={threadActions.openContextMenu}
              />
            ))
          )}
        </div>
      )}

      {/* 文件右键菜单 */}
      {fileMenuApi.state && (
        <FileContextMenu
          state={fileMenuApi.state}
          language={language}
          onClose={fileMenuApi.close}
          onDelete={fileMenuApi.trashFile}
          onOpen={fileMenuApi.openFile}
          onCopyPath={fileMenuApi.copyPath}
          onRevealInExplorer={fileMenuApi.revealInExplorer}
          onPinFile={fileMenuApi.pinFile}
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
      {threads.map((thread) => (
        <SidebarThreadItem
          key={thread.threadId}
          thread={thread}
          activeThreadId={activeThreadId}
          runningThreadIds={runningThreadIds}
          turnStatus={turnStatus}
          viewedThreadStatusAt={viewedThreadStatusAt}
          language={language}
          fallbackTitle={text.sidebar.newChat}
          onSwitchThread={onSwitchThread}
          onThreadContextMenu={onThreadContextMenu}
        />
      ))}
    </>
  );
}
