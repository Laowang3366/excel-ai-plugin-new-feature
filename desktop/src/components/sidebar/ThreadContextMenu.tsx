/**
 * ThreadContextMenu — 会话右键菜单
 *
 * 从 Sidebar.tsx 提取，包含：
 * - 置顶/取消置顶
 * - 重命名
 * - 移动到文件夹 / 从文件夹移出
 * - 删除确认流程
 */

import React, { useState } from "react";
import { useSettingsStore, type PinnedFolder, type AppLanguage } from "../../store/settingsStore";
import { getAppText } from "../../i18n";
import {
  Trash2,
  FolderOpen,
  ChevronLeft,
  ChevronRight,
  Pin,
  Pencil,
} from "../common/IconMap";

export interface ContextMenuState {
  threadId: string;
  x: number;
  y: number;
  confirming: boolean;
  showMoveMenu?: boolean;
  /** 当前会话是否在文件夹中 */
  inFolder?: boolean;
}

interface ThreadContextMenuProps {
  contextMenu: ContextMenuState;
  pinnedFolders: PinnedFolder[];
  language: AppLanguage;
  onConfirmDelete: (threadId: string) => void;
  onMoveToFolder: (threadId: string, folderId?: string) => void;
  onPinThread: (threadId: string) => void;
  onRenameThread: (threadId: string) => void;
  onClose: () => void;
  setShowMoveMenu: (show: boolean) => void;
  setConfirming: (confirming: boolean) => void;
}

export function ThreadContextMenu({
  contextMenu,
  pinnedFolders,
  language,
  onConfirmDelete,
  onMoveToFolder,
  onPinThread,
  onRenameThread,
  onClose,
  setShowMoveMenu,
  setConfirming,
}: ThreadContextMenuProps) {
  const text = getAppText(language);

  return (
    <div
      className="thread-context-menu"
      style={{ left: contextMenu.x, top: contextMenu.y }}
      onClick={(e) => e.stopPropagation()}
      onContextMenu={(e) => e.preventDefault()}
    >
      {contextMenu.confirming ? (
        <>
          <div className="thread-context-menu-title">{text.sidebar.deleteConfirm}</div>
          <div className="thread-context-menu-actions">
            <button
              className="thread-context-confirm danger"
              onClick={() => onConfirmDelete(contextMenu.threadId)}
            >
              {text.sidebar.delete}
            </button>
            <button
              className="thread-context-confirm"
              onClick={onClose}
            >
              {text.sidebar.cancel}
            </button>
          </div>
        </>
      ) : contextMenu.showMoveMenu ? (
        <>
          <button
            className="thread-context-menu-item"
            onClick={() => setShowMoveMenu(false)}
          >
            <ChevronLeft size={14} />
            <span>{text.sidebar.moveToFolder}</span>
          </button>
          <div className="thread-context-menu-divider" />
          {pinnedFolders.length === 0 ? (
            <div className="thread-context-menu-item disabled">
              <span>{text.sidebar.addFolderFirst}</span>
            </div>
          ) : (
            <>
              <button
                className="thread-context-menu-item"
                onClick={() => onMoveToFolder(contextMenu.threadId)}
              >
                <span>{text.sidebar.noFolder}</span>
              </button>
              {pinnedFolders.map((folder) => (
                <button
                  key={folder.path}
                  className="thread-context-menu-item"
                  onClick={() => onMoveToFolder(contextMenu.threadId, folder.path)}
                >
                  <FolderOpen size={14} />
                  <span>{folder.name}</span>
                </button>
              ))}
            </>
          )}
        </>
      ) : (
        <>
          <button
            className="thread-context-menu-item"
            onClick={() => onPinThread(contextMenu.threadId)}
          >
            <Pin size={14} />
            <span>{text.sidebar.pinToTop}</span>
          </button>
          <button
            className="thread-context-menu-item"
            onClick={() => onRenameThread(contextMenu.threadId)}
          >
            <Pencil size={14} />
            <span>{text.sidebar.rename}</span>
          </button>
          {contextMenu.inFolder ? (
            <>
              <div className="thread-context-menu-divider" />
              <button
                className="thread-context-menu-item"
                onClick={() => onMoveToFolder(contextMenu.threadId)}
              >
                <FolderOpen size={14} />
                <span>{text.sidebar.removeFromFolder}</span>
              </button>
            </>
          ) : (
            pinnedFolders.length > 0 && (
              <>
                <div className="thread-context-menu-divider" />
                <button
                  className="thread-context-menu-item"
                  onClick={() => setShowMoveMenu(true)}
                >
                  <FolderOpen size={14} />
                  <span>{text.sidebar.moveToFolder}</span>
                  <ChevronRight size={14} className="context-menu-arrow" />
                </button>
              </>
            )
          )}
          <div className="thread-context-menu-divider" />
          <button
            className="thread-context-menu-item danger"
            onClick={() => setConfirming(true)}
          >
            <Trash2 size={14} />
            <span>{text.sidebar.deleteThread}</span>
          </button>
        </>
      )}
    </div>
  );
}
