import React from "react";
import type { AppLanguage } from "../../store/settingsStore";
import { formatTime, getThreadDisplayStatus, getThreadStatusLabel } from "../../utils/sidebarHelpers";
import { MessageSquare, RefreshCw } from "../common/IconMap";

export interface SidebarThreadItemData {
  threadId: string;
  preview?: string;
  updatedAt: number;
  lastTurnStatus?: string;
}

export interface SidebarThreadStatusInput {
  thread: SidebarThreadItemData;
  activeThreadId: string | null;
  runningThreadIds: Record<string, boolean>;
  turnStatus: string;
  viewedThreadStatusAt: Record<string, number>;
}

export function getSidebarThreadItemStatus({
  thread,
  activeThreadId,
  runningThreadIds,
  turnStatus,
  viewedThreadStatusAt,
}: SidebarThreadStatusInput) {
  const isActiveThread = activeThreadId === thread.threadId;
  const isRunningThread = Boolean(runningThreadIds[thread.threadId]);
  const statusViewed = isActiveThread || viewedThreadStatusAt[thread.threadId] === thread.updatedAt;
  const status = getThreadDisplayStatus(
    thread.lastTurnStatus,
    isRunningThread ? "in_progress" : isActiveThread ? turnStatus as any : undefined,
    statusViewed,
  );
  return { isActiveThread, status };
}

interface SidebarThreadItemProps extends SidebarThreadStatusInput {
  language: AppLanguage;
  fallbackTitle: string;
  inFolder?: boolean;
  onSwitchThread: (threadId: string) => void;
  onThreadContextMenu: (event: React.MouseEvent, threadId: string, inFolder?: boolean) => void;
}

export const SidebarThreadItem: React.FC<SidebarThreadItemProps> = ({
  thread,
  activeThreadId,
  runningThreadIds,
  turnStatus,
  viewedThreadStatusAt,
  language,
  fallbackTitle,
  inFolder = false,
  onSwitchThread,
  onThreadContextMenu,
}) => {
  const { isActiveThread, status } = getSidebarThreadItemStatus({
    thread,
    activeThreadId,
    runningThreadIds,
    turnStatus,
    viewedThreadStatusAt,
  });
  const statusLabel = getThreadStatusLabel(status, language);

  return (
    <div
      className={`sidebar-thread-item ${inFolder ? "sidebar-thread-in-folder " : ""}${isActiveThread ? "active" : ""}`}
      onClick={() => onSwitchThread(thread.threadId)}
      onContextMenu={(event) => onThreadContextMenu(event, thread.threadId, inFolder || undefined)}
    >
      <div className="thread-item-main">
        {inFolder && <MessageSquare size={12} className="thread-item-icon" />}
        <div className="thread-item-preview">
          {thread.preview || fallbackTitle}
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
};
