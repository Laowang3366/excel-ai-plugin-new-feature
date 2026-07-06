/**
 * 聊天页面 — 主对话界面（壳组件）
 *
 * 负责组合各子模块，保持自身行数在 400 以内。
 *
 * 已拆分模块：
 * - utils/chatHelpers.tsx: 纯函数 + 小组件（分组、时长、标题、格式化等）
 * - components/common/FloatingTaskPanel.tsx: 可拖拽浮窗容器
 * - components/chat/AssistantGroupBlock.tsx: 助手消息组渲染
 * - components/chat/ChatMessageList.tsx: 消息列表区域
 * - components/chat/ComposerArea.tsx: 输入框区域
 * - hooks/useComposer.ts: 输入框状态与交互
 * - hooks/useTaskDrafts.ts: 任务面板草稿管理
 */

import React, { useEffect, useState, useCallback, useMemo } from "react";
import { useChatStore } from "../store/chatStore";
import { useSettingsStore } from "../store/settingsStore";
import type { FolderFileInfo } from "../electronApi";
import { ChatMessageList } from "./chat/ChatMessageList";
import { ComposerArea } from "./chat/ComposerArea";
import { ChatFolderBadge } from "./chat/ChatFolderBadge";
import { FloatingTaskPanel } from "./common/FloatingTaskPanel";
import { FeatureFloatingDock } from "./common/FeatureFloatingDock";
import { ToolConfirmDialog } from "./chat/ToolConfirmDialog";
import { FormulaTaskComposerPanel } from "./task/FormulaTaskComposerPanel";
import { CodeTaskComposerPanel } from "./task/CodeTaskComposerPanel";
import { OCRTaskComposerPanel } from "./task/OCRTaskComposerPanel";
import { ReportTaskComposerPanel } from "./task/ReportTaskComposerPanel";
import { SimpleTaskComposerPanel, type SimpleTaskIntent } from "./task/SimpleTaskComposerPanel";
import { OfficePreviewPanel } from "./office/OfficePreviewPanel";
import {
  getChatTitleSummary,
  MessageBubbleIcon,
  type ActiveIntentKind,
} from "../utils/chatHelpers";
import { useComposer } from "../hooks/useComposer";
import { useTaskDrafts } from "../hooks/useTaskDrafts";
import type { IntentKind } from "./Sidebar";
import { getAppText } from "../i18n";
import { ipcApi } from "../services/ipcApi";
import {
  collectOfficeEditEvents,
  getOfficePreviewToggleLocation,
} from "../utils/officeEditEvents";
import {
  Activity,
} from "./common/IconMap";
import type { SettingsSection } from "./SettingsPage";

interface ChatPageProps {
  onOpenSettings: (section?: SettingsSection) => void;
  activeIntent: IntentKind;
  onIntentClick: (intent: IntentKind) => void;
}

export const ChatPage: React.FC<ChatPageProps> = ({ onOpenSettings, activeIntent, onIntentClick }) => {
  const {
    messages,
    isStreaming,
    sendMessage,
    interruptTurn,
    pendingToolCall,
    confirmToolCall,
    cancelToolCall,
    activeThreadId,
    threads,
    pendingFolderId,
    addFilesToComposer,
  } = useChatStore();

  const { language, pinnedFolders } = useSettingsStore();
  const text = getAppText(language);

  // 当前会话所属文件夹信息
  const activeThread = threads.find((t) => t.threadId === activeThreadId);
  const currentFolderId = activeThread?.folderId || pendingFolderId;
  const [folderBadgeHidden, setFolderBadgeHidden] = useState(false);
  const currentFolder = currentFolderId ? pinnedFolders.find((f) => f.path === currentFolderId) : undefined;
  React.useEffect(() => { setFolderBadgeHidden(false); }, [currentFolderId]);
  const [currentFolderFiles, setCurrentFolderFiles] = useState<FolderFileInfo[]>([]);
  const [officePreviewOpen, setOfficePreviewOpen] = useState(false);
  useEffect(() => {
    let cancelled = false;
    if (currentFolderId) {
      ipcApi.folder.listFiles(currentFolderId).then((files) => {
        if (!cancelled) setCurrentFolderFiles(files);
      }).catch(() => {
        if (!cancelled) setCurrentFolderFiles([]);
      });
    } else {
      setCurrentFolderFiles([]);
    }
    return () => {
      cancelled = true;
    };
  }, [currentFolderId]);

  // Composer hook
  const composerDraftKey = activeThreadId ?? (pendingFolderId ? `new:${pendingFolderId}` : "new");
  const composer = useComposer(composerDraftKey);
  const { inputText, setInputText, handleSend, showFolderFileList, setShowFolderFileList } = composer;

  // 意图关闭时清理 composer
  const composerHandleSend = useCallback(() => {
    handleSend();
    onIntentClick(null);
  }, [handleSend, onIntentClick]);

  // TaskDrafts hook
  const {
    taskDrafts, setTaskDrafts,
    closeActiveTaskPanel,
    updateFormulaDraft, updateCodeDraft, updateOCRDraft, updateReportDraft,
    handleSimplePickRange,
  } = useTaskDrafts(activeIntent, onIntentClick, composerDraftKey);

  // 从 TaskComposerPanel 提交
  const handleTaskSubmit = useCallback((payload: string) => {
    setInputText(payload);
    sendMessage(payload);
    window.setTimeout(() => setInputText(""), 0);
    onIntentClick(null);
  }, [sendMessage, setInputText, onIntentClick]);

  const updateSimpleRange = useCallback((intent: SimpleTaskIntent, range: string) => {
    setTaskDrafts((prev) => ({
      ...prev,
      [intent]: {
        range,
        task: prev[intent]?.task ?? "",
      },
    }));
  }, [setTaskDrafts]);

  const updateSimpleTask = useCallback((intent: SimpleTaskIntent, task: string) => {
    setTaskDrafts((prev) => ({
      ...prev,
      [intent]: {
        range: prev[intent]?.range ?? "",
        task,
      },
    }));
  }, [setTaskDrafts]);

  const isEmpty = messages.length === 0 && !isStreaming;
  const showWelcomeComposer = isEmpty && !activeIntent;
  const chatTitle = getChatTitleSummary(messages, text.chat.newChat);
  const officeEditEvents = useMemo(() => collectOfficeEditEvents(messages), [messages]);
  const showOfficePreviewPanel = officePreviewOpen;
  const officePreviewToggleLocation = getOfficePreviewToggleLocation(showOfficePreviewPanel);

  return (
    <div className={`chat-page ${showWelcomeComposer ? "welcome-chat" : ""}`}>
      <div className="chat-workspace">
        {/* 顶部栏 */}
        <div className="chat-header">
          <h2 title={chatTitle}>
            <MessageBubbleIcon />{" "}
            <span>{chatTitle}</span>
          </h2>
          <div className="chat-header-actions">
            {currentFolder && !folderBadgeHidden && (
              <ChatFolderBadge
                folder={currentFolder}
                files={currentFolderFiles}
                open={showFolderFileList}
                text={text.chat}
                onToggleOpen={() => setShowFolderFileList((visible) => !visible)}
                onClose={() => setShowFolderFileList(false)}
                onAddFilesToComposer={addFilesToComposer}
                onHideBadge={() => setFolderBadgeHidden(true)}
              />
            )}
            {officePreviewToggleLocation === "chat-header" && (
              <button
                className="office-preview-toggle"
                type="button"
                onClick={() => setOfficePreviewOpen(true)}
                title="显示文件编辑侧栏"
                aria-label="显示文件编辑侧栏"
                aria-pressed={false}
              >
                <Activity size={14} />
              </button>
            )}
          </div>
        </div>

      {/* 消息列表 */}
      <ChatMessageList onSend={composerHandleSend} onFillInput={(t) => setInputText(t)} />

      <FeatureFloatingDock
        activeIntent={activeIntent}
        onIntentClick={onIntentClick}
      />

      {/* 任务编排浮窗 */}
      {activeIntent && (
        <FloatingTaskPanel intent={activeIntent} onClose={closeActiveTaskPanel}>
          {activeIntent === "formula" && (
            <FormulaTaskComposerPanel
              embedded
              draft={taskDrafts.formula}
              onDraftChange={updateFormulaDraft}
              onSubmit={handleTaskSubmit}
              onClose={closeActiveTaskPanel}
            />
          )}
          {activeIntent === "code" && (
            <CodeTaskComposerPanel
              embedded
              draft={taskDrafts.code}
              onDraftChange={updateCodeDraft}
              onSubmit={handleTaskSubmit}
              onClose={closeActiveTaskPanel}
            />
          )}
          {activeIntent === "ocr" && (
            <OCRTaskComposerPanel
              embedded
              draft={taskDrafts.ocr}
              onDraftChange={updateOCRDraft}
              onSubmit={handleTaskSubmit}
              onClose={closeActiveTaskPanel}
            />
          )}
          {activeIntent === "report" && (
            <ReportTaskComposerPanel
              embedded
              draft={taskDrafts.report}
              onDraftChange={updateReportDraft}
              onSubmit={handleTaskSubmit}
              onClose={closeActiveTaskPanel}
            />
          )}
          {(activeIntent === "clean" || activeIntent === "chart") && (
            <SimpleTaskComposerPanel
              intent={activeIntent}
              range={taskDrafts[activeIntent]?.range ?? ""}
              task={taskDrafts[activeIntent]?.task ?? ""}
              text={text.chat}
              onRangeChange={(range) => updateSimpleRange(activeIntent, range)}
              onTaskChange={(task) => updateSimpleTask(activeIntent, task)}
              onPickRange={handleSimplePickRange}
              onSubmit={handleTaskSubmit}
            />
          )}
        </FloatingTaskPanel>
      )}

      {/* Pill Composer 输入框 */}
      <ComposerArea
        composer={composer}
        currentFolder={currentFolder}
        currentFolderFiles={currentFolderFiles}
        showWelcomeComposer={showWelcomeComposer}
        onSend={composerHandleSend}
        onInterrupt={interruptTurn}
        onOpenSettings={onOpenSettings}
      />

      {/* 工具确认弹窗 */}
      {pendingToolCall && (
        <ToolConfirmDialog
          pendingCall={pendingToolCall}
          onConfirm={(alwaysAllow) => confirmToolCall(pendingToolCall.id, alwaysAllow)}
          onCancel={() => cancelToolCall(pendingToolCall.id)}
        />
      )}
      </div>

      <OfficePreviewPanel
        events={officeEditEvents}
        isOpen={showOfficePreviewPanel}
        onToggle={() => setOfficePreviewOpen(false)}
      />
    </div>
  );
};
