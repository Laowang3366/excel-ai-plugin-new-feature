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
import { FloatingTaskPanel } from "./common/FloatingTaskPanel";
import { FeatureFloatingDock } from "./common/FeatureFloatingDock";
import { ToolConfirmDialog } from "./chat/ToolConfirmDialog";
import { FormulaTaskComposerPanel } from "./task/FormulaTaskComposerPanel";
import { CodeTaskComposerPanel } from "./task/CodeTaskComposerPanel";
import { OCRTaskComposerPanel } from "./task/OCRTaskComposerPanel";
import { ReportTaskComposerPanel } from "./task/ReportTaskComposerPanel";
import { OfficePreviewPanel } from "./office/OfficePreviewPanel";
import {
  getChatTitleSummary,
  formatFileSize,
  MessageBubbleIcon,
  type ActiveIntentKind,
} from "../utils/chatHelpers";
import { useComposer } from "../hooks/useComposer";
import { useTaskDrafts, type TaskDrafts } from "../hooks/useTaskDrafts";
import type { IntentKind } from "./Sidebar";
import { getAppText } from "../i18n";
import { ipcApi } from "../services/ipcApi";
import {
  collectOfficeEditEvents,
  getOfficePreviewToggleLocation,
} from "../utils/officeEditEvents";
import {
  Activity,
  FolderOpen,
  FileSpreadsheet,
  Ruler,
} from "./common/IconMap";
import type { SettingsSection } from "./SettingsPage";

type SimpleTaskIntent = Extract<ActiveIntentKind, "clean" | "chart">;

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
    if (currentFolderId) {
      ipcApi.folder.listFiles(currentFolderId).then((files) => {
        setCurrentFolderFiles(files);
      }).catch(() => setCurrentFolderFiles([]));
    } else {
      setCurrentFolderFiles([]);
    }
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
              <div className="chat-folder-badge-wrapper">
                <button
                  className="chat-folder-badge"
                  title={currentFolder.path}
                  onClick={() => setShowFolderFileList((v) => !v)}
                >
                  <FolderOpen size={13} />
                  <span className="chat-folder-name">{currentFolder.name}</span>
                  {currentFolderFiles.length > 0 && (
                    <span className="chat-folder-file-count">
                      {text.chat.folderFileCount(currentFolderFiles.length)}
                    </span>
                  )}
                </button>
                {showFolderFileList && currentFolderFiles.length > 0 && (
                  <div className="folder-file-popover" onClick={(e) => e.stopPropagation()}>
                    <div className="folder-file-popover-title">{currentFolder.name}</div>
                    {currentFolderFiles.map((file) => (
                      <button
                        key={file.filePath}
                        className="folder-file-popover-item clickable"
                        title={file.filePath}
                        onClick={() => {
                          addFilesToComposer([{
                            filePath: file.filePath,
                            fileName: file.fileName,
                            fileType: "document" as const,
                            size: file.size,
                          }]);
                          setShowFolderFileList(false);
                          setFolderBadgeHidden(true);
                        }}
                      >
                        <FileSpreadsheet size={13} />
                        <span className="folder-file-popover-name">{file.fileName}</span>
                        <span className="folder-file-popover-size">{formatFileSize(file.size)}</span>
                      </button>
                    ))}
                  </div>
                )}
              </div>
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
            <div className="task-composer-panel">
              <div className="task-field">
                <label className="task-field-label">{text.chat.dataSourceRange}</label>
                <div className="range-input-row">
                  <input
                    className="task-field-input"
                    placeholder={text.chat.rangePlaceholder}
                    value={taskDrafts[activeIntent]?.range ?? ""}
                    onChange={(e) => {
                      const intent = activeIntent;
                      setTaskDrafts((prev) => ({
                        ...prev,
                        [intent]: {
                          range: e.target.value,
                          task: prev[intent]?.task ?? "",
                        },
                      }));
                    }}
                  />
                  <button className="btn-pick-range" onClick={() => handleSimplePickRange(activeIntent as SimpleTaskIntent)}><Ruler size={13} /> {text.chat.pickRange}</button>
                </div>
              </div>
              <div className="task-field">
                <label className="task-field-label">{text.chat.requirement}</label>
                <textarea
                  className="task-field-textarea"
                  value={taskDrafts[activeIntent]?.task ?? ""}
                  onChange={(e) => {
                    const intent = activeIntent;
                    setTaskDrafts((prev) => ({
                      ...prev,
                      [intent]: {
                        range: prev[intent]?.range ?? "",
                        task: e.target.value,
                      },
                    }));
                  }}
                  placeholder={
                    text.chat.simplePlaceholders[activeIntent as SimpleTaskIntent]
                  }
                />
              </div>
              <button
                className="task-submit-btn"
                onClick={() => {
                  const draft = taskDrafts[activeIntent];
                  const prefix = text.chat.simplePrefixes[activeIntent as SimpleTaskIntent];
                  const lines = [prefix];
                  if (draft?.range.trim()) lines.push(`${text.chat.dataSourceRange}: ${draft.range.trim()}`);
                  if (draft?.task.trim()) lines.push(`${text.chat.requirement}: ${draft.task.trim()}`);
                  const payload = lines.join("\n");
                  setInputText(payload);
                  sendMessage(payload);
                  window.setTimeout(() => setInputText(""), 0);
                  onIntentClick(null);
                }}
              >
                {text.chat.sendToAi}
              </button>
            </div>
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
