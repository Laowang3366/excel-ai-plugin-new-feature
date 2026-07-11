/**
 * 聊天页面 — 主对话界面（壳组件）
 *
 * 负责组合各子模块，保持自身行数在 400 以内。
 *
 * 已拆分模块：
 * - utils/chatHelpers.tsx: 纯函数 + 小组件（分组、时长、标题、格式化等）
 * - components/common/FeatureSidebarPanel.tsx: 功能模块侧栏
 * - components/chat/AssistantGroupBlock.tsx: 助手消息组渲染
 * - components/chat/ChatMessageList.tsx: 消息列表区域
 * - components/chat/ComposerArea.tsx: 输入框区域
 * - hooks/useComposer.ts: 输入框状态与交互
 * - hooks/useTaskDrafts.ts: 任务面板草稿管理
 */

import React, { useEffect, useState, useCallback, useReducer, useRef } from "react";
import { useChatStore } from "../store/chatStore";
import { useSettingsStore } from "../store/settingsStore";
import type { FolderFileInfo } from "../electronApi";
import { ChatMessageList } from "./chat/ChatMessageList";
import { ComposerArea } from "./chat/ComposerArea";
import { ChatFolderBadge } from "./chat/ChatFolderBadge";
import { WelcomeWorkspace } from "./chat/WelcomeWorkspace";
import { FeatureSidebarPanel } from "./common/FeatureSidebarPanel";
import { ToolConfirmDialog } from "./chat/ToolConfirmDialog";
import { FormulaTaskComposerPanel } from "./task/FormulaTaskComposerPanel";
import { CodeTaskComposerPanel } from "./task/CodeTaskComposerPanel";
import { OCRTaskComposerPanel } from "./task/OCRTaskComposerPanel";
import { ReportTaskComposerPanel } from "./task/ReportTaskComposerPanel";
import { SimpleTaskComposerPanel, type SimpleTaskIntent } from "./task/SimpleTaskComposerPanel";
import { getChatTitleSummary, MessageBubbleIcon } from "../utils/chatHelpers";
import { useComposer } from "../hooks/useComposer";
import { useTaskDrafts } from "../hooks/useTaskDrafts";
import type { IntentKind } from "./Sidebar";
import { getAppText } from "../i18n";
import { ipcApi } from "../services/ipcApi";
import {
  reduceFeatureSidebarState,
  shouldFocusFeatureSidebarOnToggle,
  shouldRestoreFeatureSidebarFocus,
  type FeatureSidebarCloseReason,
} from "../utils/featureSidebarState";
import { Sparkles } from "./common/IconMap";
import type { SettingsSection } from "./SettingsPage";

interface ChatPageProps {
  onOpenSettings: (section?: SettingsSection) => void;
}

function createInitialFeatureSidebarState() {
  return {
    isOpen: typeof window !== "undefined" && window.innerWidth > 1180,
    activeIntent: null,
  };
}

export const ChatPage: React.FC<ChatPageProps> = ({ onOpenSettings }) => {
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
  const [featureSidebar, dispatchFeatureSidebar] = useReducer(
    reduceFeatureSidebarState,
    undefined,
    createInitialFeatureSidebarState,
  );
  const { isOpen: featureSidebarOpen, activeIntent } = featureSidebar;
  const featureSidebarToggleRef = useRef<HTMLButtonElement>(null);

  const closeFeatureSidebar = useCallback(
    (reason: FeatureSidebarCloseReason) => {
      if (!featureSidebarOpen) return;
      dispatchFeatureSidebar({ type: "close" });
      if (shouldRestoreFeatureSidebarFocus(reason)) {
        window.requestAnimationFrame(() => featureSidebarToggleRef.current?.focus());
      }
    },
    [featureSidebarOpen],
  );

  const closeFeatureSidebarManually = useCallback(() => {
    closeFeatureSidebar("manual");
  }, [closeFeatureSidebar]);

  const closeFeatureSidebarAfterSend = useCallback(() => {
    closeFeatureSidebar("send");
  }, [closeFeatureSidebar]);

  const toggleFeatureSidebar = useCallback(() => {
    const focusFirstShortcut = shouldFocusFeatureSidebarOnToggle(featureSidebarOpen);
    dispatchFeatureSidebar({ type: "toggle" });
    if (focusFirstShortcut) {
      window.requestAnimationFrame(() => {
        featureSidebarToggleRef.current
          ?.closest(".chat-page")
          ?.querySelector<HTMLButtonElement>(".feature-sidebar-shortcut")
          ?.focus();
      });
    }
  }, [featureSidebarOpen]);

  const selectFeature = useCallback((intent: NonNullable<IntentKind>) => {
    dispatchFeatureSidebar({ type: "select", intent });
  }, []);

  // 当前会话所属文件夹信息
  const activeThread = threads.find((t) => t.threadId === activeThreadId);
  const currentFolderId = activeThread?.folderId || pendingFolderId;
  const [folderBadgeHidden, setFolderBadgeHidden] = useState(false);
  const currentFolder = currentFolderId
    ? pinnedFolders.find((f) => f.path === currentFolderId)
    : undefined;
  useEffect(() => {
    setFolderBadgeHidden(false);
  }, [currentFolderId]);
  const [currentFolderFiles, setCurrentFolderFiles] = useState<FolderFileInfo[]>([]);
  useEffect(() => {
    let cancelled = false;
    if (currentFolderId) {
      ipcApi.folder
        .listFiles(currentFolderId)
        .then((files) => {
          if (!cancelled) setCurrentFolderFiles(files);
        })
        .catch(() => {
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
  const { setInputText, handleSend, hasInput, showFolderFileList, setShowFolderFileList } =
    composer;

  // TaskDrafts hook
  const {
    taskDrafts,
    setTaskDrafts,
    updateFormulaDraft,
    updateCodeDraft,
    updateOCRDraft,
    updateReportDraft,
    handleSimplePickRange,
    moveTaskDrafts,
  } = useTaskDrafts(composerDraftKey);

  const composerHandleSend = useCallback(() => {
    const sourceDraftKey = !activeThreadId && hasInput ? composerDraftKey : null;
    const send = handleSend();
    closeFeatureSidebarAfterSend();
    if (sourceDraftKey) {
      void send.then((threadId) => {
        if (threadId) moveTaskDrafts(sourceDraftKey, threadId);
      });
    }
  }, [
    activeThreadId,
    hasInput,
    composerDraftKey,
    handleSend,
    closeFeatureSidebarAfterSend,
    moveTaskDrafts,
  ]);

  // 从 TaskComposerPanel 提交
  const handleTaskSubmit = useCallback(
    (payload: string) => {
      const sourceDraftKey = !activeThreadId ? composerDraftKey : null;
      setInputText(payload);
      const send = sendMessage(payload);
      window.setTimeout(() => setInputText(""), 0);
      closeFeatureSidebarAfterSend();
      if (sourceDraftKey) {
        void send.then((threadId) => {
          if (threadId) moveTaskDrafts(sourceDraftKey, threadId);
        });
      }
    },
    [
      activeThreadId,
      composerDraftKey,
      sendMessage,
      setInputText,
      closeFeatureSidebarAfterSend,
      moveTaskDrafts,
    ],
  );

  const updateSimpleRange = useCallback(
    (intent: SimpleTaskIntent, range: string) => {
      setTaskDrafts((prev) => ({
        ...prev,
        [intent]: {
          range,
          task: prev[intent]?.task ?? "",
        },
      }));
    },
    [setTaskDrafts],
  );

  const updateSimpleTask = useCallback(
    (intent: SimpleTaskIntent, task: string) => {
      setTaskDrafts((prev) => ({
        ...prev,
        [intent]: {
          range: prev[intent]?.range ?? "",
          task,
        },
      }));
    },
    [setTaskDrafts],
  );

  const isEmpty = messages.length === 0 && !isStreaming;
  const showWelcomeComposer = isEmpty;
  const chatTitle = getChatTitleSummary(messages, text.chat.newChat);

  return (
    <div className={`chat-page ${showWelcomeComposer ? "welcome-chat" : ""}`}>
      <div className="chat-workspace">
        {/* 顶部栏 */}
        <div className="chat-header">
          <h2 title={chatTitle}>
            <MessageBubbleIcon /> <span>{chatTitle}</span>
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
            <button
              ref={featureSidebarToggleRef}
              className={`feature-sidebar-toggle ${featureSidebarOpen ? "active" : ""}`}
              type="button"
              onClick={toggleFeatureSidebar}
              title={
                featureSidebarOpen ? text.chat.featureSidebar.close : text.chat.featureSidebar.open
              }
              aria-label={
                featureSidebarOpen ? text.chat.featureSidebar.close : text.chat.featureSidebar.open
              }
              aria-pressed={featureSidebarOpen}
            >
              <Sparkles size={15} />
            </button>
          </div>
        </div>

        {/* 消息列表 */}
        <ChatMessageList onSend={composerHandleSend} onFillInput={(t) => setInputText(t)} />

        {showWelcomeComposer && (
          <WelcomeWorkspace language={language} onIntentClick={selectFeature} />
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

      <FeatureSidebarPanel
        isOpen={featureSidebarOpen}
        activeIntent={activeIntent}
        language={language}
        onIntentClick={selectFeature}
        onClose={closeFeatureSidebarManually}
      >
        {activeIntent === "formula" && (
          <FormulaTaskComposerPanel
            key={`${composerDraftKey}:formula`}
            embedded
            draft={taskDrafts.formula}
            onDraftChange={updateFormulaDraft}
            onSubmit={handleTaskSubmit}
            onClose={closeFeatureSidebarManually}
          />
        )}
        {activeIntent === "code" && (
          <CodeTaskComposerPanel
            key={`${composerDraftKey}:code`}
            embedded
            draft={taskDrafts.code}
            onDraftChange={updateCodeDraft}
            onSubmit={handleTaskSubmit}
            onClose={closeFeatureSidebarManually}
          />
        )}
        {activeIntent === "ocr" && (
          <OCRTaskComposerPanel
            key={`${composerDraftKey}:ocr`}
            embedded
            draft={taskDrafts.ocr}
            onDraftChange={updateOCRDraft}
            onSubmit={handleTaskSubmit}
            onClose={closeFeatureSidebarManually}
          />
        )}
        {activeIntent === "report" && (
          <ReportTaskComposerPanel
            key={`${composerDraftKey}:report`}
            embedded
            draft={taskDrafts.report}
            onDraftChange={updateReportDraft}
            onSubmit={handleTaskSubmit}
            onClose={closeFeatureSidebarManually}
          />
        )}
        {(activeIntent === "clean" || activeIntent === "chart") && (
          <SimpleTaskComposerPanel
            key={`${composerDraftKey}:${activeIntent}`}
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
      </FeatureSidebarPanel>
    </div>
  );
};
