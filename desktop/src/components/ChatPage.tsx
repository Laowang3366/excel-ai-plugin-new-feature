/**
 * 聊天页面 — 主对话界面（壳组件）
 *
 * 负责组合消息区、输入区和任务面板，不持有侧栏焦点或文件夹加载生命周期。
 *
 * 已拆分模块：
 * - utils/chatHelpers.tsx: 纯函数 + 小组件（分组、时长、标题、格式化等）
 * - components/common/FeatureSidebarPanel.tsx: 功能模块侧栏
 * - components/chat/AssistantGroupBlock.tsx: 助手消息组渲染
 * - components/chat/ChatMessageList.tsx: 消息列表区域
 * - components/chat/ComposerArea.tsx: 输入框区域
 * - hooks/useComposer.ts: 输入框状态与交互
 * - hooks/useTaskDrafts.ts: 任务面板草稿管理
 * - hooks/useFeatureSidebarController.ts: 功能侧栏状态与焦点生命周期
 * - hooks/useCurrentChatFolder.ts: 当前会话文件夹与文件加载
 */

import { useCallback } from "react";
import { useChatStore } from "../store/chatStore";
import { useSettingsStore } from "../store/settingsStore";
import type { WindowDisplayMode } from "../electronApi";
import { ChatMessageList } from "./chat/ChatMessageList";
import { ComposerArea } from "./chat/ComposerArea";
import { ChatFolderBadge } from "./chat/ChatFolderBadge";
import { ChatFeatureSidebar } from "./chat/ChatFeatureSidebar";
import { OfficeLauncher } from "./chat/OfficeLauncher";
import { ToolConfirmDialog } from "./chat/ToolConfirmDialog";
import { getChatTitleSummary, MessageBubbleIcon } from "../utils/chatHelpers";
import { useComposer } from "../hooks/useComposer";
import { useCurrentChatFolder } from "../hooks/useCurrentChatFolder";
import { useFeatureSidebarController } from "../hooks/useFeatureSidebarController";
import { useTaskDrafts } from "../hooks/useTaskDrafts";
import { getAppText } from "../i18n";
import { Maximize2, Minimize2, PanelRight } from "./common/IconMap";
import type { SettingsSection } from "./SettingsPage";

interface ChatPageProps {
  displayMode: WindowDisplayMode;
  onToggleCompactMode: () => void;
  onOpenSettings: (section?: SettingsSection) => void;
}

export const ChatPage: React.FC<ChatPageProps> = ({
  displayMode,
  onToggleCompactMode,
  onOpenSettings,
}) => {
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
  const {
    activeIntent,
    closeAfterSend: closeFeatureSidebarAfterSend,
    closeManually: closeFeatureSidebarManually,
    isOpen: featureSidebarOpen,
    select: selectFeature,
    toggle: toggleFeatureSidebar,
    toggleRef: featureSidebarToggleRef,
  } = useFeatureSidebarController();

  // 当前会话所属文件夹信息
  const activeThread = threads.find((t) => t.threadId === activeThreadId);
  const currentFolderId = activeThread?.folderId || pendingFolderId;
  const currentFolder = useCurrentChatFolder(currentFolderId, pinnedFolders);

  // Composer hook
  const composerDraftKey = activeThreadId ?? (pendingFolderId ? `new:${pendingFolderId}` : "new");
  const composer = useComposer(composerDraftKey);
  const { setInputText, handleSend, hasInput, showFolderFileList, setShowFolderFileList } =
    composer;

  // TaskDrafts hook
  const taskDraftsController = useTaskDrafts(composerDraftKey);
  const { moveTaskDrafts } = taskDraftsController;

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

  const isEmpty = messages.length === 0 && !isStreaming;
  const showWelcomeComposer = isEmpty && !activeIntent;
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
            {currentFolder.folder && !currentFolder.badgeHidden && (
              <ChatFolderBadge
                folder={currentFolder.folder}
                files={currentFolder.files}
                open={showFolderFileList}
                text={text.chat}
                onToggleOpen={() => setShowFolderFileList((visible) => !visible)}
                onClose={() => setShowFolderFileList(false)}
                onAddFilesToComposer={addFilesToComposer}
                onHideBadge={currentFolder.hideBadge}
              />
            )}
            <OfficeLauncher text={text.app.officeLauncher} />
            <button
              className={`feature-sidebar-toggle chat-window-mode-toggle ${
                displayMode === "compact" ? "active" : ""
              }`}
              type="button"
              onClick={onToggleCompactMode}
              title={displayMode === "normal" ? text.app.compactWindow : text.app.restoreWindow}
              aria-label={
                displayMode === "normal" ? text.app.compactWindow : text.app.restoreWindow
              }
              aria-pressed={displayMode === "compact"}
            >
              {displayMode === "normal" ? <Minimize2 size={15} /> : <Maximize2 size={15} />}
            </button>
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
              <PanelRight size={17} />
            </button>
          </div>
        </div>

        {/* 消息列表 */}
        <ChatMessageList onFillInput={(t) => setInputText(t)} />

        {/* Pill Composer 输入框 */}
        <ComposerArea
          composer={composer}
          currentFolder={currentFolder.folder}
          currentFolderFiles={currentFolder.files}
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

      <ChatFeatureSidebar
        isOpen={featureSidebarOpen}
        activeIntent={activeIntent}
        composerDraftKey={composerDraftKey}
        controller={taskDraftsController}
        language={language}
        onIntentClick={selectFeature}
        onClose={closeFeatureSidebarManually}
        onTaskSubmit={handleTaskSubmit}
      />
    </div>
  );
};
