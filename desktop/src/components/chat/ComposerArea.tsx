/**
 * Composer 输入区 — Pill Composer 底部输入框
 *
 * 从 ChatPage.tsx 提取的输入区域，包含：
 * - 欢迎标题
 * - 文件夹上下文提示
 * - 附件列表
 * - 文本输入
 * - 工具栏（附件/权限/思考/模型切换/发送/停止）
 *
 * Props 设计：接收 useComposer() hook 返回值作为单一 `composer` prop，
 * 加上少数 ChatPage 拥有的外部状态，避免 28 props 的 prop drilling。
 */

import React from "react";
import { COMPOSER_INPUT_MAX_LENGTH } from "../../hooks/useComposer";
import { useChatStore } from "../../store/chatStore";
import { useSettingsStore } from "../../store/settingsStore";
import type { FolderFileInfo } from "../../electronApi";
import { getAppText } from "../../i18n";
import { formatEstimatedUsedTokens, formatTokensAsK } from "../../utils/modelContextWindows";
import { PermissionIcon } from "../../utils/chatHelpers";
import { ModelQuickSwitch } from "../chat/ModelQuickSwitch";
import { AttachmentImagePreview } from "../chat/AttachmentImagePreview";
import { ComposerThinkingModeButton } from "./ComposerThinkingModeButton";
import type { SettingsSection } from "../SettingsPage";
import {
  Square,
  ArrowUp,
  Plus,
  Paperclip,
  Image,
  ShieldAlert,
  ShieldCheck,
  ShieldX,
  Activity,
  FolderOpen,
  X,
  ChevronDown,
} from "../common/IconMap";
import { getComposerPrimaryAction, isComposerSubmitKey } from "./composerPrimaryAction";

export {
  getComposerPrimaryAction,
  isComposerSubmitKey,
  type ComposerPrimaryAction,
} from "./composerPrimaryAction";

type ComposerState = ReturnType<typeof import("../../hooks/useComposer").useComposer>;

interface ComposerAreaProps {
  /** useComposer hook 返回的全部状态与回调 */
  composer: ComposerState;
  /** 当前关联的文件夹 */
  currentFolder?: { path: string; name: string };
  /** 文件夹内的 Office 文件列表 */
  currentFolderFiles: FolderFileInfo[];
  /** 是否显示欢迎标题 */
  showWelcomeComposer: boolean;
  /** 发送消息回调（由 ChatPage 包装，含意图清理） */
  onSend: () => void;
  /** 中断生成回调 */
  onInterrupt: () => void;
  /** 打开设置回调 */
  onOpenSettings: (section?: SettingsSection) => void;
}

export function ComposerArea({
  composer,
  showWelcomeComposer,
  onSend,
  onInterrupt,
  onOpenSettings,
}: ComposerAreaProps) {
  const {
    inputText,
    textareaRef,
    hasInput,
    isStreaming,
    composerDragOver,
    attachedFiles,
    showAttachPopover,
    setShowAttachPopover,
    showPermissionPopover,
    setShowPermissionPopover,
    showThinkingPopover,
    setShowThinkingPopover,
    handleOpenFile,
    handleOpenImage,
    handleOpenFolder,
    removeAttachedFile,
    handleComposerDragOver,
    handleComposerDragLeave,
    handleComposerDrop,
    handleTextareaChange,
    handlePaste,
  } = composer;

  const { permissionMode, setPermissionMode, language } = useSettingsStore();
  const { contextUsage } = useChatStore();
  const text = getAppText(language);
  const primaryAction = getComposerPrimaryAction(isStreaming, hasInput);
  const handleComposerKeyDown = (event: React.KeyboardEvent<HTMLTextAreaElement>) => {
    if (!isComposerSubmitKey(event.key, event.shiftKey, event.nativeEvent.isComposing)) return;
    event.preventDefault();
    onSend();
  };

  return (
    <div className="chat-input-area">
      <div
        className={`composer${composerDragOver ? " drag-over" : ""}`}
        onDragOver={handleComposerDragOver}
        onDragLeave={handleComposerDragLeave}
        onDrop={handleComposerDrop}
      >
        {/* 附件 chip 列表 */}
        {attachedFiles.length > 0 && (
          <div className="composer-attachments">
            {attachedFiles.map((f, i) =>
              f.fileType === "image" ? (
                <AttachmentImagePreview
                  key={`${f.filePath}-${i}`}
                  attachment={f}
                  variant="composer"
                  onRemove={() => removeAttachedFile(i)}
                />
              ) : (
                <div key={`${f.filePath}-${i}`} className={`attach-chip ${f.fileType}`}>
                  <Paperclip size={12} />
                  <span className="attach-chip-name">{f.fileName}</span>
                  <button className="attach-chip-remove" onClick={() => removeAttachedFile(i)}>
                    <X size={10} />
                  </button>
                </div>
              ),
            )}
          </div>
        )}
        <textarea
          ref={textareaRef}
          className="composer-textarea"
          value={inputText}
          onChange={handleTextareaChange}
          onKeyDown={handleComposerKeyDown}
          placeholder={
            isStreaming
              ? text.chat.aiReplying
              : showWelcomeComposer
                ? text.chat.welcomePlaceholder
                : text.chat.inputPlaceholder
          }
          maxLength={COMPOSER_INPUT_MAX_LENGTH}
          rows={2}
          onPaste={handlePaste}
        />
        <div className="composer-toolbar">
          <div className="composer-left">
            {/* + 附件/图片按钮 */}
            <div className="composer-popover-wrapper">
              <button
                className={`composer-action-btn ${showAttachPopover ? "active" : ""}`}
                onClick={(e) => {
                  e.stopPropagation();
                  setShowAttachPopover(!showAttachPopover);
                  setShowPermissionPopover(false);
                  setShowThinkingPopover(false);
                }}
                title={text.chat.addAttachmentTitle}
              >
                <Plus size={19} />
              </button>
              {showAttachPopover && (
                <div className="composer-popover" onClick={(e) => e.stopPropagation()}>
                  <button className="popover-item" onClick={handleOpenFile}>
                    <Paperclip size={14} /> {text.chat.addAttachment}
                  </button>
                  <button className="popover-item" onClick={handleOpenImage}>
                    <Image size={14} /> {text.chat.uploadImage}
                  </button>
                  <button className="popover-item" onClick={handleOpenFolder}>
                    <FolderOpen size={14} /> {text.chat.addFolder}
                  </button>
                </div>
              )}
            </div>

            {/* 权限模式切换 */}
            <div className="composer-popover-wrapper">
              <button
                className={`composer-action-btn permission-btn permission-${permissionMode} ${showPermissionPopover ? "active" : ""}`}
                onClick={(e) => {
                  e.stopPropagation();
                  setShowPermissionPopover(!showPermissionPopover);
                  setShowAttachPopover(false);
                  setShowThinkingPopover(false);
                }}
                title={text.chat.permissionMode}
              >
                <PermissionIcon mode={permissionMode} />
                <span className="permission-label">
                  {text.chat.permissionLabels[permissionMode]}
                </span>
                <ChevronDown size={13} />
              </button>
              {showPermissionPopover && (
                <div className="composer-popover" onClick={(e) => e.stopPropagation()}>
                  <button
                    className={`popover-item permission-option permission-normal ${permissionMode === "normal" ? "active" : ""}`}
                    onClick={() => {
                      setPermissionMode("normal");
                      setShowPermissionPopover(false);
                    }}
                  >
                    <ShieldAlert size={15} /> {text.chat.permissionLabels.normal}
                  </button>
                  <button
                    className={`popover-item permission-option permission-auto_approve_safe ${permissionMode === "auto_approve_safe" ? "active" : ""}`}
                    onClick={() => {
                      setPermissionMode("auto_approve_safe");
                      setShowPermissionPopover(false);
                    }}
                  >
                    <ShieldCheck size={15} /> {text.chat.permissionLabels.auto_approve_safe}
                  </button>
                  <button
                    className={`popover-item permission-option permission-confirm_all ${permissionMode === "confirm_all" ? "active" : ""}`}
                    onClick={() => {
                      setPermissionMode("confirm_all");
                      setShowPermissionPopover(false);
                    }}
                  >
                    <ShieldX size={15} /> {text.chat.permissionLabels.confirm_all}
                  </button>
                </div>
              )}
            </div>
          </div>
          <div className="composer-right">
            {/* 上下文使用量指示器 */}
            {contextUsage && (
              <div
                className={`context-indicator ${
                  contextUsage.percentage >= 90
                    ? "danger"
                    : contextUsage.percentage >= 70
                      ? "warning"
                      : "normal"
                }`}
                title={`${text.chat.contextUsage}: ${contextUsage.estimatedTokens.toLocaleString()} / ${contextUsage.contextWindowSize.toLocaleString()} tokens (${contextUsage.percentage}%)`}
              >
                <Activity size={14} />
                <span className="context-indicator-pct">
                  {formatEstimatedUsedTokens(contextUsage.estimatedTokens)}/
                  {formatTokensAsK(contextUsage.contextWindowSize)}
                </span>
              </div>
            )}
            <ModelQuickSwitch onOpenSettings={onOpenSettings} />
            {/* 思考模式开关 */}
            <ComposerThinkingModeButton
              open={showThinkingPopover}
              setOpen={setShowThinkingPopover}
              closePeerPopovers={() => {
                setShowAttachPopover(false);
                setShowPermissionPopover(false);
              }}
            />
            {primaryAction === "send" ? (
              <button
                className={`btn-send-circle ${hasInput ? "active" : "inactive"}`}
                onClick={onSend}
                disabled={!hasInput}
                title={text.chat.send}
              >
                <ArrowUp size={18} strokeWidth={2.4} />
              </button>
            ) : (
              <button
                className="btn-stop-circle"
                onClick={onInterrupt}
                title={text.chat.stopGenerating}
              >
                <Square size={14} />
              </button>
            )}
          </div>
        </div>
      </div>
    </div>
  );
}
