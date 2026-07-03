/**
 * 消息气泡 — 用户消息和 AI 回复
 *
 * 无头像设计：用户消息右对齐蓝色背景，AI 消息左对齐浅色背景。
 * 推理/思考过程始终可查看（通过 ReasoningBubble 折叠展开）。
 */

import React from "react";
import type { TurnItem, FileAttachment } from "../../electronApi";
import { cleanReasoningText } from "../../utils/textCleaner";
import { getUserFacingMessageContent } from "../../utils/chatHelpers";
import { Paperclip } from "../common/IconMap";
import { AttachmentImagePreview } from "./AttachmentImagePreview";
import { CopyFeedbackButton, MarkdownContent, normalizeVisibleMarkdown } from "./MarkdownContent";
import { useSettingsStore } from "../../store/settingsStore";
import { getAppText } from "../../i18n";

interface MessageBubbleProps {
  item: TurnItem;
}

/**
 * 消息气泡 — 使用 React.memo 避免父组件重渲染时重复执行 cleanReasoningText 和 markdown 解析。
 * 比较依据是 item.id（只要消息 ID 不变就不重渲染）。
 */
export const MessageBubble: React.FC<MessageBubbleProps> = React.memo(({ item }) => {
  if (item.type === "user_message") {
    const content = getUserFacingMessageContent(item.content);
    const attachments = item.attachments;
    return (
      <div className="message-bubble user-message">
        <div className="message-content">
          {attachments && attachments.length > 0 && (
            <div className="message-attachments">
              {attachments.map((f, i) => (
                <AttachmentChip key={i} attachment={f} />
              ))}
            </div>
          )}
          {content && <div className="message-text">{content}</div>}
        </div>
        <MessageMeta timestamp={item.timestamp} copyText={content} />
      </div>
    );
  }

  if (item.type === "assistant_message") {
    // 清理正文中的 token 级别空格（国内 LLM 普通回复也存在此问题）
    const cleanedContent = cleanReasoningText(item.content);
    const showMeta = item.phase !== "commentary";
    return (
      <div className={`message-bubble assistant-message`}>
        <div className="message-content">
          <div className="message-text">
            <MarkdownContent content={cleanedContent} />
          </div>
        </div>
        {showMeta && <MessageMeta timestamp={item.timestamp} copyText={normalizeVisibleMarkdown(cleanedContent)} />}
      </div>
    );
  }

  return null;
});

/** 消息气泡中的附件芯片 */
const AttachmentChip: React.FC<{ attachment: FileAttachment }> = ({ attachment }) => {
  const isImage = attachment.fileType === "image";
  if (isImage) {
    return <AttachmentImagePreview attachment={attachment} variant="message" />;
  }

  return (
    <div className={`msg-attach-chip ${attachment.fileType}`}>
      <Paperclip size={12} />
      <span className="msg-attach-chip-name">{attachment.fileName}</span>
    </div>
  );
};

function MessageMeta({ timestamp, copyText }: { timestamp: number; copyText: string }) {
  const { language } = useSettingsStore();
  const text = getAppText(language);

  return (
    <div className="message-meta">
      <span className="message-time">{formatMessageTime(timestamp)}</span>
      <CopyFeedbackButton
        className="message-copy-btn"
        copiedTitle={text.assistant.copied}
        textToCopy={copyText}
        title={text.assistant.copy}
      />
    </div>
  );
}

function formatMessageTime(timestamp: number): string {
  return new Date(timestamp).toLocaleTimeString([], {
    hour: "2-digit",
    minute: "2-digit",
  });
}
