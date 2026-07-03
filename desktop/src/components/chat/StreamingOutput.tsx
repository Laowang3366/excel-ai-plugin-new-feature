/**
 * 流式输出区域 — AI 正在生成时的实时展示
 *
 * 无头像设计，推理过程始终可见（可折叠）。
 * 拆分为 StreamingReasoning / StreamingContent 独立组件，
 * 以便嵌入助手组渲染的正确位置。
 */

import React, { useState, useEffect, useMemo, useRef } from "react";
import { Brain, ChevronDown, ChevronRight } from "../common/IconMap";
import { useSettingsStore } from "../../store/settingsStore";
import { cleanReasoningText } from "../../utils/textCleaner";
import { MarkdownContent } from "./MarkdownContent";
import { getAppText } from "../../i18n";

/** 流式推理过程展示 — 可折叠，始终可见 */
export const StreamingReasoning: React.FC<{ reasoning: string; autoCollapse?: boolean }> = ({
  reasoning,
  autoCollapse = false,
}) => {
  const { language } = useSettingsStore();
  const text = getAppText(language);
  const [reasoningExpanded, setReasoningExpanded] = useState(true);
  const contentRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    if (reasoning) {
      setReasoningExpanded(!autoCollapse);
    }
  }, [reasoning, autoCollapse]);

  const cleanedReasoning = useMemo(() => cleanReasoningText(reasoning), [reasoning]);

  useEffect(() => {
    if (!reasoningExpanded) return;
    window.requestAnimationFrame(() => {
      const node = contentRef.current;
      if (node) {
        node.scrollTop = node.scrollHeight;
      }
    });
  }, [reasoningExpanded, cleanedReasoning]);

  if (!reasoning) return null;

  return (
    <div className="streaming-reasoning">
      <div
        className="reasoning-header"
        onClick={() => setReasoningExpanded(!reasoningExpanded)}
      >
        <span className="reasoning-icon"><Brain size={14} /></span>
        <span className="reasoning-label">
          {text.assistant.thinking}
        </span>
        <span className="streaming-cursor">▊</span>
        <span className="reasoning-toggle">
          {reasoningExpanded ? <ChevronDown size={14} /> : <ChevronRight size={14} />}
        </span>
      </div>
      {reasoningExpanded && (
        <div className="reasoning-content" ref={contentRef}>
          <MarkdownContent content={cleanedReasoning} />
        </div>
      )}
    </div>
  );
};

/** 流式正文展示 */
export const StreamingContent: React.FC<{ content: string }> = ({ content }) => {
  const cleanedContent = useMemo(() => cleanReasoningText(content), [content]);

  if (!content) return null;

  return (
    <div className="streaming-message">
      <div className="message-content">
        <div className="message-text">
          <MarkdownContent content={cleanedContent} />
          <span className="streaming-cursor">▊</span>
        </div>
      </div>
    </div>
  );
};

/** 完整流式输出（兼容旧用法） */
interface StreamingOutputProps {
  content: string;
  reasoning: string;
}

export const StreamingOutput: React.FC<StreamingOutputProps> = ({
  content,
  reasoning,
}) => {
  return (
    <div className="streaming-output">
      <StreamingReasoning reasoning={reasoning} autoCollapse={Boolean(content)} />
      <StreamingContent content={content} />
    </div>
  );
};
