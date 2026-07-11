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

export const STREAMING_REASONING_BOTTOM_THRESHOLD = 24;
export const STREAMING_REASONING_LIVE_RENDER_LIMIT = 18_000;

export function getStreamingReasoningDistanceToBottom(metrics: {
  scrollHeight: number;
  scrollTop: number;
  clientHeight: number;
}): number {
  return Math.max(0, metrics.scrollHeight - metrics.scrollTop - metrics.clientHeight);
}

export function shouldFollowStreamingReasoning(distanceToBottom: number): boolean {
  return distanceToBottom <= STREAMING_REASONING_BOTTOM_THRESHOLD;
}

export function getStreamingReasoningVisibleText(
  reasoning: string,
  limit = STREAMING_REASONING_LIVE_RENDER_LIMIT,
): string {
  if (reasoning.length <= limit) return reasoning;
  const hiddenChars = reasoning.length - limit;
  return `...已暂存前 ${hiddenChars} 字，完整思考过程会在本轮结束后保存。\n\n${reasoning.slice(-limit)}`;
}

/** 流式推理过程展示 — 可折叠，始终可见 */
export const StreamingReasoning: React.FC<{ reasoning: string; autoCollapse?: boolean }> = ({
  reasoning,
  autoCollapse = false,
}) => {
  const { language } = useSettingsStore();
  const text = getAppText(language);
  const [reasoningExpanded, setReasoningExpanded] = useState(true);
  const contentRef = useRef<HTMLDivElement>(null);
  const shouldFollowReasoningRef = useRef(true);
  const previousReasoningLengthRef = useRef(0);

  useEffect(() => {
    if (reasoning) {
      setReasoningExpanded(!autoCollapse);
    }
  }, [reasoning, autoCollapse]);

  const visibleReasoning = useMemo(() => getStreamingReasoningVisibleText(reasoning), [reasoning]);
  const cleanedReasoning = useMemo(() => cleanReasoningText(visibleReasoning), [visibleReasoning]);

  useEffect(() => {
    if (reasoning.length < previousReasoningLengthRef.current) {
      shouldFollowReasoningRef.current = true;
    }
    previousReasoningLengthRef.current = reasoning.length;
  }, [reasoning]);

  const updateReasoningScrollState = () => {
    const node = contentRef.current;
    if (!node) return;
    const distanceToBottom = getStreamingReasoningDistanceToBottom({
      scrollHeight: node.scrollHeight,
      scrollTop: node.scrollTop,
      clientHeight: node.clientHeight,
    });
    shouldFollowReasoningRef.current = shouldFollowStreamingReasoning(distanceToBottom);
  };

  const toggleReasoningExpanded = () => {
    setReasoningExpanded((expanded) => {
      const nextExpanded = !expanded;
      if (nextExpanded) {
        shouldFollowReasoningRef.current = true;
      }
      return nextExpanded;
    });
  };

  useEffect(() => {
    if (!reasoningExpanded) return;
    if (!shouldFollowReasoningRef.current) return;
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
      <div className="reasoning-header" onClick={toggleReasoningExpanded}>
        <span className="reasoning-icon">
          <Brain size={14} />
        </span>
        <span className="reasoning-label">{text.assistant.thinking}</span>
        <span className="streaming-cursor">▊</span>
        <span className="reasoning-toggle">
          {reasoningExpanded ? <ChevronDown size={14} /> : <ChevronRight size={14} />}
        </span>
      </div>
      {reasoningExpanded && (
        <div
          className="reasoning-content"
          ref={contentRef}
          onScroll={updateReasoningScrollState}
          onWheelCapture={(event) => {
            if (event.deltaY < 0) {
              shouldFollowReasoningRef.current = false;
            }
          }}
        >
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
