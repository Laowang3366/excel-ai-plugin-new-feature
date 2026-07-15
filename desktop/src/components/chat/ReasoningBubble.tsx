/**
 * 推理气泡 — 思考过程/推理链，随时可查看
 *
 * 设计：
 * - 始终渲染（不受 showReasoning 开关控制）
 * - 默认折叠，只显示一行摘要 + 展开按钮
 * - 点击展开后完整展示推理原文
 * - 完成后的推理默认折叠，避免历史推理占满对话区域
 */

import React, { useState, useEffect, useMemo, useRef } from "react";
import type { ReasoningItem } from "../../electronApi";
import { Brain, ChevronDown, ChevronRight } from "../common/IconMap";
import { useSettingsStore } from "../../store/settingsStore";
import { cleanReasoningText } from "../../utils/textCleaner";
import { MarkdownContent } from "./MarkdownContent";
import { getAppText } from "../../i18n";

interface ReasoningBubbleProps {
  item: ReasoningItem;
  elapsedDuration?: string;
  defaultExpanded?: boolean;
}

/**
 * 推理气泡 — 使用 React.memo 避免父组件重渲染时重复执行 cleanReasoningText 和 markdown 解析。
 */
export const ReasoningBubble: React.FC<ReasoningBubbleProps> = React.memo(
  ({ item, elapsedDuration, defaultExpanded = false }) => {
    const { language } = useSettingsStore();
    const text = getAppText(language);
    const [expanded, setExpanded] = useState(defaultExpanded);
    const contentRef = useRef<HTMLDivElement>(null);

    useEffect(() => {
      setExpanded(defaultExpanded);
    }, [defaultExpanded, item.id]);

    const hasRawContent = item.rawContent.length > 0 && item.rawContent.some((s) => s.trim());
    const hasSummary = item.summaryText.length > 0 && item.summaryText.some((s) => s.trim());

    // 优先展示原始推理内容，其次展示摘要
    // 注意：rawContent 是 token 级别的 delta 数组，必须用 join("") 拼接
    // join("\n") 会在每个 delta 间插入换行，导致文本散架
    const rawText = hasRawContent
      ? item.rawContent.join("")
      : hasSummary
        ? item.summaryText.join("\n")
        : text.assistant.reasoningPending;

    // 清理推理文本中的 token 级别空格（国内 LLM 的 reasoning_content 常见问题）
    const displayContent = useMemo(() => cleanReasoningText(rawText), [rawText]);

    useEffect(() => {
      if (!expanded) return;
      window.requestAnimationFrame(() => {
        const node = contentRef.current;
        if (node) {
          node.scrollTop = node.scrollHeight;
        }
      });
    }, [expanded, displayContent]);

    return (
      <div className="reasoning-bubble">
        <div className="reasoning-header" onClick={() => setExpanded(!expanded)}>
          <span className="reasoning-icon">
            <Brain size={14} />
          </span>
          <span className="reasoning-label">{text.assistant.reasoning}</span>
          {elapsedDuration && <span className="reasoning-duration">{elapsedDuration}</span>}
          <span className="reasoning-toggle">
            {expanded ? <ChevronDown size={14} /> : <ChevronRight size={14} />}
          </span>
        </div>
        {expanded && (
          <div className="reasoning-content" ref={contentRef}>
            <MarkdownContent content={displayContent} />
          </div>
        )}
      </div>
    );
  },
);
