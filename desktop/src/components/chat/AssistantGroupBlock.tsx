/**
 * 助手消息组渲染 — 将连续的助手条目折叠/展开显示
 *
 * 从 ChatPage.tsx 提取，包含：
 * - sortItemsByRound: 按 agent 轮次排序
 * - AssistantGroupBlock: 已完成轮次的折叠组
 * - StreamingAssistantGroupBlock: 流式期间的展开组
 * - renderItem: 单个 TurnItem 渲染逻辑
 * - MessageBubbleIcon: 聊天头部图标（保留此处供 ChatPage 使用）
 */

import React, { useEffect, useMemo, useState } from "react";
import { useChatStore } from "../../store/chatStore";
import { useSettingsStore } from "../../store/settingsStore";
import type { AppLanguage, PermissionMode } from "../../store/settingsStore";
import type { TurnItem } from "../../electronApi";
import { getAppText } from "../../i18n";
import {
  getItemDurationSeconds,
  getLiveTurnDurationSeconds,
  sumDurations,
  formatDuration,
} from "../../utils/chatHelpers";
import { MessageBubble } from "./MessageBubble";
import { ReasoningBubble } from "./ReasoningBubble";
import { ToolCallBubble } from "./ToolCallBubble";
import { StreamingReasoning, StreamingContent } from "./StreamingOutput";
import { ChevronRight, ChevronDown } from "../common/IconMap";

// ============================================================
// 按 agent 轮次排序
// ============================================================

/**
 * 对 TurnItem[] 做防御性排序，确保从 JSONL 历史恢复时顺序正确。
 *
 * 后端 agentLoop 按 API 真实事件顺序发出 items：
 *   reasoning → assistant_message → tool_call → tool_result
 *   （每轮内可能有多组这样的序列，最后一轮无 tool_call）
 *
 * 前端 store 按事件到达顺序 push，实时数据顺序已经正确。
 * 此函数仅作为从 JSONL 历史恢复时的兜底，按轮次分组后组内排序。
 */
export function sortItemsByRound(items: TurnItem[]): TurnItem[] {
  if (items.length <= 1) return items;

  // 将 items 分成轮次：reasoning 标志新一轮开始
  const rounds: TurnItem[][] = [];
  let currentRound: TurnItem[] = [];

  for (const item of items) {
    // reasoning 且当前轮非空 → 新一轮
    if (item.type === "reasoning" && currentRound.length > 0) {
      rounds.push(currentRound);
      currentRound = [];
    }
    currentRound.push(item);
  }
  if (currentRound.length > 0) {
    rounds.push(currentRound);
  }

  // 每轮内排序：reasoning(0) → assistant_message(1) → tool_call(2) → tool_result(3)
  const TYPE_ORDER: Record<string, number> = {
    reasoning: 0,
    assistant_message: 1,
    tool_call: 2,
    tool_result: 3,
  };

  const sortedRounds = rounds.map((round) =>
    [...round].sort((a, b) => {
      const oa = TYPE_ORDER[a.type] ?? 99;
      const ob = TYPE_ORDER[b.type] ?? 99;
      if (oa !== ob) return oa - ob;
      return 0; // 同类型保持原有相对顺序
    })
  );

  return sortedRounds.flat();
}

// ============================================================
// AssistantGroupBlock — 使用 React.memo 避免流式重渲染波及已完成组
// ============================================================

/**
 * 比较两个 group 是否相等（基于 items 内部引用，避免流式增量导致已完成组重渲染）
 */
function assistantGroupEqual(
  prev: { group: { kind: "assistant"; items: TurnItem[]; previousUserTimestamp?: number }; isLatest?: boolean },
  next: { group: { kind: "assistant"; items: TurnItem[]; previousUserTimestamp?: number }; isLatest?: boolean }
): boolean {
  if (prev.isLatest !== next.isLatest) return false;
  const a = prev.group.items;
  const b = next.group.items;
  if (a.length !== b.length) return false;
  for (let i = 0; i < a.length; i++) {
    if (a[i] !== b[i]) return false;
  }
  return true;
}

export const AssistantGroupBlock = React.memo(function AssistantGroupBlock({
  group,
  isLatest = false,
}: {
  group: { kind: "assistant"; items: TurnItem[]; previousUserTimestamp?: number };
  /** 是否为最新的助手组（用于控制初始折叠状态） */
  isLatest?: boolean;
}) {
  const { language } = useSettingsStore();
  const text = getAppText(language);
  const sortedItems = useMemo(() => sortItemsByRound(group.items), [group.items]);
  const hasFinalMessage = sortedItems.some(
    (item) => item.type === "assistant_message" && item.phase === "final"
  );
  const itemDurations = useMemo(
    () => getItemDurationSeconds(sortedItems, group.previousUserTimestamp),
    [sortedItems, group.previousUserTimestamp]
  );
  const formattedElapsed = formatDuration(sumDurations(itemDurations), language);

  // 已完成轮次：有 final 正文时默认折叠，否则展开
  const [detailsCollapsed, setDetailsCollapsed] = useState(hasFinalMessage);

  // 预计算已被 tool_call 关联的 tool_result ID 集合，避免渲染期间 Mutation
  const renderedToolResultIds = useMemo(() => {
    const ids = new Set<string>();
    for (const item of sortedItems) {
      if (item.type === "tool_call") {
        const result = sortedItems.find(
          (c) => c.type === "tool_result" && c.toolCallId === item.id
        );
        if (result?.type === "tool_result") {
          ids.add(result.id);
        }
      }
    }
    return ids;
  }, [sortedItems]);

  // 最新组：从流式（展开）切换到完成后，自动折叠
  useEffect(() => {
    if (isLatest && hasFinalMessage) {
      setDetailsCollapsed(true);
    }
  }, [isLatest, hasFinalMessage]);

  return (
    <div className="assistant-group">
      <button
        className="assistant-work-summary"
        onClick={() => setDetailsCollapsed((collapsed) => !collapsed)}
        type="button"
        title={detailsCollapsed ? text.chat.expandTurnDetails : text.chat.collapseTurnDetails}
      >
        <span className="assistant-work-label">{text.chat.turnDuration}</span>
        {formattedElapsed && (
          <span className="assistant-work-duration">{formattedElapsed}</span>
        )}
        <span className="assistant-work-toggle">
          {detailsCollapsed ? <ChevronRight size={14} /> : <ChevronDown size={14} />}
        </span>
      </button>

      {sortedItems.map((item) =>
        renderItem(item, { ...group, items: sortedItems }, detailsCollapsed, renderedToolResultIds, itemDurations, language)
      )}
    </div>
  );
}, assistantGroupEqual);

// ============================================================
// StreamingAssistantGroupBlock
// ============================================================

/**
 * 流式期间的助手组渲染 — 按真实事件顺序显示。
 *
 * 已完成轮次的 items 顺序（后端已按事件顺序发出）：
 *   assistant_message → reasoning → tool_call → tool_result
 * 当前流式轮次：streamingReasoning（思考）→ streamingContent（正文）→ tool_call（尚未到达）
 *
 * 渲染顺序：
 * 1. 已完成轮次的 items（按原始顺序）
 * 2. streamingReasoning（当前轮次的思考过程）
 * 3. streamingContent（当前轮次的正文片段）
 *
 * 使用 React.memo + 自定义比较器，仅当 group.items 内部引用变化时才重渲染。
 */
export const StreamingAssistantGroupBlock = React.memo(function StreamingAssistantGroupBlock({
  group,
}: {
  group: { kind: "assistant"; items: TurnItem[]; previousUserTimestamp?: number };
}) {
  const { language } = useSettingsStore();
  const text = getAppText(language);
  const streamingReasoning = useChatStore((s) => s.streamingReasoning);
  const streamingContent = useChatStore((s) => s.streamingContent);
  const [nowTimestamp, setNowTimestamp] = useState(() => Date.now());
  const sortedItems = useMemo(() => sortItemsByRound(group.items), [group.items]);
  const itemDurations = useMemo(
    () => getItemDurationSeconds(sortedItems, group.previousUserTimestamp),
    [sortedItems, group.previousUserTimestamp]
  );
  const liveElapsedSeconds = useMemo(
    () => getLiveTurnDurationSeconds(sortedItems, group.previousUserTimestamp, nowTimestamp),
    [sortedItems, group.previousUserTimestamp, nowTimestamp]
  );
  const formattedElapsed = formatDuration(liveElapsedSeconds, language);

  useEffect(() => {
    const timer = window.setInterval(() => {
      setNowTimestamp(Date.now());
    }, 1000);
    return () => window.clearInterval(timer);
  }, []);

  // 预计算已被 tool_call 关联的 tool_result ID 集合，避免渲染期间 Mutation
  const renderedToolResultIds = useMemo(() => {
    const ids = new Set<string>();
    for (const item of sortedItems) {
      if (item.type === "tool_call") {
        const result = sortedItems.find(
          (c) => c.type === "tool_result" && c.toolCallId === item.id
        );
        if (result?.type === "tool_result") {
          ids.add(result.id);
        }
      }
    }
    return ids;
  }, [sortedItems]);

  // 流式期间始终展开已完成轮次的所有内容（reasoning + commentary + tools），
  // 让用户实时看到完整进度。Turn 完成后切换到 AssistantGroupBlock 时才自动折叠。
  return (
    <div className="assistant-group streaming">
      <button
        className="assistant-work-summary"
        onClick={() => {}}
        type="button"
      >
        <span className="assistant-work-label">{text.chat.turnDuration}</span>
        {formattedElapsed && (
          <span className="assistant-work-duration">{formattedElapsed}</span>
        )}
        <span className="assistant-work-toggle">
          <ChevronDown size={14} />
        </span>
      </button>

      {/* 已完成轮次的 items — 流式期间全部展开 */}
      {sortedItems.map((item) =>
        renderItem(item, { ...group, items: sortedItems }, false, renderedToolResultIds, itemDurations, language)
      )}

      {/* 当前流式轮次：思考过程在前，正文片段在后 — 始终可见 */}
      {streamingReasoning && (
        <StreamingReasoning
          reasoning={streamingReasoning}
          autoCollapse={Boolean(streamingContent)}
        />
      )}
      {streamingContent && <StreamingContent content={streamingContent} />}
    </div>
  );
}, assistantGroupEqual);

// ============================================================
// renderItem — 单个 TurnItem 渲染
// ============================================================

/** 渲染单个 TurnItem（从 AssistantGroupBlock 提取的公共逻辑） */
function renderItem(
  item: TurnItem,
  group: { kind: "assistant"; items: TurnItem[]; previousUserTimestamp?: number },
  detailsCollapsed: boolean,
  renderedToolResultIds: Set<string>,
  itemDurations: Map<string, number>,
  language: AppLanguage,
): React.ReactNode {
  switch (item.type) {
    case "reasoning":
      return detailsCollapsed ? null : (
        <ReasoningBubble
          key={item.id}
          item={item}
          elapsedDuration={formatDuration(itemDurations.get(item.id), language)}
        />
      );
    case "assistant_message":
      // 折叠时只显示 final 总结正文，commentary 阶段性正文也一起折叠
      if (detailsCollapsed && item.phase === "commentary") return null;
      return <MessageBubble key={item.id} item={item} />;
    case "tool_call": {
      if (detailsCollapsed) return null;
      const result = group.items.find(
        (candidate) => candidate.type === "tool_result" && candidate.toolCallId === item.id
      );
      if (result?.type === "tool_result") {
        // renderedToolResultIds 已由 useMemo 预计算，此处不再 mutation
        return <ToolCallBubble key={item.id} item={item} result={result} />;
      }
      return <ToolCallBubble key={item.id} item={item} />;
    }
    case "tool_result":
      if (detailsCollapsed || renderedToolResultIds.has(item.id)) return null;
      return <ToolCallBubble key={item.id} item={item} />;
    default:
      return null;
  }
}
