/**
 * 消息列表区 — 显示对话消息、流式输出、恢复提示、错误
 *
 * 从 ChatPage.tsx 提取的消息渲染区域，包含：
 * - 空状态提示
 * - 消息分组渲染（单条/助手组/流式助手组）
 * - 压缩通知
 * - 中断恢复提示
 * - 错误横幅
 *
 * 性能优化：
 * - 使用原子化 Zustand 选择器，避免流式 delta 触发全量重渲染
 * - 新消息和流式内容更新时自动滚动到最新位置
 * - 分组结果使用稳定的依赖键缓存
 */

import { useRef, useEffect, useMemo, useCallback, useState } from "react";
import { useChatStore } from "../../store/chatStore";
import { useSettingsStore } from "../../store/settingsStore";
import { getAppText } from "../../i18n";
import { groupAssistantItems } from "../../utils/chatHelpers";
import { MessageBubble } from "../chat/MessageBubble";
import { CompactionNotice } from "../chat/CompactionNotice";
import { ResumeHint } from "../chat/ResumeHint";
import { StreamingReasoning, StreamingContent } from "../chat/StreamingOutput";
import { AssistantGroupBlock, StreamingAssistantGroupBlock } from "../chat/AssistantGroupBlock";
import { AlertTriangle, ArrowDown, XCircle } from "../common/IconMap";

interface ChatMessageListProps {
  onFillInput: (text: string) => void;
}

export const AUTO_SCROLL_BOTTOM_THRESHOLD = 80;
export const JUMP_TO_LATEST_THRESHOLD = 180;
export const STREAM_AUTO_SCROLL_INTERVAL_MS = 1200;
export const USER_SCROLL_AUTO_FOLLOW_PAUSE_MS = 8000;
export const MAX_RENDERED_MESSAGE_ITEMS = 500;

export function getDistanceToBottom(metrics: {
  scrollHeight: number;
  scrollTop: number;
  clientHeight: number;
}): number {
  return Math.max(0, metrics.scrollHeight - metrics.scrollTop - metrics.clientHeight);
}

export function shouldAutoFollowLatest(distanceToBottom: number): boolean {
  return distanceToBottom <= AUTO_SCROLL_BOTTOM_THRESHOLD;
}

export function shouldShowJumpToLatest(distanceToBottom: number): boolean {
  return distanceToBottom > JUMP_TO_LATEST_THRESHOLD;
}

export function shouldRunScheduledAutoScroll(params: {
  isStreaming: boolean;
  shouldFollowLatest: boolean;
  userScrollPauseActive?: boolean;
}): boolean {
  return params.isStreaming && params.shouldFollowLatest && !params.userScrollPauseActive;
}

export function isUserScrollPauseActive(params: {
  now: number;
  lastUserScrollAwayAt: number;
  pauseMs?: number;
}): boolean {
  if (params.lastUserScrollAwayAt <= 0) return false;
  return (
    params.now - params.lastUserScrollAwayAt < (params.pauseMs ?? USER_SCROLL_AUTO_FOLLOW_PAUSE_MS)
  );
}

export function getVisibleMessageItems<T>(
  messages: T[],
  limit = MAX_RENDERED_MESSAGE_ITEMS,
): {
  visibleMessages: T[];
  hiddenCount: number;
} {
  if (messages.length <= limit) return { visibleMessages: messages, hiddenCount: 0 };
  return {
    visibleMessages: messages.slice(-limit),
    hiddenCount: messages.length - limit,
  };
}

function StreamingFallbackOutput() {
  const streamingContent = useChatStore((s) => s.streamingContent);
  const streamingReasoning = useChatStore((s) => s.streamingReasoning);
  const hasLiveContent = streamingReasoning || streamingContent;

  if (!hasLiveContent) return null;

  return (
    <div className="streaming-output">
      <StreamingReasoning reasoning={streamingReasoning} autoCollapse={Boolean(streamingContent)} />
      <StreamingContent content={streamingContent} />
    </div>
  );
}

export function ChatMessageList({ onFillInput }: ChatMessageListProps) {
  // 使用原子化选择器，只订阅组件实际依赖的状态片段
  const messages = useChatStore((s) => s.messages);
  const isStreaming = useChatStore((s) => s.isStreaming);
  const turnStatus = useChatStore((s) => s.turnStatus);
  const lastInterruptContext = useChatStore((s) => s.lastInterruptContext);
  const compactionNotice = useChatStore((s) => s.compactionNotice);
  const error = useChatStore((s) => s.error);
  const clearError = useChatStore((s) => s.clearError);

  const { language } = useSettingsStore();
  const text = getAppText(language);

  const messagesContainerRef = useRef<HTMLDivElement>(null);
  const messagesEndRef = useRef<HTMLDivElement>(null);
  const [showJumpToLatest, setShowJumpToLatest] = useState(false);
  const shouldFollowLatestRef = useRef(true);
  const lastUserScrollAwayAtRef = useRef(0);

  const updateScrollState = useCallback(() => {
    const node = messagesContainerRef.current;
    if (!node) return;
    const distanceToBottom = getDistanceToBottom(node);
    const nearBottom = shouldAutoFollowLatest(distanceToBottom);
    shouldFollowLatestRef.current = nearBottom;
    if (nearBottom) {
      lastUserScrollAwayAtRef.current = 0;
    } else if (isStreaming) {
      lastUserScrollAwayAtRef.current = Date.now();
    }
    setShowJumpToLatest((visible) => {
      const nextVisible = shouldShowJumpToLatest(distanceToBottom);
      return visible === nextVisible ? visible : nextVisible;
    });
  }, [isStreaming]);

  const pauseAutoFollowForUserScroll = useCallback(() => {
    if (!isStreaming) return;
    shouldFollowLatestRef.current = false;
    lastUserScrollAwayAtRef.current = Date.now();
    setShowJumpToLatest(true);
  }, [isStreaming]);

  const scrollToLatest = useCallback((behavior: ScrollBehavior = "smooth") => {
    window.requestAnimationFrame(() => {
      const node = messagesContainerRef.current;
      if (node) {
        node.scrollTo({ top: node.scrollHeight, behavior });
      } else {
        messagesEndRef.current?.scrollIntoView({ behavior, block: "end" });
      }
      shouldFollowLatestRef.current = true;
      lastUserScrollAwayAtRef.current = 0;
      setShowJumpToLatest(false);
    });
  }, []);

  const prevMessageCount = useRef(messages.length);
  useEffect(() => {
    if (
      messages.length > prevMessageCount.current &&
      (!isStreaming || shouldFollowLatestRef.current)
    ) {
      scrollToLatest("smooth");
    }
    prevMessageCount.current = messages.length;
  }, [messages.length, isStreaming, scrollToLatest]);

  useEffect(() => {
    if (!isStreaming) return;
    const timer = window.setInterval(() => {
      const userScrollPauseActive = isUserScrollPauseActive({
        now: Date.now(),
        lastUserScrollAwayAt: lastUserScrollAwayAtRef.current,
      });
      if (
        shouldRunScheduledAutoScroll({
          isStreaming,
          shouldFollowLatest: shouldFollowLatestRef.current,
          userScrollPauseActive,
        })
      ) {
        scrollToLatest("auto");
        return;
      }
      updateScrollState();
    }, STREAM_AUTO_SCROLL_INTERVAL_MS);
    return () => window.clearInterval(timer);
  }, [isStreaming, scrollToLatest, updateScrollState]);

  useEffect(() => {
    updateScrollState();
  }, [messages.length, turnStatus, updateScrollState]);

  const { visibleMessages, hiddenCount } = useMemo(
    () => getVisibleMessageItems(messages),
    [messages],
  );

  // 使用 messages.length + 最后一条消息 id 作为分组缓存依赖，
  // 比依赖整个 messages 数组更稳定（数组引用每次 store 更新都会变）
  const groupKey = useMemo(
    () =>
      `${visibleMessages.length}-${visibleMessages.length > 0 ? visibleMessages[visibleMessages.length - 1].id : "empty"}`,
    [visibleMessages],
  );
  const groups = useMemo(() => {
    return groupAssistantItems(visibleMessages);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [groupKey]);

  return (
    <div className="chat-messages-shell">
      <div
        className="chat-messages"
        ref={messagesContainerRef}
        onScroll={updateScrollState}
        onWheelCapture={(event) => {
          if (event.deltaY < 0) {
            pauseAutoFollowForUserScroll();
          }
        }}
      >
        {hiddenCount > 0 && (
          <div className="message-window-notice">
            已隐藏较早的 {hiddenCount} 条消息以保持界面流畅，完整记录仍保存在会话历史中。
          </div>
        )}
        {groups.map((group, gi) => {
          if (group.kind === "single") {
            const item = group.items[0];
            switch (item.type) {
              case "user_message":
                return <MessageBubble key={item.id} item={item} />;
              case "error":
                return (
                  <div key={item.id} className="error-bubble">
                    <AlertTriangle size={14} style={{ verticalAlign: "middle", marginRight: 4 }} />{" "}
                    {item.message}
                  </div>
                );
              case "compacted":
                return <CompactionNotice key={item.id} item={item} />;
              case "compact_progress":
                return <CompactionNotice key={item.id} item={item} />;
              default:
                return null;
            }
          }

          // 仅当最后一个分组是助手组时，才将其渲染为 StreamingAssistantGroupBlock。
          // 否则（例如第二轮对话，最后一个分组是新用户消息），助手组属于已完成轮次，
          // 应作为普通 AssistantGroupBlock 渲染，streaming 输出走 fallback。
          const isLastGroup = gi === groups.length - 1;
          if (isStreaming && isLastGroup && group.kind === "assistant") {
            return <StreamingAssistantGroupBlock key={`ag-${gi}`} group={group} />;
          }
          return <AssistantGroupBlock key={`ag-${gi}`} group={group} />;
        })}

        {/* streaming fallback：当最后一个分组不是助手组时（通常是新用户消息后），
          将流式输出渲染在所有消息之后，确保不会跑到用户消息上方 */}
        {isStreaming && (groups.length === 0 || groups[groups.length - 1].kind === "single") && (
          <StreamingFallbackOutput />
        )}

        {compactionNotice && <CompactionNotice message={compactionNotice} />}

        {lastInterruptContext && turnStatus === "interrupted" && (
          <ResumeHint message={lastInterruptContext} onFillInput={onFillInput} />
        )}

        {error && (
          <div className="error-banner" onClick={clearError}>
            <XCircle size={14} style={{ verticalAlign: "middle", marginRight: 4 }} /> {error}{" "}
            <span className="dismiss">{text.chat.dismiss}</span>
          </div>
        )}

        <div ref={messagesEndRef} />
      </div>
      {showJumpToLatest && (
        <button
          type="button"
          className="jump-to-latest-btn"
          onClick={() => scrollToLatest("smooth")}
          title="回到最新内容"
          aria-label="回到最新内容"
        >
          <ArrowDown size={18} />
        </button>
      )}
    </div>
  );
}
