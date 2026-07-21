import { useCallback, useMemo, useState, type KeyboardEvent } from "react";
import { useChatSession } from "../chat/ChatSessionContext";
import { useStickToBottom } from "../chat/useStickToBottom";
import { ActiveProviderBar, summarizeActiveProvider } from "./ActiveProviderBar";
import { ChatApprovalCard } from "./ChatApprovalCard";
import { ChatMessageList } from "./ChatMessageList";
import { ChatToolTrace } from "./ChatToolTrace";

export function ChatPanel() {
  const { view, send, retry, stop, clear, approve, reject, store, adapter } =
    useChatSession();
  const [draft, setDraft] = useState("");
  const [composing, setComposing] = useState(false);
  const activeSummary = useMemo(
    () => summarizeActiveProvider(store),
    [store, view.status, view.bannerError, view.turns.length],
  );

  const lastAssistantLen =
    view.turns.length > 0
      ? view.turns[view.turns.length - 1]?.assistantText.length ?? 0
      : 0;
  const { containerRef, onScroll } = useStickToBottom([
    view.turns.length,
    lastAssistantLen,
    view.pendingApproval?.requestId ?? null,
    view.liveAssistant.length,
  ]);

  const onSend = useCallback(async () => {
    if (!view.canSend) return;
    const text = draft;
    const outcome = await send(text);
    if (outcome.accepted) {
      setDraft("");
    } else if (outcome.restoreText != null) {
      setDraft(outcome.restoreText);
    }
  }, [draft, send, view.canSend]);

  const onRetry = useCallback(
    async (turnId: string) => {
      if (!view.canSend) return;
      await retry(turnId);
    },
    [retry, view.canSend],
  );

  const onKeyDown = (e: KeyboardEvent<HTMLTextAreaElement>) => {
    if (e.key !== "Enter" || e.shiftKey) return;
    if (composing || e.nativeEvent.isComposing) return;
    if (view.pendingApproval) return;
    e.preventDefault();
    void onSend();
  };

  if (!adapter) {
    return (
      <div className="card muted">正在检测宿主…聊天将在宿主就绪后可用。</div>
    );
  }

  return (
    <section className="card chat-panel">
      <div className="chat-banner" role="status">
        变更操作会在执行前等待你的批准；未批准不会写入/删除。任务入口与聊天共享同一会话。
      </div>

      <ActiveProviderBar summary={activeSummary} />

      {view.bannerError && (
        <div className="chat-banner error" role="alert">
          {view.bannerError}
        </div>
      )}

      <div
        className="chat-messages"
        role="log"
        aria-live="polite"
        ref={containerRef}
        onScroll={onScroll}
      >
        <ChatMessageList
          turns={view.turns}
          canRetry={view.canSend}
          onRetry={(id) => {
            void onRetry(id);
          }}
        />
      </div>

      {view.pendingApproval && (
        <ChatApprovalCard
          request={view.pendingApproval}
          disabled={!view.canApprove || view.status === "stopping"}
          onApprove={(id) => {
            approve(id);
          }}
          onReject={(id) => {
            reject(id);
          }}
        />
      )}

      <details className="chat-trace-box">
        <summary>工具轨迹</summary>
        <ChatToolTrace turns={view.turns} />
      </details>

      <div className="chat-composer">
        <textarea
          rows={3}
          value={draft}
          placeholder="输入问题… 写删改会弹出批准；Enter 发送，Shift+Enter 换行"
          onChange={(e) => setDraft(e.target.value)}
          onKeyDown={onKeyDown}
          onCompositionStart={() => setComposing(true)}
          onCompositionEnd={() => setComposing(false)}
          disabled={view.status === "stopping"}
          aria-label="聊天输入"
        />
        <div className="row chat-actions">
          <button
            type="button"
            onClick={() => {
              void onSend();
            }}
            disabled={!view.canSend || draft.trim() === ""}
          >
            发送
          </button>
          {view.canStop && (
            <button type="button" className="danger" onClick={stop}>
              停止
            </button>
          )}
          <button
            type="button"
            onClick={clear}
            disabled={!view.canClear}
            title="清空会话"
          >
            清空
          </button>
        </div>
      </div>
    </section>
  );
}
