import { useCallback, useState, type KeyboardEvent } from "react";
import type { ChatController, ChatControllerDeps } from "@shared/agentChat";
import type { HostAdapter } from "@shared/host";
import type { ProviderStore } from "@shared/provider";
import { useChatController } from "../chat/useChatController";
import { ChatMessageList } from "./ChatMessageList";
import { ChatToolTrace } from "./ChatToolTrace";

interface Props {
  store: ProviderStore;
  adapter: HostAdapter | null;
  /** Test seam only; production omits this and uses real ChatController. */
  createController?: (deps: ChatControllerDeps) => ChatController;
}

export function ChatPanel({ store, adapter, createController }: Props) {
  const { view, send, stop, clear } = useChatController({
    store,
    adapter,
    createController,
  });
  const [draft, setDraft] = useState("");
  const [composing, setComposing] = useState(false);

  const onSend = useCallback(() => {
    if (!view.canSend) return;
    const text = draft;
    setDraft("");
    void send(text);
  }, [draft, send, view.canSend]);

  const onKeyDown = (e: KeyboardEvent<HTMLTextAreaElement>) => {
    if (e.key !== "Enter" || e.shiftKey) return;
    if (composing || e.nativeEvent.isComposing) return;
    e.preventDefault();
    onSend();
  };

  if (!adapter) {
    return <div className="card muted">正在检测宿主…聊天将在宿主就绪后可用。</div>;
  }

  return (
    <section className="card chat-panel">
      <div className="chat-banner" role="status">
        只读模式：当前仅可查询，不会写入或删除工作簿
      </div>

      {view.bannerError && (
        <div className="chat-banner error" role="alert">
          {view.bannerError}
        </div>
      )}

      <ChatMessageList turns={view.turns} />

      <details className="chat-trace-box">
        <summary>工具轨迹</summary>
        <ChatToolTrace turns={view.turns} />
      </details>

      <div className="chat-composer">
        <textarea
          rows={3}
          value={draft}
          placeholder="输入只读问题… Enter 发送，Shift+Enter 换行"
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
            onClick={onSend}
            disabled={!view.canSend || draft.trim() === ""}
          >
            发送
          </button>
          {view.canStop && (
            <button type="button" className="danger" onClick={stop}>
              停止
            </button>
          )}
          {view.status === "stopping" && (
            <span className="muted">
              正在停止…进行中的表格读取可能仍会完成
            </span>
          )}
          <button type="button" onClick={clear} disabled={!view.canClear}>
            清空
          </button>
        </div>
      </div>
    </section>
  );
}
