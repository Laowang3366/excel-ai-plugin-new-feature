import type { DisplayTurn } from "../chat/chatPresentation";

interface Props {
  turns: DisplayTurn[];
  canRetry: boolean;
  onRetry?: (turnId: string) => void;
}

export function isRetryableTurn(turn: DisplayTurn): boolean {
  if (turn.pending) return false;
  if (turn.turnStatus === "failed" || turn.turnStatus === "aborted") return true;
  if (turn.errorText && turn.turnStatus !== "completed") return true;
  return false;
}

export function ChatMessageList({ turns, canRetry, onRetry }: Props) {
  if (turns.length === 0) {
    return (
      <div className="chat-empty muted">
        输入问题开始对话。变更类工具会在执行前请求你的批准。
      </div>
    );
  }
  return (
    <div className="chat-messages-inner">
      {turns.map((turn) => (
        <div key={turn.id} className="chat-turn">
          <div className="chat-bubble user">
            <div className="chat-role">你</div>
            <div className="chat-content">{turn.userText}</div>
          </div>
          {(turn.assistantText || turn.pending || turn.errorText) && (
            <div
              className={`chat-bubble assistant${turn.pending ? " pending" : ""}`}
            >
              <div className="chat-role">助手</div>
              <div className="chat-content">
                {turn.assistantText || (turn.pending ? "…" : "")}
              </div>
              {turn.errorText && (
                <div className="chat-error" role="alert">
                  {turn.errorText}
                </div>
              )}
              {turn.turnStatus &&
                !turn.pending &&
                turn.turnStatus !== "completed" && (
                  <div className="chat-status muted">
                    状态：{statusLabel(turn.turnStatus)}
                  </div>
                )}
              {canRetry && isRetryableTurn(turn) && onRetry && (
                <div className="chat-retry-row">
                  <button
                    type="button"
                    className="chat-retry-btn"
                    onClick={() => onRetry(turn.id)}
                    aria-label={`重试：${turn.userText.slice(0, 40)}`}
                  >
                    重试
                  </button>
                </div>
              )}
            </div>
          )}
        </div>
      ))}
    </div>
  );
}

function statusLabel(status: string): string {
  switch (status) {
    case "aborted":
      return "已停止";
    case "max_rounds":
      return "达到轮数上限";
    case "failed":
    case "preflight_failed":
      return "失败";
    case "busy":
      return "忙碌";
    case "empty":
      return "空输入";
    default:
      return status;
  }
}
