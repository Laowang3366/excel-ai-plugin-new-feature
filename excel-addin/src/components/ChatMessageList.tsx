import type { DisplayTurn } from "../chat/chatPresentation";

interface Props {
  turns: DisplayTurn[];
}

export function ChatMessageList({ turns }: Props) {
  if (turns.length === 0) {
    return (
      <div className="chat-empty muted">
        输入问题开始只读查询。模型只能读取工作簿，不会写入或删除。
      </div>
    );
  }
  return (
    <div className="chat-messages" role="log" aria-live="polite">
      {turns.map((turn) => (
        <div key={turn.id} className="chat-turn">
          <div className="chat-bubble user">
            <div className="chat-role">你</div>
            <div className="chat-content">{turn.userText}</div>
          </div>
          {(turn.assistantText || turn.pending || turn.errorText) && (
            <div className={`chat-bubble assistant${turn.pending ? " pending" : ""}`}>
              <div className="chat-role">助手</div>
              <div className="chat-content">
                {turn.assistantText || (turn.pending ? "…" : "")}
              </div>
              {turn.errorText && (
                <div className="chat-error" role="alert">
                  {turn.errorText}
                </div>
              )}
              {turn.turnStatus && !turn.pending && turn.turnStatus !== "completed" && (
                <div className="chat-status muted">状态：{statusLabel(turn.turnStatus)}</div>
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
