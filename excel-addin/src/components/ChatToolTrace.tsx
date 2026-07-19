import type { DisplayTraceItem, DisplayTurn } from "../chat/chatPresentation";

interface Props {
  turns: DisplayTurn[];
}

export function ChatToolTrace({ turns }: Props) {
  const items: Array<DisplayTraceItem & { turnId: string }> = [];
  for (const turn of turns) {
    for (const tr of turn.traces) {
      items.push({ ...tr, turnId: turn.id });
    }
  }
  if (items.length === 0) {
    return <div className="chat-trace muted">暂无工具轨迹</div>;
  }
  return (
    <ul className="chat-trace-list" aria-label="工具轨迹">
      {items.map((item) => (
        <li
          key={`${item.turnId}-${item.id}`}
          className={`chat-trace-item tone-${item.tone ?? "info"}`}
        >
          {item.text}
        </li>
      ))}
    </ul>
  );
}
