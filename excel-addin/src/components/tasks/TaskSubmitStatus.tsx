export function TaskSubmitStatus({
  busy,
  error,
  lastResult,
}: {
  busy: boolean;
  error?: string;
  lastResult?: string;
}) {
  return (
    <div className="task-status" role="status" aria-live="polite">
      {busy && <p className="muted">任务运行中…可在「聊天」查看完整会话与批准。</p>}
      {error && (
        <p className="chat-error" role="alert">
          {error}
        </p>
      )}
      {lastResult && !busy && <p className="muted">{lastResult}</p>}
    </div>
  );
}
