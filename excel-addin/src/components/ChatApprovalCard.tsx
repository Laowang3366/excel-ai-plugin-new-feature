import { useEffect, useId, useRef } from "react";
import type { ApprovalRequest } from "@shared/agentChat";
import { safeJson } from "../chat/chatPresentation";

interface Props {
  request: ApprovalRequest;
  disabled?: boolean;
  onApprove: (requestId: string) => void;
  onReject: (requestId: string) => void;
}

export function ChatApprovalCard({
  request,
  disabled = false,
  onApprove,
  onReject,
}: Props) {
  const titleId = useId();
  const descId = useId();
  const rejectRef = useRef<HTMLButtonElement>(null);
  const acted = useRef(false);

  useEffect(() => {
    acted.current = false;
    rejectRef.current?.focus();
  }, [request.requestId]);

  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (e.key === "Escape") {
        e.preventDefault();
        e.stopPropagation();
        if (!disabled && !acted.current) {
          acted.current = true;
          onReject(request.requestId);
        }
      }
    };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, [disabled, onReject, request.requestId]);

  const riskLabel =
    request.riskLevel === "dangerous"
      ? "危险"
      : request.riskLevel === "moderate"
        ? "需批准"
        : request.riskLevel;
  const previewText = safeJson(request.argsPreview);

  const clickApprove = () => {
    if (disabled || acted.current) return;
    acted.current = true;
    onApprove(request.requestId);
  };
  const clickReject = () => {
    if (disabled || acted.current) return;
    acted.current = true;
    onReject(request.requestId);
  };

  return (
    <div
      className="chat-approval-card"
      role="alertdialog"
      aria-modal="true"
      aria-labelledby={titleId}
      aria-describedby={descId}
    >
      <div className="chat-approval-header">
        <h3 id={titleId}>等待批准：{request.name}</h3>
        <div className="chat-approval-tags">
          <span className={`risk-badge risk-${request.riskLevel}`}>{riskLabel}</span>
          {request.destructive && (
            <span className="risk-badge risk-destructive">破坏性操作</span>
          )}
        </div>
      </div>
      <p id={descId} className="chat-approval-impact">
        {request.impactHint}
      </p>
      <pre className="chat-approval-args" aria-label="参数预览">
        {previewText}
      </pre>
      <div className="chat-approval-actions">
        <button
          ref={rejectRef}
          type="button"
          className="danger"
          disabled={disabled}
          onClick={clickReject}
        >
          拒绝
        </button>
        <button type="button" disabled={disabled} onClick={clickApprove}>
          批准
        </button>
      </div>
    </div>
  );
}
