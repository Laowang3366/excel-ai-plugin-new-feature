/**
 * 工具确认弹窗 — 当工具调用需要用户确认时显示
 *
 * 对齐桌面端工具确认控件：
 * - 显示工具名称、参数、风险等级
 * - 确认 / 取消 按钮
 * - "始终允许此类操作" 快捷选项
 */

import React from "react";
import type { LucideIcon } from "lucide-react";
import type { ToolRiskLevel } from "../../../electron/agent/shared/types";
import { Wrench, ShieldCheck, ShieldAlert, ShieldX, Check, X } from "../common/IconMap";

export interface PendingToolCall {
  id: string;
  toolName: string;
  arguments: Record<string, unknown>;
  riskLevel: ToolRiskLevel;
  /** 工具描述 */
  description?: string;
  canAlwaysAllow?: boolean;
}

interface ToolConfirmDialogProps {
  pendingCall: PendingToolCall;
  onConfirm: (alwaysAllow?: boolean) => void;
  onCancel: () => void;
}

const RISK_CONFIG: Record<ToolRiskLevel, { label: string; color: string; bg: string; Icon: LucideIcon }> = {
  safe: { label: "安全", color: "var(--success-text)", bg: "var(--success-bg)", Icon: ShieldCheck },
  moderate: { label: "中等风险", color: "var(--warning-text)", bg: "var(--warning-bg)", Icon: ShieldAlert },
  dangerous: { label: "高风险", color: "var(--danger-text)", bg: "var(--danger-bg)", Icon: ShieldX },
};

export const ToolConfirmDialog: React.FC<ToolConfirmDialogProps> = ({
  pendingCall,
  onConfirm,
  onCancel,
}) => {
  const risk = RISK_CONFIG[pendingCall.riskLevel];
  const RiskIcon = risk.Icon;
  // 格式化参数显示
  const argEntries = Object.entries(pendingCall.arguments);
  const formatValue = (v: unknown): string => {
    if (typeof v === "string") return v.length > 80 ? v.slice(0, 80) + "…" : v;
    if (Array.isArray(v)) return `[${v.length} 项]`;
    if (typeof v === "object" && v !== null) return JSON.stringify(v).slice(0, 80);
    return String(v);
  };

  return (
    <div className="tool-confirm-overlay">
      <div className="tool-confirm-dialog">
        {/* 标题栏 */}
        <div className="tool-confirm-header">
          <span className="tool-confirm-icon"><Wrench size={16} /></span>
          <span className="tool-confirm-title">工具执行确认</span>
          <span
            className="tool-confirm-risk"
            style={{ color: risk.color, background: risk.bg }}
          >
            <RiskIcon size={13} /> {risk.label}
          </span>
        </div>

        {/* 工具信息 */}
        <div className="tool-confirm-body">
          <div className="tool-confirm-name">{pendingCall.toolName}</div>
          {pendingCall.description && (
            <div className="tool-confirm-desc">{pendingCall.description}</div>
          )}

          {/* 参数列表 */}
          {argEntries.length > 0 && (
            <div className="tool-confirm-args">
              <div className="tool-confirm-args-title">参数：</div>
              {argEntries.map(([key, value]) => (
                <div key={key} className="tool-confirm-arg-row">
                  <span className="tool-confirm-arg-key">{key}</span>
                  <span className="tool-confirm-arg-value">{formatValue(value)}</span>
                </div>
              ))}
            </div>
          )}
        </div>

        {/* 操作按钮 */}
        <div className="tool-confirm-actions">
          <button className="tool-confirm-cancel" onClick={onCancel}>
            <X size={13} /> 取消
          </button>
          {pendingCall.canAlwaysAllow === true && (
            <button
              className="tool-confirm-always"
              onClick={() => onConfirm(true)}
            >
              <Check size={13} /> 始终允许
            </button>
          )}
          <button
            className="tool-confirm-approve"
            onClick={() => onConfirm(false)}
          >
            <Check size={13} /> 确认执行
          </button>
        </div>
      </div>
    </div>
  );
};
