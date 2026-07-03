/**
 * 工具调用气泡 — 展示工具执行过程和结果
 */

import React, { useState } from "react";
import type { LucideIcon } from "lucide-react";
import type { ToolCallItem, ToolResultItem, TurnItem } from "../../electronApi";
import { useSettingsStore, type AppLanguage } from "../../store/settingsStore";
import { getAppText } from "../../i18n";
import {
  Clock,
  RefreshCw,
  CheckCircle,
  XCircle,
  BookOpen,
  PenLine,
  MousePointerClick,
  Hash,
  Wrench,
  ClipboardList,
  ChevronDown,
  ChevronRight,
  AlertTriangle,
  Check,
} from "../common/IconMap";

interface ToolCallBubbleProps {
  item: TurnItem;
  result?: ToolResultItem;
}

/**
 * 工具调用气泡 — 使用 React.memo 避免父组件重渲染时重复解析。
 */
export const ToolCallBubble: React.FC<ToolCallBubbleProps> = React.memo(({ item, result }) => {
  if (item.type === "tool_call") {
    return <ToolCallDisplay item={item} result={result} />;
  }

  if (item.type === "tool_result") {
    return <ToolResultDisplay item={item} />;
  }

  return null;
});

const STATUS_ICONS: Record<string, LucideIcon> = {
  pending: Clock,
  running: RefreshCw,
  completed: CheckCircle,
  failed: XCircle,
};

const TOOL_PREFIX_ICONS: Record<string, LucideIcon> = {
  "workbook.": BookOpen,
  "range.": PenLine,
  "selection.": MousePointerClick,
  "formula.": Hash,
  "vba.": Wrench,
  "sheet.": ClipboardList,
};

const ToolCallDisplay: React.FC<{ item: ToolCallItem; result?: ToolResultItem }> = ({ item, result }) => {
  const { language } = useSettingsStore();
  const [expanded, setExpanded] = useState(false);

  const displayStatus = getToolDisplayStatus(item, result);
  const StatusIcon = STATUS_ICONS[displayStatus] || XCircle;
  const statusLabel = getToolStatusLabel(displayStatus, language);

  // 简化工具名显示 — 用图标替代 emoji 前缀
  let displayToolName = item.toolName;
  let PrefixIcon: LucideIcon | null = null;
  for (const [prefix, icon] of Object.entries(TOOL_PREFIX_ICONS)) {
    if (item.toolName.startsWith(prefix)) {
      PrefixIcon = icon;
      displayToolName = item.toolName.slice(prefix.length);
      break;
    }
  }

  return (
    <div className={`tool-call-bubble status-${displayStatus}`}>
      <div
        className="tool-call-header"
        onClick={() => setExpanded(!expanded)}
      >
        <span className="tool-status"><StatusIcon size={14} /></span>
        <span className="tool-name">{statusLabel}</span>
        <span className="tool-toggle">{expanded ? <ChevronDown size={14} /> : <ChevronRight size={14} />}</span>
      </div>
      {expanded && (
        <div className="tool-detail-card">
          <div className="tool-command-line">
            {PrefixIcon && <PrefixIcon size={13} className="tool-command-icon" />}
            <span className="tool-command-prefix">$</span>
            <code>{displayToolName}</code>
          </div>
          <pre>{JSON.stringify(item.arguments, null, 2)}</pre>
          {result && (
            <pre>{formatToolResult(result)}</pre>
          )}
        </div>
      )}
    </div>
  );
};

function getToolDisplayStatus(item: ToolCallItem, result?: ToolResultItem): string {
  if (result?.isError || item.status === "failed") return "failed";
  if (result) return "completed";
  return item.status;
}

function getToolStatusLabel(status: string, language: AppLanguage): string {
  const text = getAppText(language).assistant.toolStatus;
  if (status === "failed") return text.failed;
  if (status === "running") return text.running;
  if (status === "pending") return text.pending;
  return text.completed;
}

function formatToolResult(item: ToolResultItem): string {
  const resultStr = typeof item.result === "string"
    ? item.result
    : JSON.stringify(item.result, null, 2);
  return item.isError ? resultStr : resultStr || "(completed with no output)";
}

const ToolResultDisplay: React.FC<{ item: ToolResultItem }> = ({ item }) => {
  const { language } = useSettingsStore();
  const text = getAppText(language);
  const [expanded, setExpanded] = useState(false);

  const resultStr = typeof item.result === "string"
    ? item.result
    : JSON.stringify(item.result, null, 2);

  // 截断过长的结果
  const isTruncated = resultStr.length > 500;
  const displayResult = isTruncated && !expanded
    ? resultStr.slice(0, 500) + "..."
    : resultStr;

  return (
    <div className={`tool-result-bubble ${item.isError ? "error" : "success"}`}>
      <div className="tool-result-header" onClick={() => setExpanded(!expanded)}>
        <span className="tool-result-icon">{item.isError ? <AlertTriangle size={14} /> : <Check size={14} />}</span>
        <span className="tool-result-label">
          {item.isError ? text.assistant.toolError : text.assistant.toolResult}
        </span>
        <span className="tool-result-toggle">
          {expanded ? <ChevronDown size={14} /> : <ChevronRight size={14} />}
        </span>
      </div>
      <div className="tool-detail-card tool-result-content">
        <pre>{displayResult}</pre>
      </div>
    </div>
  );
};
