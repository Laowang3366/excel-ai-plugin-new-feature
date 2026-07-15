import type { OfficeAutomationDocument, OfficeAutomationResult } from "../../electronApi";
import {
  AlertTriangle,
  Files,
  History,
  LayoutTemplate,
  RefreshCw,
  Workflow,
  X,
} from "../common/IconMap";
import { officeStatusLabel } from "./officeAutomationViewModel";

export type AutomationTab = "documents" | "workflows" | "transactions" | "templates";

export const AUTOMATION_TABS = [
  { id: "documents" as const, label: "文档与对象", icon: Files },
  { id: "workflows" as const, label: "工作流", icon: Workflow },
  { id: "transactions" as const, label: "事务", icon: History },
  { id: "templates" as const, label: "模板", icon: LayoutTemplate },
];

export function AutomationTabList({
  tab,
  onChange,
}: {
  tab: AutomationTab;
  onChange: (tab: AutomationTab) => void;
}) {
  return (
    <div className="office-automation-tabs" role="tablist" aria-label="Office 自动化管理">
      {AUTOMATION_TABS.map((item) => (
        <button
          key={item.id}
          type="button"
          role="tab"
          aria-label={item.label}
          title={item.label}
          aria-selected={tab === item.id}
          className={tab === item.id ? "active" : ""}
          onClick={() => onChange(item.id)}
        >
          <item.icon size={15} />
          <span>{item.label}</span>
        </button>
      ))}
    </div>
  );
}

export function AutomationErrorBanner({ error, onClose }: { error: string; onClose: () => void }) {
  if (!error) return null;
  return (
    <div className="office-automation-error" role="alert">
      <AlertTriangle size={14} />
      <span>{error}</span>
      <button type="button" title="关闭" onClick={onClose}>
        <X size={14} />
      </button>
    </div>
  );
}

export function Toolbar({
  title,
  count,
  onRefresh,
  busy,
}: {
  title: string;
  count: number;
  onRefresh: () => void;
  busy: boolean;
}) {
  return (
    <div className="office-view-toolbar">
      <div>
        <strong>{title}</strong>
        <span className="office-count">{count}</span>
      </div>
      <button
        type="button"
        className="office-icon-button"
        title="刷新"
        onClick={onRefresh}
        disabled={busy}
      >
        <RefreshCw size={15} className={busy ? "spin" : ""} />
      </button>
    </div>
  );
}

export function EmptyState({ text }: { text: string }) {
  return <div className="office-empty-state">{text}</div>;
}

export function StatusBadge({ status }: { status: string }) {
  return <span className={`office-status-badge ${status}`}>{officeStatusLabel(status)}</span>;
}

export function documentKey(document?: OfficeAutomationDocument): string {
  return document
    ? `${document.instanceId}|${document.fullName || document.name}|${document.index}`
    : "";
}

export function unwrap<T>(response: OfficeAutomationResult<T>): T {
  if (!response.success || response.data === undefined)
    throw new Error(response.error || "Office 自动化操作失败");
  return response.data;
}

export function errorText(reason: unknown): string {
  return reason instanceof Error ? reason.message : String(reason);
}
