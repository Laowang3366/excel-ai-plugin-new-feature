import type { ProviderPublicView, ProviderStore } from "@shared/provider";

export interface ActiveProviderSummary {
  present: boolean;
  name: string;
  apiFormat: string;
  connectionMode: "direct" | "gateway";
  keyLabel: string;
}

export function summarizeActiveProvider(
  store: ProviderStore,
): ActiveProviderSummary {
  const activeId = store.getActiveId();
  if (!activeId) {
    return {
      present: false,
      name: "未选择",
      apiFormat: "—",
      connectionMode: "direct",
      keyLabel: "—",
    };
  }
  const list = store.list();
  const view: ProviderPublicView | undefined = list.find((p) => p.id === activeId);
  if (!view) {
    return {
      present: false,
      name: "未选择",
      apiFormat: "—",
      connectionMode: "direct",
      keyLabel: "—",
    };
  }
  const keyLabel =
    view.connectionMode === "gateway"
      ? "无需浏览器 Key"
      : view.hasApiKey
        ? "Key 已设"
        : "Key 未设";
  return {
    present: true,
    name: view.name,
    apiFormat: view.apiFormat,
    connectionMode: view.connectionMode,
    keyLabel,
  };
}

interface Props {
  summary: ActiveProviderSummary;
}

export function ActiveProviderBar({ summary }: Props) {
  return (
    <div
      className="chat-active-provider"
      role="status"
      aria-label="当前活动模型供应商"
    >
      <span className="muted">活动供应商：</span>
      {summary.present ? (
        <span>
          <strong>{summary.name}</strong>
          <span className="badge">{summary.apiFormat}</span>
          <span className="badge">
            {summary.connectionMode === "gateway" ? "Gateway" : "直连"}
          </span>
          <span className="badge">{summary.keyLabel}</span>
        </span>
      ) : (
        <span className="muted">未选择 — 请到「模型供应商」页配置并设为当前</span>
      )}
    </div>
  );
}
