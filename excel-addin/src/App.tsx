import { useEffect, useMemo, useState } from "react";
import {
  createHostAdapter,
  waitForOfficeReady,
  type HostAdapter,
  type HostKind,
} from "@shared/host";
import { ProviderStore } from "@shared/provider";
import { ChatPanel } from "./components/ChatPanel";
import { HostStatusPanel } from "./components/HostStatusPanel";
import { ProviderSettingsPanel } from "./components/ProviderSettingsPanel";
import { ToolDemoPanel } from "./components/ToolDemoPanel";

type Tab = "chat" | "host" | "tools" | "providers";

export function App() {
  const [tab, setTab] = useState<Tab>("chat");
  const [hostKind, setHostKind] = useState<HostKind>("unknown");
  const [adapter, setAdapter] = useState<HostAdapter | null>(null);
  // Single shared store for the task pane lifetime.
  const providerStore = useMemo(() => new ProviderStore(), []);

  useEffect(() => {
    let cancelled = false;
    void waitForOfficeReady().then((kind) => {
      if (cancelled) return;
      setHostKind(kind);
      setAdapter(createHostAdapter(kind === "unknown" ? "office-js" : kind));
    });
    return () => {
      cancelled = true;
    };
  }, []);

  return (
    <div className="app">
      <header>
        <h1>文格 Excel 加载项验证</h1>
        <p className="muted">
          独立 Office.js / WPS JSA 任务窗格 · 宿主{" "}
          <span className="badge">{hostKind}</span>
        </p>
      </header>

      <nav className="tabs">
        <button
          type="button"
          className={tab === "chat" ? "active" : ""}
          onClick={() => setTab("chat")}
        >
          聊天
        </button>
        <button
          type="button"
          className={tab === "host" ? "active" : ""}
          onClick={() => setTab("host")}
        >
          宿主
        </button>
        <button
          type="button"
          className={tab === "tools" ? "active" : ""}
          onClick={() => setTab("tools")}
        >
          工具
        </button>
        <button
          type="button"
          className={tab === "providers" ? "active" : ""}
          onClick={() => setTab("providers")}
        >
          模型供应商
        </button>
      </nav>

      {tab === "chat" && <ChatPanel store={providerStore} adapter={adapter} />}
      {tab === "host" && adapter && <HostStatusPanel adapter={adapter} />}
      {tab === "tools" && adapter && <ToolDemoPanel adapter={adapter} />}
      {tab === "providers" && <ProviderSettingsPanel store={providerStore} />}
      {!adapter && tab !== "chat" && tab !== "providers" && (
        <div className="card muted">正在检测宿主…</div>
      )}
    </div>
  );
}
