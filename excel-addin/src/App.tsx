import { useEffect, useMemo, useState } from "react";
import {
  createHostAdapter,
  waitForOfficeReady,
  type HostAdapter,
  type HostKind,
} from "@shared/host";
import {
  getBrowserProviderPersistenceStorage,
  MemorySecretStore,
  ProviderStore,
} from "@shared/provider";
import {
  getBrowserPermissionModeStore,
  PERMISSION_MODE_LABELS,
  PERMISSION_MODES,
  type PermissionMode,
} from "@shared/agentChat";
import { ChatSessionProvider } from "./chat/ChatSessionContext";
import { ChatPanel } from "./components/ChatPanel";
import { HostStatusPanel } from "./components/HostStatusPanel";
import { ProviderSettingsPanel } from "./components/ProviderSettingsPanel";
import { ToolDemoPanel } from "./components/ToolDemoPanel";
import {
  ChartTaskPanel,
  CleanTaskPanel,
  FormulaTaskPanel,
  OcrTaskPanel,
  ReportTaskPanel,
} from "./components/tasks";
import { resolveInitialTabFromSearch, type AppTab } from "./resolveInitialTab";

const TAB_LABELS: { id: AppTab; label: string }[] = [
  { id: "chat", label: "聊天" },
  { id: "formula", label: "公式助手" },
  { id: "clean", label: "数据清洗" },
  { id: "ocr", label: "OCR识别" },
  { id: "chart", label: "图表制作" },
  { id: "report", label: "报告生成" },
  { id: "host", label: "宿主" },
  { id: "tools", label: "工具" },
  { id: "providers", label: "模型供应商" },
];

export function App() {
  const [tab, setTab] = useState<AppTab>(() =>
    resolveInitialTabFromSearch(
      typeof window !== "undefined" ? window.location.search : "",
    ),
  );
  const [hostKind, setHostKind] = useState<HostKind>("unknown");
  const [adapter, setAdapter] = useState<HostAdapter | null>(null);
  const providerStore = useMemo(
    () =>
      new ProviderStore(
        new MemorySecretStore(),
        getBrowserProviderPersistenceStorage(),
      ),
    [],
  );
  const permissionStore = useMemo(() => getBrowserPermissionModeStore(), []);
  const [permissionMode, setPermissionMode] = useState<PermissionMode>(() =>
    permissionStore.get(),
  );

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

  const appClassName = hostKind === "wps-jsa" ? "app app--wps-jsa" : "app";

  return (
    <div className={appClassName} data-host={hostKind}>
      <header>
        <h1>文格 Excel 加载项验证</h1>
        <p className="muted">
          独立 Office.js / WPS JSA 任务窗格 · 宿主{" "}
          <span className="badge">{hostKind}</span>
        </p>
      </header>

      <nav className="tabs" aria-label="主导航">
        {TAB_LABELS.map(({ id, label }) => (
          <button
            key={id}
            type="button"
            className={tab === id ? "active" : ""}
            onClick={() => setTab(id)}
          >
            {label}
          </button>
        ))}
      </nav>

      <ChatSessionProvider store={providerStore} adapter={adapter}>
        {tab === "chat" && (
          <>
            <div className="card permission-mode-bar">
              <label htmlFor="permission-mode-select">
                审批模式
                <select
                  id="permission-mode-select"
                  value={permissionMode}
                  onChange={(e) => {
                    const next = permissionStore.set(e.target.value);
                    setPermissionMode(next);
                  }}
                  aria-label="审批模式"
                >
                  {PERMISSION_MODES.map((mode) => (
                    <option key={mode} value={mode}>
                      {PERMISSION_MODE_LABELS[mode]}
                    </option>
                  ))}
                </select>
              </label>
              <p className="muted permission-mode-hint">
                {permissionMode === "normal" && "每次工具调用都需确认（含安全读取）"}
                {permissionMode === "auto_approve_safe" &&
                  "安全操作自动执行；中高风险需批准（默认）"}
                {permissionMode === "confirm_all" &&
                  "完整权限：所有工具自动执行，请谨慎使用"}
              </p>
            </div>
            <ChatPanel />
          </>
        )}
        {tab === "formula" && <FormulaTaskPanel />}
        {tab === "clean" && <CleanTaskPanel />}
        {tab === "ocr" && <OcrTaskPanel />}
        {tab === "chart" && <ChartTaskPanel />}
        {tab === "report" && <ReportTaskPanel />}
        {tab === "host" && adapter && <HostStatusPanel adapter={adapter} />}
        {tab === "tools" && adapter && <ToolDemoPanel adapter={adapter} />}
        {tab === "providers" && (
          <ProviderSettingsPanel store={providerStore} />
        )}
        {!adapter &&
          (tab === "host" || tab === "tools") && (
            <div className="card muted">正在检测宿主…</div>
          )}
      </ChatSessionProvider>
    </div>
  );
}
