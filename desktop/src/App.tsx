/**
 * 根组件 — 应用框架
 *
 * 布局结构（聊天页）：
 * ┌──────────┬─────────────────────┐
 * │ 侧边栏   │     主内容区         │
 * │          │                     │
 * │ 会话列表  │     聊天页面         │
 * │ 功能快捷  │                     │
 * │ 连接状态  │                     │
 * │ 设置入口  │                     │
 * └──────────┴─────────────────────┘
 *
 * 布局结构（设置页）：
 * ┌──────────┬─────────────────────┐
 * │ 设置侧栏  │    设置主视图        │
 * │ 返回工作台 │                     │
 * │ 常规设置   │                     │
 * │ 模型配置   │                     │
 * │ 使用统计   │                     │
 * └──────────┴─────────────────────┘
 *
 * 设置页是独立全页，不显示主侧边栏。
 */

import React, { useEffect, useState } from "react";
import { useSettingsStore } from "./store/settingsStore";
import { Sidebar, type IntentKind } from "./components/Sidebar";
import { ChatPage } from "./components/ChatPage";
import { SettingsPage, type SettingsSection } from "./components/SettingsPage";
import { ErrorBoundary } from "./components/common/ErrorBoundary";
import { ChevronLeft, Maximize2, Menu, Pin } from "./components/common/IconMap";
import { logWarn } from "./utils/rendererLogger";
import { getAppText } from "./i18n";
import { ipcApi } from "./services/ipcApi";
import { useExcelConnection } from "./hooks/useExcelConnection";
import { useOfficeConnection } from "./hooks/useOfficeConnection";
import type { WindowDisplayMode } from "./electronApi";

/** 主内容区可显示的页面 */
export type AppPage = "chat" | "settings";

function requestLayoutReflow(): void {
  void document.documentElement.offsetWidth;
  window.requestAnimationFrame(() => {
    void document.documentElement.offsetWidth;
  });
}

export const App: React.FC = () => {
  const {
    isConfigured,
    isLoading,
    loadSettings,
    language,
    theme,
    officeAutoCompactEnabled,
  } = useSettingsStore();
  const [currentPage, setCurrentPage] = useState<AppPage>("chat");
  const [sidebarCollapsed, setSidebarCollapsed] = useState(false);
  const [settingsSidebarCollapsed, setSettingsSidebarCollapsed] = useState(false);
  const [activeIntent, setActiveIntent] = useState<IntentKind>(null);
  const [settingsSection, setSettingsSection] = useState<SettingsSection>("general");
  const [alwaysOnTop, setAlwaysOnTop] = useState(true);
  const [displayMode, setDisplayMode] = useState<WindowDisplayMode>("normal");
  const { excelStatus } = useExcelConnection();
  const { wordStatus, presentationStatus } = useOfficeConnection();
  const text = getAppText(language);
  const hasConnectedOffice = excelStatus.connected || wordStatus.connected || presentationStatus.connected;
  const chatSidebarCollapsed = displayMode === "compact" || sidebarCollapsed;
  const settingsNavCollapsed = displayMode === "compact" || settingsSidebarCollapsed;

  useEffect(() => {
    loadSettings();
  }, [loadSettings]);

  useEffect(() => {
    document.documentElement.dataset.theme = theme;
    document.documentElement.lang = language;
  }, [language, theme]);

  useEffect(() => {
    ipcApi.window.getAlwaysOnTop()
      .then(setAlwaysOnTop)
      .catch(() => {
        logWarn("App", "获取窗口置顶状态失败，使用默认值");
        setAlwaysOnTop(true);
      });
  }, []);

  useEffect(() => {
    let disposed = false;
    let resizeSyncTimer: number | undefined;

    const applyActualMode = (mode: WindowDisplayMode, forceReflow = false) => {
      if (disposed) return;
      setDisplayMode(mode);
      if (forceReflow) {
        requestLayoutReflow();
      }
    };

    const syncActualMode = () => {
      ipcApi.window.getDisplayMode()
        .then((mode) => applyActualMode(mode))
        .catch(() => {
          logWarn("App", "获取窗口显示模式失败");
          applyActualMode("normal");
        });
    };

    const handleResize = () => {
      if (resizeSyncTimer !== undefined) {
        window.clearTimeout(resizeSyncTimer);
      }
      resizeSyncTimer = window.setTimeout(syncActualMode, 120);
    };

    const unsubscribeDisplayMode = ipcApi.window.onDisplayModeChanged((mode) => {
      applyActualMode(mode, true);
    });

    syncActualMode();
    window.addEventListener("resize", handleResize);

    return () => {
      disposed = true;
      if (resizeSyncTimer !== undefined) {
        window.clearTimeout(resizeSyncTimer);
      }
      unsubscribeDisplayMode();
      window.removeEventListener("resize", handleResize);
    };
  }, []);

  useEffect(() => {
    if (!officeAutoCompactEnabled || !hasConnectedOffice) return;

    const handleBlur = () => {
      setDisplayMode((currentMode) => {
        if (currentMode !== "normal") return currentMode;
        ipcApi.window.setDisplayMode("compact")
          .then(setDisplayMode)
          .catch(() => {
            logWarn("App", "设置紧凑模式失败");
            setDisplayMode(currentMode);
          });
        return "compact";
      });
    };

    window.addEventListener("blur", handleBlur);
    return () => window.removeEventListener("blur", handleBlur);
  }, [officeAutoCompactEnabled, hasConnectedOffice]);

  const toggleAlwaysOnTop = async () => {
    const next = !alwaysOnTop;
    setAlwaysOnTop(next);
    try {
      const actual = await ipcApi.window.setAlwaysOnTop(next);
      if (typeof actual === "boolean") {
        setAlwaysOnTop(actual);
      }
    } catch {
      logWarn("App", "切换窗口置顶失败");
      setAlwaysOnTop(alwaysOnTop);
    }
  };

  const setWindowMode = async (mode: WindowDisplayMode) => {
    setDisplayMode(mode);
    try {
      const actual = await ipcApi.window.setDisplayMode(mode);
      setDisplayMode(actual);
    } catch {
      logWarn("App", "切换窗口模式失败");
      setDisplayMode("normal");
    }
  };

  const toggleCompactMode = () => {
    setWindowMode(displayMode === "normal" ? "compact" : "normal");
  };

  const renderTitlebar = (
    showSidebarToggle: boolean,
    collapsed = false,
    onToggleSidebar?: () => void
  ) => (
    <div className="app-titlebar">
      {showSidebarToggle && (
        <button
          className="titlebar-sidebar-toggle"
          onClick={onToggleSidebar}
          title={collapsed ? text.app.expandSidebar : text.app.collapseSidebar}
        >
          <Menu size={16} />
        </button>
      )}
      <button
        className={`titlebar-window-mode-toggle ${displayMode === "compact" ? "active" : ""}`}
        onClick={toggleCompactMode}
        title={displayMode === "normal" ? text.app.compactWindow : text.app.restoreWindow}
        aria-pressed={displayMode === "compact"}
      >
        {displayMode === "normal" ? <ChevronLeft size={15} /> : <Maximize2 size={15} />}
      </button>
      <button
        className={`titlebar-pin-toggle ${alwaysOnTop ? "active" : ""}`}
        onClick={toggleAlwaysOnTop}
        title={alwaysOnTop ? text.app.pinOff : text.app.pinOn}
        aria-pressed={alwaysOnTop}
      >
        <Pin size={15} />
      </button>
    </div>
  );

  // 首次未配置时自动跳转到设置页
  useEffect(() => {
    if (!isLoading && !isConfigured) {
      setCurrentPage("settings");
    }
  }, [isLoading, isConfigured]);

  if (isLoading) {
    return (
      <div className={`app-shell ${displayMode}-mode`}>
        {renderTitlebar(false)}
        <div className="app-loading">
          <div className="spinner" />
          <p>{text.app.loading}</p>
        </div>
      </div>
    );
  }

  // 设置页：独立全页，不显示主侧边栏
  if (currentPage === "settings") {
    return (
      <ErrorBoundary>
        <div className={`app-shell ${displayMode}-mode`}>
          {renderTitlebar(true, settingsNavCollapsed, () => setSettingsSidebarCollapsed((collapsed) => !collapsed))}
          <div className="app-view">
            <SettingsPage
              onBack={() => setCurrentPage("chat")}
              initialSection={settingsSection}
              sidebarCollapsed={settingsNavCollapsed}
            />
          </div>
        </div>
      </ErrorBoundary>
    );
  }

  // 聊天页：主侧边栏 + 聊天区
  return (
    <ErrorBoundary>
      <div className={`app-shell ${displayMode}-mode`}>
        {renderTitlebar(true, chatSidebarCollapsed, () => setSidebarCollapsed((collapsed) => !collapsed))}
        <div className={`app ${chatSidebarCollapsed ? "sidebar-collapsed" : ""}`}>
          <Sidebar
            collapsed={chatSidebarCollapsed}
            currentPage={currentPage}
            onNavigate={setCurrentPage}
            onOpenSettingsSection={(section) => {
              setSettingsSection(section);
              setCurrentPage("settings");
            }}
          />
          <main className="app-main">
            <ChatPage
              onOpenSettings={(section = "general") => {
                setSettingsSection(section);
                setCurrentPage("settings");
              }}
              activeIntent={activeIntent}
              onIntentClick={setActiveIntent}
            />
          </main>
        </div>
      </div>
    </ErrorBoundary>
  );
};
