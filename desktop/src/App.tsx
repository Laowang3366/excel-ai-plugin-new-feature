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
import { Sidebar } from "./components/Sidebar";
import { ChatPage } from "./components/ChatPage";
import type { SettingsSection } from "./components/SettingsPage";
import { AppTitlebar } from "./components/AppTitlebar";
import { ErrorBoundary } from "./components/common/ErrorBoundary";
import { HotPatchHealthAck } from "./components/HotPatchHealthAck";
import { getAppText } from "./i18n";
import { useExcelConnection } from "./hooks/useExcelConnection";
import { useOfficeConnection } from "./hooks/useOfficeConnection";
import { useWindowDisplayState } from "./hooks/useWindowDisplayState";

/** 主内容区可显示的页面 */
export type AppPage = "chat" | "settings";

const SettingsPage = React.lazy(() =>
  import("./components/SettingsPage").then((module) => ({ default: module.SettingsPage })),
);

export const App: React.FC = () => {
  const {
    isConfigured,
    isLoading,
    loadSettings,
    language,
    theme,
    officeAutoCompactEnabled,
    windowOpacity,
    setWindowOpacity,
  } = useSettingsStore();
  const [currentPage, setCurrentPage] = useState<AppPage>("chat");
  const [sidebarCollapsed, setSidebarCollapsed] = useState(false);
  const [settingsSidebarCollapsed, setSettingsSidebarCollapsed] = useState(false);
  const [settingsSection, setSettingsSection] = useState<SettingsSection>("general");
  const { excelStatus } = useExcelConnection();
  const { wordStatus, presentationStatus } = useOfficeConnection();
  const text = getAppText(language);
  const hasConnectedOffice =
    excelStatus.connected || wordStatus.connected || presentationStatus.connected;
  const { alwaysOnTop, displayMode, toggleAlwaysOnTop, toggleCompactMode } = useWindowDisplayState({
    autoCompactEnabled: officeAutoCompactEnabled,
    hasConnectedOffice,
  });
  const chatSidebarCollapsed = displayMode === "compact" || sidebarCollapsed;
  const settingsNavCollapsed = displayMode === "compact" || settingsSidebarCollapsed;

  useEffect(() => {
    loadSettings();
  }, [loadSettings]);

  useEffect(() => {
    document.documentElement.dataset.theme = theme;
    document.documentElement.lang = language;
  }, [language, theme]);

  // 首次未配置时自动跳转到设置页
  useEffect(() => {
    if (!isLoading && !isConfigured) {
      setCurrentPage("settings");
    }
  }, [isLoading, isConfigured]);

  if (isLoading) {
    return (
      <div className={`app-shell ${displayMode}-mode`}>
        <AppTitlebar
          alwaysOnTop={alwaysOnTop}
          displayMode={displayMode}
          onSetWindowOpacity={setWindowOpacity}
          onToggleAlwaysOnTop={toggleAlwaysOnTop}
          onToggleCompactMode={toggleCompactMode}
          showSidebarToggle={false}
          text={text}
          windowOpacity={windowOpacity}
        />
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
          <AppTitlebar
            alwaysOnTop={alwaysOnTop}
            collapsed={settingsNavCollapsed}
            displayMode={displayMode}
            onSetWindowOpacity={setWindowOpacity}
            onToggleAlwaysOnTop={toggleAlwaysOnTop}
            onToggleCompactMode={toggleCompactMode}
            onToggleSidebar={() => setSettingsSidebarCollapsed((collapsed) => !collapsed)}
            showSidebarToggle
            text={text}
            windowOpacity={windowOpacity}
          />
          <div className="app-view">
            <React.Suspense
              fallback={
                <div className="app-loading">
                  <div className="spinner" />
                  <p>{text.app.loading}</p>
                </div>
              }
            >
              <SettingsPage
                onBack={() => setCurrentPage("chat")}
                initialSection={settingsSection}
                sidebarCollapsed={settingsNavCollapsed}
              />
              <HotPatchHealthAck />
            </React.Suspense>
          </div>
        </div>
      </ErrorBoundary>
    );
  }

  // 聊天页：主侧边栏 + 聊天区
  return (
    <ErrorBoundary>
      <div className={`app-shell ${displayMode}-mode`}>
        <AppTitlebar
          alwaysOnTop={alwaysOnTop}
          collapsed={chatSidebarCollapsed}
          displayMode={displayMode}
          onSetWindowOpacity={setWindowOpacity}
          onToggleAlwaysOnTop={toggleAlwaysOnTop}
          onToggleCompactMode={toggleCompactMode}
          onToggleSidebar={() => setSidebarCollapsed((collapsed) => !collapsed)}
          showSidebarToggle
          showWindowModeToggle={false}
          text={text}
          windowOpacity={windowOpacity}
        />
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
              displayMode={displayMode}
              onToggleCompactMode={toggleCompactMode}
              onOpenSettings={(section = "general") => {
                setSettingsSection(section);
                setCurrentPage("settings");
              }}
            />
          </main>
        </div>
        <HotPatchHealthAck />
      </div>
    </ErrorBoundary>
  );
};
