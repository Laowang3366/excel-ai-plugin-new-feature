/**
 * 设置页面 — 独立全页布局（侧栏 + 主视图）
 *
 * 对齐用户需求：
 * - Settings 作为独立全页，有自己的侧栏导航
 * - 侧栏项：返回工作台、常规设置、模型配置、使用统计
 * - 主视图根据选中项显示对应模块
 */

import React, { useState, useEffect } from "react";
import { useSettingsStore } from "../store/settingsStore";
import {
  ArrowLeft,
  User,
  Wrench,
  Bot,
  BarChart3,
  ShieldAlert,
  BookOpen,
  Package,
} from "./common/IconMap";
import { GeneralSettings } from "./settings/GeneralSettings";
import { ModelSettings } from "./settings/ModelSettings";
import { UsageStats } from "./settings/UsageStats";
import { ExecPolicySettings } from "./settings/ExecPolicySettings";
import { KnowledgeSettings } from "./settings/KnowledgeSettings";
import { OpenSourceSettings } from "./settings/OpenSourceSettings";

export type SettingsSection = "profile" | "general" | "model" | "usage" | "safety" | "knowledge" | "opensource";

const SETTINGS_TEXT = {
  "zh-CN": {
    back: "返回主页",
    loading: "加载设置中...",
    sections: {
      profile: "个人资料",
      general: "常规设置",
      model: "模型配置",
      usage: "使用统计",
      safety: "安全策略",
      knowledge: "知识库",
      opensource: "开源项目",
    },
    profileTitle: "个人资料",
    profileDesc: "查看当前桌面端账户信息。",
    localUser: "本地用户",
    localAccount: "桌面端本地账户",
  },
  "en-US": {
    back: "Back home",
    loading: "Loading settings...",
    sections: {
      profile: "Profile",
      general: "General",
      model: "Models",
      usage: "Usage",
      safety: "Safety",
      knowledge: "Knowledge",
      opensource: "Open Source",
    },
    profileTitle: "Profile",
    profileDesc: "View the current desktop account information.",
    localUser: "Local user",
    localAccount: "Desktop local account",
  },
} as const;

const SECTIONS = [
  { key: "profile" as const, icon: User },
  { key: "general" as const, icon: Wrench },
  { key: "model" as const, icon: Bot },
  { key: "usage" as const, icon: BarChart3 },
  { key: "safety" as const, icon: ShieldAlert },
  { key: "knowledge" as const, icon: BookOpen },
  { key: "opensource" as const, icon: Package },
];

interface SettingsPageProps {
  onBack?: () => void;
  initialSection?: SettingsSection;
  sidebarCollapsed?: boolean;
}

export const SettingsPage: React.FC<SettingsPageProps> = ({ onBack, initialSection = "general", sidebarCollapsed = false }) => {
  const [section, setSection] = useState<SettingsSection>(initialSection);
  const { isLoading, language, loadSettings } = useSettingsStore();
  const text = SETTINGS_TEXT[language];

  useEffect(() => {
    loadSettings();
  }, [loadSettings]);

  useEffect(() => {
    setSection(initialSection);
  }, [initialSection]);

  if (isLoading) {
    return (
      <div className="settings-shell loading" lang={language}>
        <div className="spinner" />
        <p>{text.loading}</p>
      </div>
    );
  }

  return (
    <div className={`settings-shell ${sidebarCollapsed ? "settings-sidebar-collapsed" : ""}`} lang={language}>
      {/* 设置侧栏 */}
      <aside className="settings-sidebar">
        <button className="settings-back-btn" onClick={onBack} title={text.back}>
          <ArrowLeft size={16} />
          <span>{text.back}</span>
        </button>

        <nav className="settings-sidebar-nav">
          {SECTIONS.map((s) => (
            <button
              key={s.key}
              className={`settings-sidebar-item ${section === s.key ? "active" : ""}`}
              onClick={() => setSection(s.key)}
              title={text.sections[s.key]}
            >
              <s.icon size={16} />
              <span>{text.sections[s.key]}</span>
            </button>
          ))}
        </nav>
      </aside>

      {/* 设置主视图 */}
      <main className="settings-main">
        {section === "profile" && <ProfileSettings text={text} />}
        {section === "general" && <GeneralSettings />}
        {section === "model" && <ModelSettings />}
        {section === "usage" && <UsageStats />}
        {section === "safety" && <ExecPolicySettings />}
        {section === "knowledge" && <KnowledgeSettings />}
        {section === "opensource" && <OpenSourceSettings />}
      </main>
    </div>
  );
};

function ProfileSettings({ text }: { text: typeof SETTINGS_TEXT["zh-CN"] | typeof SETTINGS_TEXT["en-US"] }) {
  return (
    <div className="settings-section-content">
      <h2>{text.profileTitle}</h2>
      <p className="section-desc">{text.profileDesc}</p>

      <div className="account-profile-panel">
        <div className="account-avatar">
          <User size={22} />
        </div>
        <div className="account-profile-main">
          <div className="account-profile-name">{text.localUser}</div>
          <div className="account-profile-desc">{text.localAccount}</div>
        </div>
      </div>
    </div>
  );
}
