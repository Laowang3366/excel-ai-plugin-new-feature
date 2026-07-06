import React from "react";
import type { AppPage } from "../../App";
import type { ExcelStatus } from "../../utils/sidebarHelpers";
import type { OfficeAppStatus } from "../../hooks/useOfficeConnection";
import type { getAppText } from "../../i18n";
import {
  FolderOpen,
  Plus,
  RefreshCw,
  Search,
  Settings,
} from "../common/IconMap";

interface SidebarCollapsedProps {
  currentPage: AppPage;
  text: ReturnType<typeof getAppText>;
  creatingNewThread: boolean;
  excelStatus: ExcelStatus;
  wordStatus: OfficeAppStatus;
  presentationStatus: OfficeAppStatus;
  onCreateNewThread: () => void;
  onToggleSearch: () => void;
  onAddFolder: () => void;
  onOpenSettings: () => void;
}

export function SidebarCollapsed({
  currentPage,
  text,
  creatingNewThread,
  excelStatus,
  wordStatus,
  presentationStatus,
  onCreateNewThread,
  onToggleSearch,
  onAddFolder,
  onOpenSettings,
}: SidebarCollapsedProps) {
  return (
    <aside className="sidebar sidebar-collapsed">
      <button
        className={`sidebar-icon-btn${creatingNewThread ? " creating" : ""}`}
        onClick={onCreateNewThread}
        disabled={creatingNewThread}
        title={text.sidebar.newThread}
      >
        {creatingNewThread ? <RefreshCw size={18} className="spin" /> : <Plus size={18} />}
      </button>
      <div className="sidebar-spacer" />
      <button className="sidebar-icon-btn" onClick={onToggleSearch} title={text.sidebar.search}>
        <Search size={16} />
      </button>
      <button className="sidebar-icon-btn" onClick={onAddFolder} title={text.sidebar.addFolder}>
        <FolderOpen size={16} />
      </button>
      <div
        className={`sidebar-status-dot ${excelStatus.connected ? "connected" : "disconnected"}`}
        title={`Excel: ${excelStatus.connected ? "已连接" : "未连接"} | Word: ${wordStatus.connected ? "已连接" : "未连接"} | PPT: ${presentationStatus.connected ? "已连接" : "未连接"}`}
      />
      <button
        className={`sidebar-icon-btn ${currentPage === "settings" ? "active" : ""}`}
        onClick={onOpenSettings}
        title={text.sidebar.settings}
      >
        <Settings size={16} />
      </button>
    </aside>
  );
}
