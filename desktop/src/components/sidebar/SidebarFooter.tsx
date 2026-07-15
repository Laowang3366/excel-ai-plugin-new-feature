import React from "react";
import type { AppPage } from "../../App";
import type { SettingsSection } from "../SettingsPage";
import type { ExcelStatus } from "../../utils/sidebarHelpers";
import type { OfficeAppStatus } from "../../hooks/useOfficeConnection";
import type { getAppText } from "../../i18n";
import { LogOut, Package, RefreshCw, Settings, User } from "../common/IconMap";

interface SidebarFooterProps {
  currentPage: AppPage;
  text: ReturnType<typeof getAppText>;
  excelStatus: ExcelStatus;
  wordStatus: OfficeAppStatus;
  presentationStatus: OfficeAppStatus;
  connectFailed: boolean;
  connecting: boolean;
  pulseDot: boolean;
  settingsMenuOpen: boolean;
  onConnect: () => void;
  onToggleSettingsMenu: (event: React.MouseEvent) => void;
  onOpenSettingsSection: (section: SettingsSection) => void;
  onCloseSettingsMenu: () => void;
}

export function SidebarFooter({
  currentPage,
  text,
  excelStatus,
  wordStatus,
  presentationStatus,
  connectFailed,
  connecting,
  pulseDot,
  settingsMenuOpen,
  onConnect,
  onToggleSettingsMenu,
  onOpenSettingsSection,
  onCloseSettingsMenu,
}: SidebarFooterProps) {
  return (
    <div className="sidebar-footer">
      <div className={`sidebar-connection ${connectFailed ? "connection-failed" : ""}`}>
        <div className="sidebar-connection-apps">
          <div
            className={`connection-indicator ${excelStatus.connected ? "connected" : "disconnected"}`}
            title={excelStatus.connected ? `Excel ${excelStatus.version || ""}` : "Excel 未连接"}
          >
            <span className={`connection-dot ${pulseDot ? "pulse" : ""}`} />
            <span className="connection-text">
              {excelStatus.connected
                ? `${text.sidebar.connectedExcelPrefix}${excelStatus.version ? ` (${excelStatus.version})` : ""}`
                : text.sidebar.excelDisconnected}
            </span>
          </div>
          <div
            className={`connection-indicator ${wordStatus.connected ? "connected" : "disconnected"}`}
            title={wordStatus.connected ? `Word ${wordStatus.version || ""}` : "Word 未连接"}
          >
            <span className="connection-dot" />
            <span className="connection-text">
              {wordStatus.connected
                ? `${text.sidebar.connectedWordPrefix}${wordStatus.version ? ` (${wordStatus.version})` : ""}`
                : text.sidebar.wordDisconnected}
            </span>
          </div>
          <div
            className={`connection-indicator ${presentationStatus.connected ? "connected" : "disconnected"}`}
            title={
              presentationStatus.connected
                ? `PowerPoint ${presentationStatus.version || ""}`
                : "PowerPoint 未连接"
            }
          >
            <span className="connection-dot" />
            <span className="connection-text">
              {presentationStatus.connected
                ? `${text.sidebar.connectedPresentationPrefix}${presentationStatus.version ? ` (${presentationStatus.version})` : ""}`
                : text.sidebar.presentationDisconnected}
            </span>
          </div>
        </div>
        {excelStatus.connected ? (
          <button
            className="btn-connect btn-connect-icon"
            onClick={onConnect}
            title={text.sidebar.reconnect}
            disabled={connecting}
          >
            <RefreshCw size={10} className={connecting ? "spin" : ""} />
          </button>
        ) : (
          <button className="btn-connect" onClick={onConnect} disabled={connecting}>
            {connecting ? text.sidebar.connecting : text.sidebar.connect}
          </button>
        )}
      </div>

      <div className="sidebar-settings-menu-wrap">
        <button
          className={`sidebar-nav-btn ${currentPage === "settings" ? "active" : ""}`}
          onClick={onToggleSettingsMenu}
        >
          <Settings size={16} />
          <span>{text.sidebar.settings}</span>
        </button>

        {settingsMenuOpen && (
          <div className="sidebar-settings-menu" onClick={(e) => e.stopPropagation()}>
            <div className="sidebar-settings-account">
              <User size={16} />
              <span>{text.sidebar.localAccount}</span>
            </div>
            <div className="sidebar-settings-divider" />
            <button
              className="sidebar-settings-menu-item"
              onClick={() => onOpenSettingsSection("profile")}
            >
              <User size={16} />
              <span>{text.sidebar.profile}</span>
            </button>
            <button
              className="sidebar-settings-menu-item"
              onClick={() => onOpenSettingsSection("general")}
            >
              <Settings size={16} />
              <span>{text.sidebar.settings}</span>
            </button>
            <button
              className="sidebar-settings-menu-item"
              onClick={() => onOpenSettingsSection("opensource")}
            >
              <Package size={16} />
              <span>{text.sidebar.openSource}</span>
            </button>
            <button className="sidebar-settings-menu-item" onClick={onCloseSettingsMenu}>
              <LogOut size={16} />
              <span>{text.sidebar.logout}</span>
            </button>
          </div>
        )}
      </div>
    </div>
  );
}
