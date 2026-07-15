import type { CSSProperties } from "react";

import type { WindowDisplayMode } from "../electronApi";
import type { getAppText } from "../i18n";
import { MAX_WINDOW_OPACITY, MIN_WINDOW_OPACITY } from "../store/settingsValues";
import { Eye, Maximize2, Minimize2, PanelLeft, Pin } from "./common/IconMap";

interface AppTitlebarProps {
  alwaysOnTop: boolean;
  collapsed?: boolean;
  displayMode: WindowDisplayMode;
  onSetWindowOpacity: (opacity: number) => void;
  onToggleAlwaysOnTop: () => void;
  onToggleCompactMode: () => void;
  onToggleSidebar?: () => void;
  showSidebarToggle: boolean;
  showWindowModeToggle?: boolean;
  text: ReturnType<typeof getAppText>;
  windowOpacity: number;
}

export function getOpacityPresentation(windowOpacity: number) {
  const valuePercent = Math.round(windowOpacity * 100);
  const minPercent = Math.round(MIN_WINDOW_OPACITY * 100);
  const maxPercent = Math.round(MAX_WINDOW_OPACITY * 100);
  const fillPercent = ((valuePercent - minPercent) / (maxPercent - minPercent)) * 100;
  return {
    fillPercent,
    maxPercent,
    minPercent,
    thumbNearValue: fillPercent >= 35 && fillPercent <= 65,
    valuePercent,
  };
}

export function AppTitlebar({
  alwaysOnTop,
  collapsed = false,
  displayMode,
  onSetWindowOpacity,
  onToggleAlwaysOnTop,
  onToggleCompactMode,
  onToggleSidebar,
  showSidebarToggle,
  showWindowModeToggle = true,
  text,
  windowOpacity,
}: AppTitlebarProps) {
  const opacity = getOpacityPresentation(windowOpacity);

  return (
    <div className="app-titlebar">
      {showSidebarToggle && (
        <button
          className={`titlebar-sidebar-toggle${collapsed ? "" : " active"}`}
          onClick={onToggleSidebar}
          title={collapsed ? text.app.expandSidebar : text.app.collapseSidebar}
          aria-pressed={!collapsed}
        >
          <PanelLeft size={17} />
        </button>
      )}
      {showWindowModeToggle && (
        <button
          className={`titlebar-window-mode-toggle ${displayMode === "compact" ? "active" : ""}`}
          onClick={onToggleCompactMode}
          title={displayMode === "normal" ? text.app.compactWindow : text.app.restoreWindow}
          aria-pressed={displayMode === "compact"}
        >
          {displayMode === "normal" ? <Minimize2 size={15} /> : <Maximize2 size={15} />}
        </button>
      )}
      <label
        className="titlebar-opacity-control"
        title={`${text.app.windowOpacity}: ${opacity.valuePercent}%`}
      >
        <Eye size={14} aria-hidden="true" />
        <span className="titlebar-opacity-track">
          <input
            type="range"
            className="titlebar-opacity-slider"
            min={opacity.minPercent}
            max={opacity.maxPercent}
            step={1}
            value={opacity.valuePercent}
            aria-label={text.app.windowOpacity}
            style={{ "--slider-fill": `${opacity.fillPercent}%` } as CSSProperties}
            onChange={(event) => onSetWindowOpacity(Number(event.target.value) / 100)}
          />
          <span
            className={`titlebar-opacity-value${
              opacity.thumbNearValue ? " avoid-thumb" : opacity.fillPercent > 65 ? " over-fill" : ""
            }`}
          >
            {opacity.valuePercent}%
          </span>
        </span>
      </label>
      <button
        className={`titlebar-pin-toggle ${alwaysOnTop ? "active" : ""}`}
        onClick={onToggleAlwaysOnTop}
        title={alwaysOnTop ? text.app.pinOff : text.app.pinOn}
        aria-pressed={alwaysOnTop}
      >
        <Pin size={15} />
      </button>
    </div>
  );
}
