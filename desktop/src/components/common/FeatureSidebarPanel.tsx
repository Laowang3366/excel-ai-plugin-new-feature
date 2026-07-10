import type React from "react";
import type { AppLanguage } from "../../store/settingsStore";
import { getAppText } from "../../i18n";
import { INTENT_SHORTCUTS, type IntentKind } from "../../utils/sidebarHelpers";
import { Sparkles, X } from "./IconMap";

interface FeatureSidebarPanelProps {
  isOpen: boolean;
  activeIntent: IntentKind;
  language: AppLanguage;
  onIntentClick: (intent: NonNullable<IntentKind>) => void;
  onClose: () => void;
  children?: React.ReactNode;
}

export function FeatureSidebarPanel({
  isOpen,
  activeIntent,
  language,
  onIntentClick,
  onClose,
  children,
}: FeatureSidebarPanelProps) {
  const text = getAppText(language);

  return (
    <aside
      className={`feature-sidebar-panel ${isOpen ? "open" : "collapsed"}`}
      aria-hidden={!isOpen}
    >
      <div className="feature-sidebar-content">
        <div className="feature-sidebar-header">
          <div className="feature-sidebar-title">
            <Sparkles size={18} />
            <span>{text.chat.featureSidebar.title}</span>
          </div>
          <button
            className="feature-sidebar-close"
            type="button"
            onClick={onClose}
            title={text.chat.featureSidebar.close}
            aria-label={text.chat.featureSidebar.close}
          >
            <X size={18} />
          </button>
        </div>

        <div
          className="feature-sidebar-shortcuts"
          role="listbox"
          aria-label={text.chat.featureSidebar.title}
        >
          {INTENT_SHORTCUTS.map((shortcut) => (
            <button
              key={shortcut.key}
              className={`feature-sidebar-shortcut ${activeIntent === shortcut.key ? "active" : ""}`}
              type="button"
              role="option"
              aria-selected={activeIntent === shortcut.key}
              onClick={() => onIntentClick(shortcut.key)}
            >
              <shortcut.icon size={16} />
              <span>{text.sidebar.intents[shortcut.key]}</span>
            </button>
          ))}
        </div>

        <div className="feature-sidebar-form">
          {activeIntent ? (
            children
          ) : (
            <div className="feature-sidebar-empty">{text.chat.featureSidebar.empty}</div>
          )}
        </div>
      </div>
    </aside>
  );
}
