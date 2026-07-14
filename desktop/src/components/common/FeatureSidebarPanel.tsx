import type React from "react";
import type { AppLanguage } from "../../store/settingsStore";
import { getAppText } from "../../i18n";
import { INTENT_SHORTCUTS, type IntentKind } from "../../utils/sidebarHelpers";
import { Sparkles, X } from "./IconMap";

const FEATURE_SIDEBAR_TITLE_ID = "feature-sidebar-title";

interface InertAttributes {
  inert?: "";
}

interface FeatureSidebarPanelProps {
  isOpen: boolean;
  activeIntent: IntentKind;
  language: AppLanguage;
  onIntentClick: (intent: NonNullable<IntentKind>) => void;
  onClose?: () => void;
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
  const inertAttributes: InertAttributes = isOpen ? {} : { inert: "" };

  return (
    <aside
      {...inertAttributes}
      className={`feature-sidebar-panel ${isOpen ? "open" : "collapsed"}${activeIntent === "office" ? " office-automation-mode" : ""}`}
      aria-hidden={!isOpen}
      aria-labelledby={FEATURE_SIDEBAR_TITLE_ID}
    >
      <div className="feature-sidebar-content">
        <div className="feature-sidebar-header">
          <div className="feature-sidebar-title">
            <Sparkles size={18} />
            <span id={FEATURE_SIDEBAR_TITLE_ID}>{text.chat.featureSidebar.title}</span>
          </div>
          {onClose && <button className="feature-sidebar-close" type="button" title={text.chat.featureSidebar.close} aria-label={text.chat.featureSidebar.close} onClick={onClose}><X size={15} /></button>}
        </div>

        <div
          className="feature-sidebar-shortcuts"
          role="group"
          aria-label={text.chat.featureSidebar.title}
        >
          {INTENT_SHORTCUTS.map((shortcut) => (
            <button
              key={shortcut.key}
              className={`feature-sidebar-shortcut feature-${shortcut.key}${activeIntent === shortcut.key ? " active" : ""}`}
              type="button"
              aria-pressed={activeIntent === shortcut.key}
              aria-label={text.sidebar.intents[shortcut.key]}
              title={text.sidebar.intents[shortcut.key]}
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
