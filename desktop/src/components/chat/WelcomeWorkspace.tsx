import type { LucideIcon } from "lucide-react";
import type { AppLanguage } from "../../store/settingsStore";
import type { IntentKind } from "../Sidebar";
import { getAppText } from "../../i18n";
import { BarChart3, Bot, Code, FileSpreadsheet, FileText, Sparkles } from "../common/IconMap";

interface WelcomeWorkspaceProps {
  language: AppLanguage;
  onIntentClick: (intent: NonNullable<IntentKind>) => void;
}

interface QuickAction {
  key: "analyze" | "code" | "summary" | "chart";
  intent: NonNullable<IntentKind>;
  icon: LucideIcon;
}

const QUICK_ACTIONS: QuickAction[] = [
  { key: "analyze", intent: "clean", icon: FileSpreadsheet },
  { key: "code", intent: "code", icon: Code },
  { key: "summary", intent: "report", icon: FileText },
  { key: "chart", intent: "chart", icon: BarChart3 },
];

export function WelcomeWorkspace({ language, onIntentClick }: WelcomeWorkspaceProps) {
  const text = getAppText(language);

  return (
    <section className="welcome-workspace" aria-labelledby="welcome-workspace-title">
      <div className="welcome-visual" aria-hidden="true">
        <span className="welcome-spark welcome-spark-left">
          <Sparkles size={16} />
        </span>
        <span className="welcome-spark welcome-spark-right">
          <Sparkles size={13} />
        </span>
        <span className="welcome-bot-ring">
          <span className="welcome-bot-face">
            <Bot size={38} />
          </span>
        </span>
        <span className="welcome-bubble welcome-bubble-top" />
        <span className="welcome-bubble welcome-bubble-bottom" />
      </div>

      <h1 id="welcome-workspace-title">{text.chat.welcomeTitle}</h1>
      <p>{text.chat.welcomeSubtitle}</p>

      <div className="welcome-quick-actions" aria-label={text.chat.quickActionsLabel}>
        {QUICK_ACTIONS.map((action) => {
          const actionText = text.chat.quickActions[action.key];
          return (
            <button
              key={action.key}
              className={`welcome-quick-action welcome-quick-action-${action.key}`}
              type="button"
              onClick={() => onIntentClick(action.intent)}
            >
              <span className="welcome-quick-icon" aria-hidden="true">
                <action.icon size={20} />
              </span>
              <span className="welcome-quick-copy">
                <strong>{actionText.title}</strong>
                <span>{actionText.description}</span>
              </span>
            </button>
          );
        })}
      </div>
    </section>
  );
}
