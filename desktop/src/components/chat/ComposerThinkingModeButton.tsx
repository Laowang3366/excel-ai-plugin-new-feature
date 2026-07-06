import React from "react";
import type { AiProviderConfig, ReasoningMode } from "../../electronApi";
import { getAppText } from "../../i18n";
import { PROVIDER_TEMPLATES, useSettingsStore } from "../../store/settingsStore";
import type { ProviderTemplate } from "../../store/settingsProviderTemplates";
import {
  buildReasoningOptions,
  coerceReasoningMode,
  defaultReasoningModeForOptions,
  resolveReasoningOptionValues,
} from "../../utils/reasoningSupport";
import { Brain, ChevronDown } from "../common/IconMap";

export interface ComposerThinkingModeState {
  activeProviderId: string;
  currentMode: ReasoningMode;
  isReasoningActive: boolean;
  options: Array<{ value: ReasoningMode; label: string }>;
}

export function resolveComposerThinkingModeState({
  providers,
  activeProviderId,
  templates,
  language,
}: {
  providers: Record<string, AiProviderConfig>;
  activeProviderId: string;
  templates: ProviderTemplate[];
  language: "zh-CN" | "en-US";
}): ComposerThinkingModeState | null {
  const activeProvider = providers[activeProviderId];
  if (!activeProvider) return null;

  const activeTemplate = templates.find((template) => template.provider === activeProvider.provider);
  const activeModel = activeProvider.model || "";
  const activeModelConfig = (activeProvider.modelConfigs || []).find((modelConfig) => modelConfig.name === activeModel);
  const reasoningOptionValues = resolveReasoningOptionValues(activeProvider, activeTemplate, activeModelConfig);
  const options = buildReasoningOptions(reasoningOptionValues, language);
  if (options.length === 0) return null;

  const defaultMode = defaultReasoningModeForOptions(
    reasoningOptionValues,
    activeTemplate?.defaultReasoningMode,
  );
  const currentMode = coerceReasoningMode(
    activeModelConfig?.reasoningMode || activeProvider.reasoningMode,
    reasoningOptionValues,
    defaultMode,
  );

  return {
    activeProviderId,
    currentMode,
    isReasoningActive: currentMode !== "off",
    options,
  };
}

interface ComposerThinkingModeButtonProps {
  open: boolean;
  setOpen: (open: boolean) => void;
  closePeerPopovers: () => void;
}

export const ComposerThinkingModeButton: React.FC<ComposerThinkingModeButtonProps> = ({
  open,
  setOpen,
  closePeerPopovers,
}) => {
  const { language, providers, activeProviderId, updateProvider } = useSettingsStore();
  const text = getAppText(language).chat;
  const state = resolveComposerThinkingModeState({
    providers,
    activeProviderId,
    templates: PROVIDER_TEMPLATES,
    language,
  });

  if (!state) return null;

  const currentLabel =
    state.options.find((option) => option.value === state.currentMode)?.label || state.currentMode;

  return (
    <div className="composer-popover-wrapper thinking-mode-wrapper">
      <button
        className={`composer-action-btn thinking-mode-btn ${state.isReasoningActive ? "active" : ""}`}
        onClick={(event) => {
          event.stopPropagation();
          setOpen(!open);
          closePeerPopovers();
        }}
        title={text.thinkingMode}
      >
        <Brain size={17} />
        <span className="thinking-mode-label">
          {state.isReasoningActive ? currentLabel : text.thinkingOff}
        </span>
        <ChevronDown size={13} />
      </button>
      {open && (
        <div className="composer-popover thinking-mode-popover" onClick={(event) => event.stopPropagation()}>
          {state.options.map((option) => (
            <button
              key={option.value}
              className={`popover-item ${state.currentMode === option.value ? "active" : ""}`}
              onClick={() => {
                if (state.currentMode !== option.value) {
                  updateProvider(state.activeProviderId, { reasoningMode: option.value });
                }
                setOpen(false);
              }}
            >
              <Brain size={14} /> {option.label}
            </button>
          ))}
        </div>
      )}
    </div>
  );
};
