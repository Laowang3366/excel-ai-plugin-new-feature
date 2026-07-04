/**
 * 模型快速切换 — 嵌入输入框工具栏的下拉选择器
 *
 * 显示所有已配置供应商下的可用模型列表（跨厂商），
 * 而非厂商本身。选择某个模型时自动切换到对应供应商
 * 并设置该模型为当前使用。
 *
 * 布局：
 * - 按供应商分组，每组显示图标+厂商名
 * - 每个模型显示名称，活跃模型带 ✓
 * - 底部"管理模型配置"入口
 */

import React, { useState, useRef, useEffect, useLayoutEffect } from "react";
import { useSettingsStore, PROVIDER_TEMPLATES } from "../../store/settingsStore";
import { Check, ChevronDown, Settings } from "../common/IconMap";
import { getAppText } from "../../i18n";
import type { SettingsSection } from "../SettingsPage";
import {
  coerceReasoningMode,
  defaultReasoningModeForOptions,
  resolveReasoningOptionValues,
} from "../../utils/reasoningSupport";

interface ModelQuickSwitchProps {
  onOpenSettings?: (section?: SettingsSection) => void;
}

/** 分组模型条目 */
interface ModelEntry {
  modelId: string;
  providerId: string;
  providerName: string;
}

export const ModelQuickSwitch: React.FC<ModelQuickSwitchProps> = ({ onOpenSettings }) => {
  const { providers, activeProviderId, setActiveProvider, updateProvider, language } = useSettingsStore();
  const [open, setOpen] = useState(false);
  const [dropdownStyle, setDropdownStyle] = useState<React.CSSProperties>({});
  const dropdownRef = useRef<HTMLDivElement>(null);
  const buttonRef = useRef<HTMLButtonElement>(null);

  const activeProvider = providers[activeProviderId];
  const text = getAppText(language);

  // 点击外部关闭
  useEffect(() => {
    if (!open) return;
    const handleClickOutside = (e: MouseEvent) => {
      if (dropdownRef.current && !dropdownRef.current.contains(e.target as Node)) {
        setOpen(false);
      }
    };
    document.addEventListener("mousedown", handleClickOutside);
    return () => document.removeEventListener("mousedown", handleClickOutside);
  }, [open]);

  useLayoutEffect(() => {
    if (!open) return;

    const updateDropdownPosition = () => {
      const buttonRect = buttonRef.current?.getBoundingClientRect();
      if (!buttonRect) return;

      const dropdownWidth = 260;
      const sidePadding = 12;
      const horizontalOffset = 8;
      const left = Math.min(
        Math.max(buttonRect.left - horizontalOffset, sidePadding),
        window.innerWidth - dropdownWidth - sidePadding
      );
      const bottom = window.innerHeight - buttonRect.top + 8;
      const availableHeight = Math.max(180, buttonRect.top - 18);
      const dropdownMaxHeight = Math.max(130, availableHeight - 50);

      setDropdownStyle({
        position: "fixed",
        left,
        bottom,
        width: dropdownWidth,
        maxHeight: Math.min(310, dropdownMaxHeight),
      });
    };

    updateDropdownPosition();
    window.addEventListener("resize", updateDropdownPosition);
    return () => window.removeEventListener("resize", updateDropdownPosition);
  }, [open]);

  /**
   * 选择一个模型
   *
   * 自动切换到模型所属的供应商，并设置该模型为当前使用。
   * 如果供应商没有 reasoningMode，自动应用模板的默认值。
   */
  const handleSelectModel = (providerId: string, modelId: string) => {
    // 先切换供应商
    setActiveProvider(providerId);
    // 更新供应商的模型选择
    const provider = providers[providerId];
    const patch: Record<string, unknown> = { model: modelId };
    if (provider) {
      const template = PROVIDER_TEMPLATES.find(t => t.provider === provider.provider);
      const modelConfig = provider.modelConfigs?.find((item) => item.name === modelId);
      const optionValues = resolveReasoningOptionValues(
        { ...provider, model: modelId },
        template,
        modelConfig,
      );
      const defaultMode = defaultReasoningModeForOptions(optionValues, template?.defaultReasoningMode);
      patch.reasoningMode = coerceReasoningMode(
        modelConfig?.reasoningMode || provider.reasoningMode,
        optionValues,
        defaultMode,
      );
    }
    updateProvider(providerId, patch);
    setOpen(false);
  };

  // 收集所有已配置供应商下的可用模型，按供应商分组
  const configuredProviders = Object.values(providers).filter((p) => p.apiKey);

  // 构建分组模型列表
  const groupedModels: { providerId: string; providerName: string; models: ModelEntry[] }[] = [];

  for (const provider of configuredProviders) {
    const template = PROVIDER_TEMPLATES.find((t) => t.provider === provider.provider);
    const isAggregation = template?.category === "aggregation";

    // 聚合平台：优先用 modelConfigs，否则用当前单个 model
    // 直连供应商：优先用已获取的 models 列表，否则用预设 presetModels
    let modelNames: string[];
    if (isAggregation && provider.modelConfigs && provider.modelConfigs.length > 0) {
      modelNames = provider.modelConfigs.map((m) => m.name);
    } else if (isAggregation) {
      // 无 modelConfigs 的聚合平台，显示当前配置的单个 model
      modelNames = provider.model ? [provider.model] : [];
    } else {
      const presetModels = template?.presetModels || [];
      const storedModels = provider.models || [];
      modelNames = storedModels.length > 0 ? storedModels : presetModels;
    }

    if (modelNames.length === 0) {
      // 如果没有模型列表，至少显示当前配置的模型
      if (provider.model) {
        groupedModels.push({
          providerId: provider.id,
          providerName: provider.name,
          models: [{
            modelId: provider.model,
            providerId: provider.id,
            providerName: provider.name,
          }],
        });
      }
    } else {
      const entries: ModelEntry[] = modelNames.map((m) => ({
        modelId: m,
        providerId: provider.id,
        providerName: provider.name,
      }));
      groupedModels.push({
        providerId: provider.id,
        providerName: provider.name,
        models: entries,
      });
    }
  }

  // 当前活跃的模型名称
  const currentModelName = activeProvider?.model || text.modelSwitch.selectModel;

  return (
    <div className="model-quick-switch" ref={dropdownRef}>
      <button
        ref={buttonRef}
        className={`model-quick-switch-btn ${open ? "open" : ""}`}
        onClick={() => setOpen(!open)}
        title={text.modelSwitch.switchModel}
      >
        <span className="model-qs-name">
          {currentModelName}
        </span>
        <span className="model-qs-arrow"><ChevronDown size={12} /></span>
      </button>

      {open && (
        <div className="model-qs-dropdown" style={dropdownStyle}>
          {groupedModels.length === 0 ? (
            <div className="model-qs-empty">
              <p>{text.modelSwitch.noProviders}</p>
              <button
                className="model-qs-manage-btn"
                onClick={() => {
                  setOpen(false);
                  onOpenSettings?.("model");
                }}
              >
                <Settings size={13} /> {text.modelSwitch.configure}
              </button>
            </div>
          ) : (
            <>
              <div className="model-qs-list">
                {groupedModels.map((group) => {
                  const isGroupActive = activeProviderId === group.providerId;

                  return (
                    <div key={group.providerId} className="model-qs-group">
                      {/* 供应商分组标题 */}
                      <div className="model-qs-group-header">
                        <span className="model-qs-group-name">{group.providerName}</span>
                      </div>
                      {/* 该供应商下的模型列表 */}
                      {group.models.map((entry) => {
                        const isActive = isGroupActive && activeProvider?.model === entry.modelId;

                        return (
                          <button
                            key={`${entry.providerId}-${entry.modelId}`}
                            className={`model-qs-option ${isActive ? "active" : ""}`}
                            onClick={() => handleSelectModel(entry.providerId, entry.modelId)}
                          >
                            <span className="model-qs-option-name">{entry.modelId}</span>
                            {isActive && <span className="model-qs-option-check"><Check size={13} /></span>}
                          </button>
                        );
                      })}
                    </div>
                  );
                })}
              </div>
              <div className="model-qs-divider" />
              <button
                className="model-qs-option manage"
                onClick={() => {
                  setOpen(false);
                  onOpenSettings?.("model");
                }}
              >
                <Settings size={14} />
                <span>{text.modelSwitch.manage}</span>
              </button>
            </>
          )}
        </div>
      )}
    </div>
  );
};
