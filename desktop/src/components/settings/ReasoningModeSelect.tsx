/**
 * ReasoningModeSelect — 思考等级选择共享组件
 *
 * 从 ProviderCard / AddProviderDialog 提取，统一：
 * - reasoningOptions.length > 2 → 下拉选择
 * - reasoningOptions.length ≤ 2 → 复选框开关
 * - 标签、提示文本
 */

import React from "react";
import type { ReasoningMode } from "../../electronApi";
import { MODEL_TEXT } from "./modelSettingsI18n";
import { useSettingsStore } from "../../store/settingsStore";

interface ReasoningOption {
  value: string;
  label: string;
}

export interface ReasoningModeSelectProps {
  reasoningOptions: ReasoningOption[];
  value: ReasoningMode;
  defaultMode: ReasoningMode;
  onChange: (mode: ReasoningMode) => void;
}

export const ReasoningModeSelect: React.FC<ReasoningModeSelectProps> = ({
  reasoningOptions,
  value,
  defaultMode,
  onChange,
}) => {
  const { language } = useSettingsStore();
  const text = MODEL_TEXT[language];

  if (reasoningOptions.length > 2) {
    // 等级型：下拉选择
    return (
      <div className="form-group">
        <label>{text.reasoningMode}</label>
        <select
          className="form-input reasoning-select"
          value={value}
          onChange={(e) => onChange(e.target.value as ReasoningMode)}
        >
          {reasoningOptions.map((opt) => (
            <option key={opt.value} value={opt.value}>{opt.label}</option>
          ))}
        </select>
        <span className="form-hint">{text.reasoningModeHint}</span>
      </div>
    );
  }

  // 开关型：复选框
  return (
    <div className="form-group checkbox-group">
      <label>
        <input
          type="checkbox"
          checked={value !== "off"}
          onChange={(e) => onChange(e.target.checked ? defaultMode : "off")}
        />
        {text.enableThinking}
      </label>
      <span className="form-hint">{text.thinkingHint}</span>
    </div>
  );
};
