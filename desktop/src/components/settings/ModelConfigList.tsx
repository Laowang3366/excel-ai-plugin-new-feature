/**
 * ModelConfigList — 聚合平台模型列表管理共享组件
 *
 * 从 ProviderCard / AddProviderDialog 提取，统一：
 * - 模型列表渲染（名称 + 上下文窗口 + 推理等级 + 删除按钮）
 * - 添加模型（Enter 键 / Plus 按钮）
 * - 删除模型时自动检测当前选中模型并通知
 * - 重名检测
 */

import React, { useRef } from "react";
import type { AiProviderConfig, ModelConfig } from "../../electronApi";
import { Trash2, Plus, Brain } from "../common/IconMap";
import { MODEL_TEXT } from "./modelSettingsI18n";
import { useSettingsStore } from "../../store/settingsStore";

/** 默认推理等级选项映射（用于 UI 标签） */
const REASONING_LABELS: Record<string, string> = {
  off: "关闭",
  low: "低",
  medium: "中",
  high: "高",
  max: "极高",
};
const REASONING_LABELS_EN: Record<string, string> = {
  off: "Off",
  low: "Low",
  medium: "Medium",
  high: "High",
  max: "Max",
};

/** 预定义的推理等级选项组 */
export const REASONING_OPTIONS_PRESETS = {
  /** off/low/medium/high/max — 完整五档 */
  FULL: ["off", "low", "medium", "high", "max"] as const,
  /** off/high/max — 仅高/极高两档（DeepSeek, 智谱） */
  HIGH_MAX: ["off", "high", "max"] as const,
  /** off/high — 开关型（Kimi, 聚合平台） */
  TOGGLE: ["off", "high"] as const,
};

export interface ModelConfigListProps {
  modelConfigs: ModelConfig[];
  currentModel: string;
  onModelConfigsChange: (
    newConfigs: ModelConfig[],
    extraPatch?: Partial<AiProviderConfig>
  ) => void;
  /** 新增模型后自动选中（仅 AddProviderDialog 使用） */
  onAddModelAutoSelect?: (name: string) => void;
}

export const ModelConfigList: React.FC<ModelConfigListProps> = ({
  modelConfigs,
  currentModel,
  onModelConfigsChange,
  onAddModelAutoSelect,
}) => {
  const { language } = useSettingsStore();
  const text = MODEL_TEXT[language];
  const addInputRef = useRef<HTMLInputElement>(null);
  const labels = language === "zh-CN" ? REASONING_LABELS : REASONING_LABELS_EN;

  /** 添加一个模型 */
  const addModel = (name: string) => {
    const trimmed = name.trim();
    if (!trimmed) return;
    if (modelConfigs.some((m) => m.name === trimmed)) return;
    const newConfigs = [...modelConfigs, { name: trimmed }];
    onModelConfigsChange(newConfigs);
    onAddModelAutoSelect?.(trimmed);
  };

  /** 删除第 idx 个模型 */
  const removeModel = (idx: number) => {
    const removed = modelConfigs[idx];
    const newConfigs = modelConfigs.filter((_, i) => i !== idx);
    const extraPatch: Partial<AiProviderConfig> = {};
    if (removed && currentModel === removed.name) {
      extraPatch.model = "";
    }
    onModelConfigsChange(newConfigs, extraPatch);
  };

  /** 更新第 idx 个模型的 contextWindowSize */
  const updateContextWindow = (idx: number, val: number | undefined) => {
    const newConfigs = [...modelConfigs];
    newConfigs[idx] = {
      ...newConfigs[idx],
      contextWindowSize: val && val > 0 ? val : undefined,
    };
    onModelConfigsChange(newConfigs);
  };

  /** 更新第 idx 个模型的 reasoningOptions */
  const updateReasoningOptions = (idx: number, options: string[] | undefined) => {
    const newConfigs = [...modelConfigs];
    newConfigs[idx] = {
      ...newConfigs[idx],
      reasoningOptions: options,
    };
    onModelConfigsChange(newConfigs);
  };

  return (
    <div className="model-config-list">
      {modelConfigs.map((mc, idx) => (
        <div key={mc.name} className="model-config-item">
          <div className="model-config-row">
            <span className="model-config-name">{mc.name}</span>
            <input
              type="number"
              className="form-input model-config-cw"
              value={mc.contextWindowSize ?? ""}
              onChange={(e) => {
                const val = e.target.value
                  ? parseInt(e.target.value, 10)
                  : undefined;
                updateContextWindow(idx, val);
              }}
              placeholder={text.contextWindowPlaceholder}
              min={1000}
              step={1000}
              title={text.contextWindowSize}
            />
            {/* 推理等级选项（可覆盖供应商级默认值） */}
            <select
              className="form-input model-config-reasoning"
              value={mc.reasoningOptions?.join(",") || ""}
              onChange={(e) => {
                const val = e.target.value;
                updateReasoningOptions(
                  idx,
                  val ? val.split(",") : undefined
                );
              }}
              title={text.reasoningMode}
            >
              <option value="">{text.defaultValue}</option>
              <option value="off,high">{labels.off} / {labels.high}</option>
              <option value="off,high,max">{labels.off} / {labels.high} / {labels.max}</option>
              <option value="off,low,medium,high,max">{labels.off} / {labels.low} / {labels.medium} / {labels.high} / {labels.max}</option>
            </select>
            <button
              className="btn-icon-danger"
              onClick={() => removeModel(idx)}
              title={text.removeModel}
            >
              <Trash2 size={13} />
            </button>
          </div>
          {mc.reasoningOptions && (
            <div className="model-config-hint">
              <Brain size={12} /> {mc.reasoningOptions.map((r) => labels[r] || r).join(" → ")}
            </div>
          )}
        </div>
      ))}
      <div className="model-config-add">
        <input
          ref={addInputRef}
          type="text"
          className="form-input model-config-add-input"
          placeholder={text.modelInputPlaceholder}
          onKeyDown={(e) => {
            if (e.key === "Enter") {
              const input = e.target as HTMLInputElement;
              addModel(input.value);
              input.value = "";
            }
          }}
        />
        <button
          className="btn-icon-hint"
          onClick={() => {
            const input = addInputRef.current;
            if (!input) return;
            addModel(input.value);
            input.value = "";
          }}
          title={text.addModel}
        >
          <Plus size={14} />
        </button>
      </div>
      <span className="form-hint">{text.modelListHint}</span>
    </div>
  );
};
