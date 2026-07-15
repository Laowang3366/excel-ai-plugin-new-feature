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
import { Trash2, Plus } from "../common/IconMap";
import { MODEL_TEXT } from "./modelSettingsI18n";
import { useSettingsStore } from "../../store/settingsStore";

export interface ModelConfigListProps {
  modelConfigs: ModelConfig[];
  currentModel: string;
  onModelConfigsChange: (newConfigs: ModelConfig[], extraPatch?: Partial<AiProviderConfig>) => void;
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
                const val = e.target.value ? parseInt(e.target.value, 10) : undefined;
                updateContextWindow(idx, val);
              }}
              placeholder={text.contextWindowPlaceholder}
              min={1000}
              step={1000}
              title={text.contextWindowSize}
            />
            <button
              className="btn-icon-danger"
              onClick={() => removeModel(idx)}
              title={text.removeModel}
            >
              <Trash2 size={13} />
            </button>
          </div>
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
