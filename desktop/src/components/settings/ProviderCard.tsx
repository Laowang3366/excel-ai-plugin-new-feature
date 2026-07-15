/**
 * ProviderCard — 供应商卡片子组件
 *
 * 展示供应商基本信息（名称、模型、API Key 遮蔽），不含内联编辑。
 * 编辑操作由 ModelSettings 管理的 EditProviderDialog 弹窗处理。
 *
 * 关联模块：
 * - modelSettingsI18n.ts — 提供 MODEL_TEXT 双语文本
 * - settingsStore — useSettingsStore（获取 language）
 */

import React from "react";
import { useSettingsStore } from "../../store/settingsStore";
import type { AiProviderConfig } from "../../electronApi";
import { Pencil, Trash2 } from "../common/IconMap";
import { MODEL_TEXT } from "./modelSettingsI18n";
import { SETTINGS_SECRET_MASK } from "../../../electron/shared/settingsSecretContract";

// ============================================================
// 类型定义
// ============================================================

export interface ProviderCardProps {
  provider: AiProviderConfig;
  isActive: boolean;
  isEditing: boolean;
  onEdit: () => void;
  onDelete: () => void;
  onSetActive: () => void;
  onUpdate: (patch: Partial<AiProviderConfig>) => void;
  onModelsLoaded: (models: string[]) => void;
}

// ============================================================
// 组件实现
// ============================================================

export const ProviderCard: React.FC<ProviderCardProps> = ({
  provider,
  isActive,
  onEdit,
  onDelete,
}) => {
  const { language } = useSettingsStore();
  const text = MODEL_TEXT[language];

  // 遮蔽 API Key 显示
  const maskedKey =
    provider.apiKey === SETTINGS_SECRET_MASK
      ? SETTINGS_SECRET_MASK
      : provider.apiKey
        ? provider.apiKey.slice(0, 6) + "••••" + provider.apiKey.slice(-4)
        : text.unset;

  return (
    <div className={`provider-card ${isActive ? "active" : ""}`}>
      {/* 卡片头部 */}
      <div className="provider-card-header" onClick={onEdit}>
        <div className="provider-card-info">
          <div className="provider-card-text">
            <span className="provider-card-name">{provider.name}</span>
            <span className="provider-card-detail">
              {provider.model || text.noModel} · {maskedKey}
            </span>
          </div>
        </div>
        <div className="provider-card-badges">
          {isActive && <span className="active-badge">{text.active}</span>}
          <button
            className="card-action-btn"
            onClick={(e) => {
              e.stopPropagation();
              onEdit();
            }}
            title="编辑"
          >
            <Pencil size={14} />
          </button>
          <button
            className="card-action-btn danger"
            onClick={(e) => {
              e.stopPropagation();
              onDelete();
            }}
            title="删除"
          >
            <Trash2 size={14} />
          </button>
        </div>
      </div>
    </div>
  );
};
