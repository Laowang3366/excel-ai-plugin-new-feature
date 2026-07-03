/**
 * ModelSettings — 供应商管理主组件（编排器）
 *
 * 职责：组合 ProviderCard、AddProviderDialog、EditProviderDialog、DeleteConfirmDialog，
 * 管理编辑状态（editingId、showAddDialog、deleteConfirmId），从 settingsStore 读写数据。
 *
 * 子组件拆分自本文件（原 1,186 行 → 本文件 ~120 行 + 5 个子模块）：
 * - modelSettingsI18n.ts — 双语文本常量 + 模板分组（被所有子组件引用）
 * - ProviderCard.tsx — 供应商卡片（展示/选择/删除）
 * - AddProviderDialog.tsx — 添加供应商弹窗（模板选择/表单）
 * - EditProviderDialog.tsx — 编辑供应商弹窗（修改配置/保存）
 * - DeleteConfirmDialog.tsx — 删除确认弹窗
 *
 * 外部依赖：
 * - settingsStore — providers、activeProviderId、CRUD 方法、language
 * - IconMap — Plus、Bot 图标
 */

import React, { useState } from "react";
import { useSettingsStore } from "../../store/settingsStore";
import { Plus, Bot } from "../common/IconMap";
import { MODEL_TEXT } from "./modelSettingsI18n";
import { ProviderCard } from "./ProviderCard";
import { AddProviderDialog } from "./AddProviderDialog";
import { EditProviderDialog } from "./EditProviderDialog";
import { DeleteConfirmDialog } from "./DeleteConfirmDialog";

// ============================================================
// 主组件
// ============================================================

export const ModelSettings: React.FC = () => {
  const {
    providers,
    activeProviderId,
    addProvider,
    updateProvider,
    removeProvider,
    setProviderModels,
    setActiveProvider,
    generateId,
    language,
  } = useSettingsStore();
  const text = MODEL_TEXT[language];

  const [editingId, setEditingId] = useState<string | null>(null);
  const [showAddDialog, setShowAddDialog] = useState(false);
  const [deleteConfirmId, setDeleteConfirmId] = useState<string | null>(null);

  const providerList = Object.values(providers);
  const isEmpty = providerList.length === 0;

  return (
    <div className="settings-section-content">
      <div className="section-header-row">
        <div>
          <h2>{text.title}</h2>
          <p className="section-desc">{text.desc}</p>
        </div>
        <button className="btn-add-provider" onClick={() => setShowAddDialog(true)}>
          <Plus size={14} /> {text.addProvider}
        </button>
      </div>

      {/* 空状态 */}
      {isEmpty && (
        <div className="provider-empty">
          <Bot size={36} style={{ opacity: 0.3 }} />
          <p style={{ color: "var(--text-faint)", fontSize: 13 }}>
            {text.empty}
          </p>
          <button className="btn-add-provider" onClick={() => setShowAddDialog(true)}>
            <Plus size={14} /> {text.addProvider}
          </button>
        </div>
      )}

      {/* 供应商卡片列表 */}
      {providerList.map((provider) => (
        <ProviderCard
          key={provider.id}
          provider={provider}
          isActive={provider.id === activeProviderId}
          isEditing={editingId === provider.id}
          onEdit={() => setEditingId(editingId === provider.id ? null : provider.id)}
          onDelete={() => setDeleteConfirmId(provider.id)}
          onSetActive={() => setActiveProvider(provider.id)}
          onUpdate={(patch) => updateProvider(provider.id, patch)}
          onModelsLoaded={(models) => setProviderModels(provider.id, models)}
        />
      ))}

      {/* 添加供应商对话框 */}
      {showAddDialog && (
        <AddProviderDialog
          onAdd={(config) => {
            addProvider(config);
            setShowAddDialog(false);
          }}
          onClose={() => setShowAddDialog(false)}
          generateId={generateId}
        />
      )}

      {/* 编辑供应商对话框 */}
      {editingId && providers[editingId] && (
        <EditProviderDialog
          provider={providers[editingId]}
          onSave={(patch) => {
            updateProvider(editingId, patch);
            setEditingId(null);
          }}
          onClose={() => setEditingId(null)}
        />
      )}

      {/* 删除确认弹窗 */}
      {deleteConfirmId && (
        <DeleteConfirmDialog
          name={providers[deleteConfirmId]?.name || ""}
          onConfirm={() => {
            removeProvider(deleteConfirmId);
            setDeleteConfirmId(null);
            if (editingId === deleteConfirmId) setEditingId(null);
          }}
          onCancel={() => setDeleteConfirmId(null)}
        />
      )}
    </div>
  );
};
