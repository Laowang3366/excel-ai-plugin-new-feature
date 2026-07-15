/**
 * DeleteConfirmDialog — 删除确认弹窗子组件
 *
 * 从 ModelSettings.tsx 提取，负责：
 * - 显示删除确认对话框
 * - 确认/取消操作
 *
 * 关联模块：
 * - modelSettingsI18n.ts — 提供 MODEL_TEXT 双语文本
 * - settingsStore — useSettingsStore（获取 language）
 * - IconMap — X 图标
 */

import React from "react";
import { useSettingsStore } from "../../store/settingsStore";
import { X } from "../common/IconMap";
import { MODEL_TEXT } from "./modelSettingsI18n";

// ============================================================
// 类型定义
// ============================================================

export interface DeleteConfirmDialogProps {
  name: string;
  onConfirm: () => void;
  onCancel: () => void;
}

// ============================================================
// 组件实现
// ============================================================

export const DeleteConfirmDialog: React.FC<DeleteConfirmDialogProps> = ({
  name,
  onConfirm,
  onCancel,
}) => {
  const { language } = useSettingsStore();
  const text = MODEL_TEXT[language];

  return (
    <div className="dialog-overlay" onClick={onCancel}>
      <div className="dialog delete-confirm-dialog" onClick={(e) => e.stopPropagation()}>
        <div className="dialog-header">
          <h3>{text.deleteTitle}</h3>
          <button className="dialog-close" onClick={onCancel}>
            <X size={16} />
          </button>
        </div>
        <div className="dialog-body">
          <p>{text.deleteMessage(name)}</p>
        </div>
        <div className="dialog-actions">
          <button className="btn-secondary" onClick={onCancel}>
            {text.cancel}
          </button>
          <button className="btn-danger" onClick={onConfirm}>
            {text.delete}
          </button>
        </div>
      </div>
    </div>
  );
};
