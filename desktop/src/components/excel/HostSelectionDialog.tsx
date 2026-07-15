/**
 * HostSelectionDialog — 宿主选择弹窗
 *
 * 当 Office Excel 和 WPS 表格同时运行时弹出，
 * 让用户选择要连接的目标程序。
 *
 * 复用 dialog.css 定义的 .dialog-overlay / .dialog / .dialog-header / .dialog-body / .dialog-actions 样式。
 */

import React from "react";
import { useSettingsStore } from "../../store/settingsStore";
import { getAppText } from "../../i18n";
import { FileSpreadsheet } from "../common/IconMap";

// ============================================================
// 类型定义
// ============================================================

export interface HostSelectionDialogProps {
  /** 可用宿主列表（如 ["excel", "wps"]） */
  availableHosts: string[];
  /** 用户选择后的回调 */
  onSelect: (host: "excel" | "wps") => void;
  /** 关闭弹窗（不选择，断开检测） */
  onDismiss: () => void;
}

// ============================================================
// 组件实现
// ============================================================

export const HostSelectionDialog: React.FC<HostSelectionDialogProps> = ({
  availableHosts,
  onSelect,
  onDismiss,
}) => {
  const { language } = useSettingsStore();
  const text = getAppText(language);

  const hasExcel = availableHosts.includes("excel");
  const hasWps = availableHosts.includes("wps");

  return (
    <div className="dialog-overlay" onClick={onDismiss}>
      <div className="dialog host-selection-dialog" onClick={(e) => e.stopPropagation()}>
        <div className="dialog-header">
          <h3>
            <FileSpreadsheet size={16} style={{ verticalAlign: "middle", marginRight: 6 }} />
            {text.sidebar.selectHostTitle}
          </h3>
        </div>
        <div className="dialog-body">
          <p>{text.sidebar.selectHostDescription}</p>
          <div className="host-selection-options">
            {hasExcel && (
              <button className="host-selection-btn" onClick={() => onSelect("excel")}>
                <span className="host-selection-icon">
                  <FileSpreadsheet size={24} />
                </span>
                <span className="host-selection-label">
                  <strong>{text.sidebar.selectHostExcel}</strong>
                  <span className="host-selection-subtitle">ProgID: Excel.Application</span>
                </span>
              </button>
            )}
            {hasWps && (
              <button className="host-selection-btn" onClick={() => onSelect("wps")}>
                <span className="host-selection-icon">
                  <FileSpreadsheet size={24} />
                </span>
                <span className="host-selection-label">
                  <strong>{text.sidebar.selectHostWps}</strong>
                  <span className="host-selection-subtitle">ProgID: Ket.Application</span>
                </span>
              </button>
            )}
          </div>
        </div>
        <div className="dialog-actions">
          <button className="btn-secondary" onClick={onDismiss}>
            {text.sidebar.cancel}
          </button>
        </div>
      </div>
    </div>
  );
};
