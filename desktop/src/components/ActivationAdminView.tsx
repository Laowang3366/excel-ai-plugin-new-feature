/**
 * 激活信息管理视图 — 设置页中的"激活信息"板块
 *
 * 显示当前激活状态、卡密信息，支持反激活。
 */

import React, { useEffect, useState } from "react";
import { useActivationStore } from "../store/activationStore";
import { useSettingsStore } from "../store/settingsStore";

export const ActivationAdminView: React.FC = () => {
  const {
    activated,
    info,
    isLoading,
    loadStatus,
    deactivate,
  } = useActivationStore();
  const { language } = useSettingsStore();

  const [machineInfo, setMachineInfo] = useState<{ machineId: string; machineName: string } | null>(null);
  const [confirmDeactivate, setConfirmDeactivate] = useState(false);

  useEffect(() => {
    loadStatus();
    // 获取设备信息
    import("../services/ipcApi").then(({ ipcApi }) => {
      ipcApi.activation.getMachineInfo().then(setMachineInfo).catch(() => {});
    });
  }, []);

  const handleDeactivate = async () => {
    await deactivate();
    setConfirmDeactivate(false);
  };

  const formatDate = (dateStr: string | null) => {
    if (!dateStr) return "—";
    try {
      const d = new Date(dateStr);
      return d.toLocaleString("zh-CN");
    } catch {
      return dateStr;
    }
  };

  const isExpired = info?.expiresAt ? new Date(info.expiresAt).getTime() < Date.now() : false;

  const text = {
    "zh-CN": {
      title: "激活信息",
      desc: "查看和管理应用激活状态",
      activated: "已激活",
      notActivated: "未激活",
      deactivate: "反激活",
      confirmDeactivate: "确认反激活？",
      cancel: "取消",
      confirm: "确认",
      key: "激活卡密",
      machineId: "设备标识",
      machineName: "设备名称",
      activatedAt: "激活时间",
      lastVerified: "最后验证",
      expiresAt: "过期时间",
      expired: "已过期",
      permanent: "永久有效",
      loading: "加载中...",
    },
    "en-US": {
      title: "Activation",
      desc: "View and manage application activation status",
      activated: "Activated",
      notActivated: "Not Activated",
      deactivate: "Deactivate",
      confirmDeactivate: "Confirm deactivation?",
      cancel: "Cancel",
      confirm: "Confirm",
      key: "License Key",
      machineId: "Machine ID",
      machineName: "Machine Name",
      activatedAt: "Activated At",
      lastVerified: "Last Verified",
      expiresAt: "Expires At",
      expired: "Expired",
      permanent: "Permanent",
      loading: "Loading...",
    },
  }[language === "en-US" ? "en-US" : "zh-CN"];

  return (
    <div className="settings-section-content">
      <h2>{text.title}</h2>
      <p className="section-desc">{text.desc}</p>

      {isLoading ? (
        <div className="loading-spinner" style={{ padding: 32, textAlign: "center" }}>
          <div className="spinner" />
          <p style={{ marginTop: 8, color: "#94a3b8", fontSize: 13 }}>{text.loading}</p>
        </div>
      ) : (
        <>
          {/* 激活状态卡片 */}
          <div className={`activation-status-card ${activated ? "activated" : "not-activated"}`}>
            <div className="activation-status-icon">
              {activated ? "✅" : "❌"}
            </div>
            <div className="activation-status-info">
              <div className="activation-status-label">
                {activated ? text.activated : text.notActivated}
              </div>
              {activated && info && isExpired && (
                <div className="activation-status-warning">{text.expired}</div>
              )}
            </div>
          </div>

          {activated && info && (
            <div className="activation-detail-card">
              <div className="activation-detail-row">
                <span className="detail-label">{text.key}</span>
                <span className="detail-value mono">{info.key}</span>
              </div>
              <div className="activation-detail-row">
                <span className="detail-label">{text.machineName}</span>
                <span className="detail-value">{machineInfo?.machineName || "—"}</span>
              </div>
              <div className="activation-detail-row">
                <span className="detail-label">{text.machineId}</span>
                <span className="detail-value mono small">{info.machineId}</span>
              </div>
              <div className="activation-detail-row">
                <span className="detail-label">{text.activatedAt}</span>
                <span className="detail-value">{formatDate(info.activatedAt)}</span>
              </div>
              <div className="activation-detail-row">
                <span className="detail-label">{text.lastVerified}</span>
                <span className="detail-value">{formatDate(info.lastVerifiedAt)}</span>
              </div>
              <div className="activation-detail-row">
                <span className="detail-label">{text.expiresAt}</span>
                <span className="detail-value" style={{ color: isExpired ? "#ef4444" : undefined }}>
                  {info.expiresAt ? formatDate(info.expiresAt) : text.permanent}
                  {isExpired && ` (${text.expired})`}
                </span>
              </div>
            </div>
          )}

          {/* 反激活 */}
          {activated && (
            <div style={{ marginTop: 16 }}>
              {confirmDeactivate ? (
                <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
                  <span style={{ fontSize: 13, color: "#ef4444" }}>{text.confirmDeactivate}</span>
                  <button
                    className="btn btn-sm btn-danger"
                    onClick={handleDeactivate}
                  >
                    {text.confirm}
                  </button>
                  <button
                    className="btn btn-sm btn-outline"
                    onClick={() => setConfirmDeactivate(false)}
                  >
                    {text.cancel}
                  </button>
                </div>
              ) : (
                <button
                  className="btn btn-sm btn-danger"
                  onClick={() => setConfirmDeactivate(true)}
                >
                  {text.deactivate}
                </button>
              )}
            </div>
          )}
        </>
      )}

      <style>{`
        .activation-status-card {
          display: flex;
          align-items: center;
          gap: 12px;
          padding: 16px;
          border-radius: 8px;
          margin-bottom: 16px;
        }
        .activation-status-card.activated {
          background: rgba(34, 197, 94, 0.1);
          border: 1px solid rgba(34, 197, 94, 0.2);
        }
        .activation-status-card.not-activated {
          background: rgba(239, 68, 68, 0.1);
          border: 1px solid rgba(239, 68, 68, 0.2);
        }
        .activation-status-icon {
          font-size: 24px;
        }
        .activation-status-label {
          font-size: 16px;
          font-weight: 600;
          color: var(--gray-800);
        }
        .activation-status-warning {
          font-size: 12px;
          color: #ef4444;
          margin-top: 2px;
        }
        .activation-detail-card {
          background: white;
          border: 1px solid var(--gray-200);
          border-radius: 8px;
          padding: 16px;
        }
        .activation-detail-row {
          display: flex;
          justify-content: space-between;
          align-items: center;
          padding: 8px 0;
          border-bottom: 1px solid var(--gray-100);
        }
        .activation-detail-row:last-child {
          border-bottom: none;
        }
        .detail-label {
          font-size: 13px;
          color: var(--gray-500);
        }
        .detail-value {
          font-size: 13px;
          color: var(--gray-800);
          font-weight: 500;
        }
        .detail-value.mono {
          font-family: "SF Mono", "Fira Code", Consolas, monospace;
        }
        .detail-value.small {
          font-size: 12px;
        }
      `}</style>
    </div>
  );
};
