import React, { useEffect, useState } from "react";
import { ipcApi } from "../services/ipcApi";
import { useActivationStore } from "../store/activationStore";

interface BoundDevice {
  machine_id: string;
  machine_name: string | null;
  activated_at: string;
  last_heartbeat: string | null;
  last_heartbeat_ago: string;
  is_current: boolean;
  is_online: boolean;
  online_duration_formatted: string;
}

export const DeviceManagementView: React.FC = () => {
  const { activated, loadStatus } = useActivationStore();
  const [devices, setDevices] = useState<BoundDevice[]>([]);
  const [maxMachines, setMaxMachines] = useState(0);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");
  const [unbindingId, setUnbindingId] = useState<string | null>(null);

  const loadDevices = async () => {
    if (!activated) return;
    setLoading(true);
    setError("");
    try {
      const result = await ipcApi.activation.listDevices();
      if (!result.success) {
        setError(result.error || "设备列表加载失败");
        return;
      }
      setDevices(result.data?.machines || []);
      setMaxMachines(result.data?.max_machines || 0);
    } catch (err: any) {
      setError(err.message || "设备列表加载失败");
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    loadDevices();
  }, [activated]);

  const handleUnbind = async (device: BoundDevice) => {
    const message = device.is_current
      ? "解绑当前设备后，本机需要重新激活才能继续使用。确认解绑？"
      : `确认解绑设备「${device.machine_name || device.machine_id}」？`;
    if (!window.confirm(message)) return;

    setUnbindingId(device.machine_id);
    setError("");
    try {
      const result = await ipcApi.activation.unbindDevice(device.machine_id);
      if (!result.success) {
        setError(result.error || "设备解绑失败");
        return;
      }
      if (result.currentDeviceUnbound) {
        await loadStatus();
        return;
      }
      await loadDevices();
    } catch (err: any) {
      setError(err.message || "设备解绑失败");
    } finally {
      setUnbindingId(null);
    }
  };

  if (!activated) {
    return null;
  }

  return (
    <div className="device-management-card">
      <div className="device-management-header">
        <div>
          <h3>设备管理</h3>
          <p>已绑定 {devices.length} / {maxMachines || "—"} 台设备</p>
        </div>
        <button className="btn btn-sm btn-outline" onClick={loadDevices} disabled={loading}>
          {loading ? "刷新中..." : "刷新"}
        </button>
      </div>

      {error && <div className="device-management-error">{error}</div>}

      <div className="device-list">
        {devices.length === 0 && !loading ? (
          <div className="device-empty">暂无绑定设备</div>
        ) : devices.map((device) => (
          <div key={device.machine_id} className="device-row">
            <div className="device-main">
              <div className="device-title">
                <span>{device.machine_name || "未命名设备"}</span>
                {device.is_current && <span className="device-tag current">当前设备</span>}
                <span className={`device-tag ${device.is_online ? "online" : "offline"}`}>
                  {device.is_online ? "在线" : "离线"}
                </span>
              </div>
              <div className="device-meta">
                <span>{device.machine_id}</span>
                <span>激活：{device.activated_at || "—"}</span>
                <span>心跳：{device.last_heartbeat || "—"}（{device.last_heartbeat_ago || "—"}）</span>
                <span>累计在线：{device.online_duration_formatted}</span>
              </div>
            </div>
            <button
              className="btn btn-sm btn-danger"
              onClick={() => handleUnbind(device)}
              disabled={unbindingId === device.machine_id}
            >
              {unbindingId === device.machine_id ? "解绑中..." : "解绑"}
            </button>
          </div>
        ))}
      </div>

      <style>{`
        .device-management-card {
          background: white;
          border: 1px solid var(--gray-200);
          border-radius: 8px;
          padding: 16px;
          margin-top: 16px;
        }
        .device-management-header {
          display: flex;
          justify-content: space-between;
          align-items: center;
          gap: 12px;
          margin-bottom: 12px;
        }
        .device-management-header h3 {
          font-size: 15px;
          font-weight: 600;
          color: var(--gray-800);
          margin: 0 0 4px;
        }
        .device-management-header p,
        .device-meta {
          color: var(--gray-500);
          font-size: 12px;
        }
        .device-management-error {
          background: rgba(239, 68, 68, 0.08);
          border: 1px solid rgba(239, 68, 68, 0.2);
          color: #ef4444;
          border-radius: 8px;
          padding: 8px 10px;
          font-size: 13px;
          margin-bottom: 10px;
        }
        .device-list {
          display: flex;
          flex-direction: column;
          gap: 8px;
        }
        .device-row {
          display: flex;
          justify-content: space-between;
          align-items: center;
          gap: 12px;
          padding: 12px;
          border: 1px solid var(--gray-100);
          border-radius: 8px;
          background: var(--gray-50);
        }
        .device-main {
          min-width: 0;
          flex: 1;
        }
        .device-title {
          display: flex;
          align-items: center;
          gap: 6px;
          font-size: 14px;
          font-weight: 600;
          color: var(--gray-800);
          margin-bottom: 6px;
        }
        .device-meta {
          display: flex;
          flex-wrap: wrap;
          gap: 8px 14px;
          line-height: 1.5;
        }
        .device-tag {
          border-radius: 999px;
          padding: 2px 8px;
          font-size: 11px;
          font-weight: 500;
        }
        .device-tag.current {
          background: #e0e7ff;
          color: #4338ca;
        }
        .device-tag.online {
          background: #dcfce7;
          color: #166534;
        }
        .device-tag.offline {
          background: var(--gray-200);
          color: var(--gray-600);
        }
        .device-empty {
          padding: 20px;
          text-align: center;
          color: var(--gray-400);
          font-size: 13px;
        }
      `}</style>
    </div>
  );
};
