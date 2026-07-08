/**
 * 激活弹窗组件
 *
 * 应用首次启动或激活失效时显示，要求用户输入卡密。
 */

import React, { useState } from "react";
import { useActivationStore } from "../store/activationStore";

interface ActivationDialogProps {
  onActivated?: () => void;
}

export const ActivationDialog: React.FC<ActivationDialogProps> = ({ onActivated }) => {
  const { activate, isLoading } = useActivationStore();

  const [key, setKey] = useState("");
  const [error, setError] = useState("");
  const [activating, setActivating] = useState(false);
  const [success, setSuccess] = useState(false);

  const handleActivate = async (e: React.FormEvent) => {
    e.preventDefault();
    const trimmedKey = key.trim();
    if (!trimmedKey) {
      setError("请输入卡密");
      return;
    }

    setError("");
    setActivating(true);

    try {
      const result = await activate(trimmedKey);
      if (result.success) {
        setSuccess(true);
        setTimeout(() => {
          onActivated?.();
        }, 1500);
      } else {
        setError(result.error || "激活失败，请检查卡密");
      }
    } catch (err: any) {
      setError(err.message || "网络错误");
    } finally {
      setActivating(false);
    }
  };

  if (isLoading) {
    return (
      <div className="activation-overlay">
        <div className="activation-card">
          <div className="spinner" />
          <p style={{ marginTop: 16, color: "#94a3b8" }}>加载中...</p>
        </div>
      </div>
    );
  }

  return (
    <div className="activation-overlay">
      <div className="activation-card">
        {success ? (
          <>
            <div className="activation-icon success">✓</div>
            <h2>激活成功！</h2>
            <p className="activation-desc">正在进入应用...</p>
          </>
        ) : (
          <>
            <div className="activation-icon">🔐</div>
            <h2>激活应用</h2>
            <p className="activation-desc">请输入激活卡密以继续使用</p>

            <form onSubmit={handleActivate} className="activation-form">
              <div className="activation-input-group">
                <input
                  type="text"
                  className="activation-input"
                  placeholder="请输入卡密（如：XXXX-XXXX-XXXX-XXXX）"
                  value={key}
                  onChange={(e) => {
                    setKey(e.target.value.toUpperCase());
                    setError("");
                  }}
                  autoFocus
                  disabled={activating}
                />
              </div>

              {error && <div className="activation-error">{error}</div>}

              <button
                type="submit"
                className="activation-btn"
                disabled={activating || !key.trim()}
              >
                {activating ? "验证中..." : "激 活"}
              </button>
            </form>
          </>
        )}
      </div>

      <style>{`
        .activation-overlay {
          position: fixed;
          top: 0;
          left: 0;
          right: 0;
          bottom: 0;
          background: linear-gradient(135deg, #1e293b 0%, #0f172a 100%);
          display: flex;
          align-items: center;
          justify-content: center;
          z-index: 9999;
        }
        .activation-card {
          background: #1e293b;
          border: 1px solid #334155;
          border-radius: 16px;
          padding: 40px;
          width: 420px;
          max-width: 90vw;
          text-align: center;
          box-shadow: 0 25px 50px rgba(0, 0, 0, 0.5);
        }
        .activation-icon {
          font-size: 48px;
          margin-bottom: 16px;
        }
        .activation-icon.success {
          width: 64px;
          height: 64px;
          background: #22c55e;
          border-radius: 50%;
          display: flex;
          align-items: center;
          justify-content: center;
          margin: 0 auto 16px;
          font-size: 28px;
          color: white;
          font-weight: bold;
        }
        .activation-card h2 {
          color: #f1f5f9;
          font-size: 22px;
          font-weight: 600;
          margin-bottom: 8px;
        }
        .activation-desc {
          color: #94a3b8;
          font-size: 14px;
          margin-bottom: 24px;
        }
        .activation-form {
          display: flex;
          flex-direction: column;
          gap: 12px;
        }
        .activation-input-group {
          position: relative;
        }
        .activation-input {
          width: 100%;
          padding: 12px 16px;
          background: #0f172a;
          border: 1px solid #334155;
          border-radius: 8px;
          color: #f1f5f9;
          font-size: 16px;
          font-family: "SF Mono", "Fira Code", Consolas, monospace;
          letter-spacing: 2px;
          text-align: center;
          outline: none;
          transition: border-color 0.2s;
          box-sizing: border-box;
        }
        .activation-input:focus {
          border-color: #6366f1;
          box-shadow: 0 0 0 3px rgba(99, 102, 241, 0.2);
        }
        .activation-input::placeholder {
          color: #475569;
          font-size: 13px;
          letter-spacing: 0;
          font-family: inherit;
        }
        .activation-error {
          background: rgba(239, 68, 68, 0.1);
          border: 1px solid rgba(239, 68, 68, 0.3);
          color: #fca5a5;
          padding: 8px 12px;
          border-radius: 8px;
          font-size: 13px;
          text-align: center;
        }
        .activation-btn {
          padding: 12px;
          background: linear-gradient(135deg, #6366f1, #8b5cf6);
          border: none;
          border-radius: 8px;
          color: white;
          font-size: 16px;
          font-weight: 600;
          cursor: pointer;
          transition: opacity 0.2s, transform 0.1s;
        }
        .activation-btn:hover:not(:disabled) {
          opacity: 0.9;
        }
        .activation-btn:active:not(:disabled) {
          transform: scale(0.98);
        }
        .activation-btn:disabled {
          opacity: 0.5;
          cursor: not-allowed;
        }
      `}</style>
    </div>
  );
};
