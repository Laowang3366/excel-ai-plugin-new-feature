/**
 * 激活弹窗组件
 *
 * 状态机：loading → input → activating → success / error
 * - loading：首次启动，正在从主进程获取激活状态。
 * - input：展示卡密输入框，等待用户提交。
 * - activating：已提交，正在调用 IPC 执行激活。
 * - success：激活成功，展示对勾动画，1.5 秒后自动关闭弹窗。
 * - error：激活失败，在输入框下方展示红色错误提示。
 *
 * 视觉设计：
 * - 使用 fixed + z-index:9999 的全屏遮罩层（activation-overlay）
 * - 居中卡片（activation-card），深色渐变背景
 * - 成功状态展示绿色圆圈对勾，失败状态保留输入框以便重试
 *
 * 当 isLoading（Store 层正在加载）时，组件不论内部状态都优先展示 loading spinner。
 */

import React, { useState } from "react";
import { useActivationStore } from "../store/activationStore";

interface ActivationDialogProps {
  onActivated?: () => void;
}

export const ActivationDialog: React.FC<ActivationDialogProps> = ({ onActivated }) => {
  const { activate, isLoading } = useActivationStore();

  /**
   * 组件内部状态
   * - key: 卡密输入框的值，始终大写（见 onChange 处理）
   * - error: 激活失败时的错误文案，清空输入时自动重置
   * - activating: 正在调用 IPC 执行激活，按钮展示"验证中..."
   * - success: IPC 返回成功，展示成功界面后自动关闭
   */
  const [key, setKey] = useState("");
  const [error, setError] = useState("");
  const [activating, setActivating] = useState(false);
  const [success, setSuccess] = useState(false);

  /**
   * 提交激活表单
   *
   * 流程：
   * 1. 去除首尾空格后做空值校验
   * 2. 调用 Store 的 activate（内部通过 IPC → 主进程 → 许可证服务器）
   * 3. 成功 → 展示 success 界面，1.5s 后调用 onActivated 回调关闭弹窗
   * 4. 失败 → 在输入框下方展示错误信息，用户可重新输入
   */
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
    /**
     * loading 状态：Store 尚未完成初始化，展示 spinner
     * 此时不渲染表单，避免用户看到空白输入框
     */
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
                    /** 统一转为大写字母，提升卡密输入体验（卡密通常仅含大写字母与数字） */
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
