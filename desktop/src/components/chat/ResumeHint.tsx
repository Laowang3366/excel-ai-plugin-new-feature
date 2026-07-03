/**
 * 中断恢复提示 — 用户最需要的功能
 *
 * 当对话因 max token 或中断而停止时，展示蓝色提示条，
 * 引导用户输入继续消息，AI 会从断点恢复。
 */

import React from "react";
import { Info, Play, RefreshCw, ClipboardList } from "../common/IconMap";
import { useSettingsStore } from "../../store/settingsStore";
import { getAppText } from "../../i18n";

interface ResumeHintProps {
  message: string;
  onResume?: () => void;
  /** 快速恢复时填入输入框的回调 */
  onFillInput?: (text: string) => void;
}

export const ResumeHint: React.FC<ResumeHintProps> = ({ message, onResume, onFillInput }) => {
  const { language } = useSettingsStore();
  const text = getAppText(language);

  const handleQuickResume = (text: string) => {
    if (onFillInput) {
      onFillInput(text);
    }
  };

  return (
    <div className="resume-hint">
      <div className="resume-hint-icon"><Info size={16} /></div>
      <div className="resume-hint-content">
        <div className="resume-hint-text">{message}</div>
        <div className="resume-quick-actions">
          <button
            className="resume-btn"
            onClick={() => handleQuickResume(text.assistant.resumeContinuePrompt)}
          >
            <Play size={13} /> {text.assistant.resumeContinue}
          </button>
          <button
            className="resume-btn"
            onClick={() => handleQuickResume(text.assistant.resumeBreakpointPrompt)}
          >
            <RefreshCw size={13} /> {text.assistant.resumeFromBreakpoint}
          </button>
          <button
            className="resume-btn"
            onClick={() => handleQuickResume(text.assistant.resumeSummarizePrompt)}
          >
            <ClipboardList size={13} /> {text.assistant.resumeSummarize}
          </button>
        </div>
      </div>
    </div>
  );
};
