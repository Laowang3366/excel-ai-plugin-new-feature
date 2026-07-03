/**
 * OfficePreviewPanel — 右侧 Office 文件编辑侧栏
 *
 * 关联模块：
 * - utils/officeEditEvents.ts: 从工具结果提取 Open XML 编辑事件。
 * - components/ChatPage.tsx: 将当前会话事件传入本组件。
 */

import React from "react";
import type { OfficeEditEvent } from "../../utils/officeEditEvents";
import { Activity, CheckCircle, FileText } from "../common/IconMap";

interface OfficePreviewPanelProps {
  events: OfficeEditEvent[];
  isOpen: boolean;
  onToggle: () => void;
}

function fileNameOf(filePath: string): string {
  const parts = filePath.split(/[\\/]/);
  return parts[parts.length - 1] || filePath;
}

function timeOf(timestamp: number): string {
  return new Date(timestamp).toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" });
}

export const OfficePreviewPanel: React.FC<OfficePreviewPanelProps> = ({ events, isOpen, onToggle }) => {
  const latest = events[events.length - 1];

  return (
    <aside
      className={`office-preview-panel ${isOpen ? "open" : "collapsed"}`}
      aria-hidden={!isOpen}
    >
      {isOpen && (
        <>
          <div className="office-preview-header">
            <button
              className="office-preview-toggle office-preview-toggle-in-panel active"
              type="button"
              onClick={onToggle}
              title="隐藏文件编辑侧栏"
              aria-label="隐藏文件编辑侧栏"
              aria-pressed={true}
            >
              <Activity size={14} />
            </button>
          </div>

          {latest ? (
            <div className="office-preview-current">
              <div className="office-preview-file-icon">
                <FileText size={18} />
              </div>
              <div className="office-preview-file-meta">
                <div className="office-preview-file-name" title={latest.outputPath || latest.filePath}>
                  {fileNameOf(latest.outputPath || latest.filePath)}
                </div>
                <div className="office-preview-file-path" title={latest.outputPath || latest.filePath}>
                  {latest.outputPath || latest.filePath}
                </div>
              </div>
            </div>
          ) : (
            <div className="office-preview-empty">
              <FileText size={20} />
              <span>暂无文件编辑</span>
            </div>
          )}

          <div className="office-preview-events">
            {events.map((event) => {
              const actionSummary = typeof event.detail.summary === "string" ? event.detail.summary : "";
              return (
                <div key={event.id} className="office-preview-event">
                  <div className="office-preview-event-mark">
                    <CheckCircle size={14} />
                  </div>
                  <div className="office-preview-event-body">
                    <div className="office-preview-event-top">
                      <span className="office-preview-event-summary">{event.summary}</span>
                      <span className="office-preview-event-time">{timeOf(event.timestamp)}</span>
                    </div>
                    <div className="office-preview-event-file" title={event.filePath}>
                      {fileNameOf(event.filePath)}
                    </div>
                    {actionSummary && (
                      <div className="office-preview-text-preview">
                        {actionSummary.slice(0, 160)}
                      </div>
                    )}
                  </div>
                </div>
              );
            })}
          </div>
        </>
      )}
    </aside>
  );
};
