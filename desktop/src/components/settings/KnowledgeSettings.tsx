/**
 * 知识库设置 — 管理 RAG 知识来源
 *
 * 功能：
 * - 启用/禁用知识库
 * - 查看已索引的知识来源列表
 * - 添加文件/文件夹索引
 * - 删除/重建索引
 */

import React, { useEffect, useState, useCallback } from "react";
import {
  BookOpen,
  Database,
  FolderOpen,
  FileScan,
  Trash2,
  RefreshCw,
  Search,
  Check,
  X,
  FileText,
} from "../common/IconMap";
import { useSettingsStore } from "../../store/settingsStore";
import { ipcApi } from "../../services/ipcApi";
import {
  formatKnowledgeFolderIndexSuccess,
  formatKnowledgeSourceStats,
  formatKnowledgeTime,
  getKnowledgeSourceTypeLabel,
  KNOWLEDGE_TEXT,
  type KnowledgeSource,
} from "./knowledgeSettingsText";

// ============================================================
// 组件
// ============================================================

export const KnowledgeSettings: React.FC = () => {
  const { language, knowledgeEnabled, setKnowledgeEnabled } = useSettingsStore();
  const text = KNOWLEDGE_TEXT[language];

  const [sources, setSources] = useState<KnowledgeSource[]>([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [reindexing, setReindexing] = useState(false);
  const [reindexProgress, setReindexProgress] = useState<{ current: number; total: number } | null>(
    null,
  );
  const [successMsg, setSuccessMsg] = useState<string | null>(null);
  const totalEntries = sources.reduce((sum, source) => sum + source.entryCount, 0);

  // ── 加载来源列表 ──
  const loadSources = useCallback(async () => {
    try {
      setLoading(true);
      const list = await ipcApi.knowledge.listSources();
      setSources(list as KnowledgeSource[]);
      setError(null);
    } catch (err: any) {
      setError(err.message || text.error);
    } finally {
      setLoading(false);
    }
  }, [text.error]);

  useEffect(() => {
    loadSources();
  }, [loadSources]);

  // ── 自动清除成功消息 ──
  useEffect(() => {
    if (successMsg) {
      const timer = setTimeout(() => setSuccessMsg(null), 3000);
      return () => clearTimeout(timer);
    }
  }, [successMsg]);

  // ── 添加文件 ──
  const handleAddFile = async () => {
    try {
      const result = await ipcApi.dialog.openFile();
      if (result.canceled || result.filePaths.length === 0) return;
      const filePath = result.filePaths[0];
      const r = await ipcApi.knowledge.indexFile(filePath);
      if (r.success) {
        setSuccessMsg(text.indexFileSuccess.replace("{count}", String(r.entryCount)));
        await loadSources();
      } else {
        setError(r.error || text.error);
      }
    } catch (err: any) {
      setError(err.message || text.error);
    }
  };

  // ── 添加文件夹 ──
  const handleAddFolder = async () => {
    try {
      const result = await ipcApi.dialog.openFolder();
      if (result.canceled || result.filePaths.length === 0) return;
      const folderPath = result.filePaths[0];
      const r = await ipcApi.knowledge.indexFolder(folderPath);
      if (Array.isArray(r)) {
        setSuccessMsg(formatKnowledgeFolderIndexSuccess(text, r));
        await loadSources();
      } else {
        setError((r as any)?.error || text.error);
      }
    } catch (err: any) {
      setError(err.message || text.error);
    }
  };

  // ── 删除来源 ──
  const handleDelete = async (sourcePath: string) => {
    if (!window.confirm(text.deleteConfirm)) return;
    try {
      const r = await ipcApi.knowledge.deleteFile(sourcePath);
      if (r.success) {
        setSuccessMsg(text.deleteSuccess);
        await loadSources();
      } else {
        setError(r.error || text.error);
      }
    } catch (err: any) {
      setError(err.message || text.error);
    }
  };

  // ── 重建全部索引 ──
  const handleReindexAll = async () => {
    if (!window.confirm(text.reindexConfirm)) return;
    try {
      setReindexing(true);
      setReindexProgress({ current: 0, total: 0 });
      const r = await ipcApi.knowledge.reindexAll();
      if (r.success) {
        setSuccessMsg(
          Array.isArray(r.results)
            ? formatKnowledgeFolderIndexSuccess(text, r.results)
            : text.indexFolderSuccess
                .replace("{success}", "0")
                .replace("{failed}", "0")
                .replace("{count}", "0"),
        );
        await loadSources();
      } else {
        setError(r.error || text.error);
      }
    } catch (err: any) {
      setError(err.message || text.error);
    } finally {
      setReindexing(false);
      setReindexProgress(null);
    }
  };

  return (
    <div className="settings-section-content">
      <h2>{text.title}</h2>
      <p className="settings-desc">{text.desc}</p>

      {/* 成功/错误提示 */}
      {successMsg && (
        <div className="settings-success-banner">
          <Check size={16} /> {successMsg}
        </div>
      )}
      {error && (
        <div className="settings-error-banner">
          <X size={16} /> {error}
          <button className="settings-dismiss-btn" onClick={() => setError(null)}>
            ×
          </button>
        </div>
      )}

      {/* 启用/禁用开关 */}
      <div className="settings-card">
        <div className="settings-switch-row knowledge-enable-row">
          <div className="settings-switch-info">
            <div className="settings-switch-label">
              <BookOpen size={18} />
              <span>{text.enableTitle}</span>
            </div>
            <div className="form-hint">{text.enableHint}</div>
          </div>
          <label className="switch">
            <input
              type="checkbox"
              checked={knowledgeEnabled}
              onChange={(e) => setKnowledgeEnabled(e.target.checked)}
            />
            <span className="switch-slider" />
          </label>
        </div>
      </div>

      {/* 知识来源列表 */}
      <div className="settings-card">
        <div className="settings-card-header knowledge-card-header">
          <Database size={18} />
          <span>{text.sourcesTitle}</span>
        </div>
        <p className="form-hint">{text.sourcesDesc}</p>

        {loading ? (
          <p className="form-hint">{text.indexing}</p>
        ) : sources.length === 0 ? (
          <div className="settings-empty-state">
            <Search size={32} />
            <p>{text.noSources}</p>
          </div>
        ) : (
          <div className="knowledge-source-list">
            <div className="knowledge-source-summary">
              {formatKnowledgeSourceStats(text, sources.length, totalEntries)}
            </div>
            {sources.map((s) => (
              <div className="knowledge-source-item" key={s.sourcePath}>
                <div className="knowledge-source-icon">
                  <FileText size={18} />
                </div>
                <div className="knowledge-source-main">
                  <div className="knowledge-source-title-row">
                    <span className="knowledge-source-name" title={s.sourceName}>
                      {s.sourceName}
                    </span>
                    <span className="type-badge">
                      {getKnowledgeSourceTypeLabel(text, s.sourceType)}
                    </span>
                  </div>
                  <div className="knowledge-source-path" title={s.sourcePath}>
                    {text.sourcePathLabel}: {s.sourcePath}
                  </div>
                  <div className="knowledge-source-meta">
                    <span>
                      {text.entriesCol}: {s.entryCount}
                    </span>
                    <span>
                      {text.indexedCol}: {formatKnowledgeTime(s.lastIndexed, language)}
                    </span>
                  </div>
                </div>
                <button
                  className="settings-action-btn icon-btn danger knowledge-source-delete"
                  onClick={() => handleDelete(s.sourcePath)}
                  title={text.deleteSource}
                >
                  <Trash2 size={14} />
                </button>
              </div>
            ))}
          </div>
        )}
      </div>

      {/* 操作按钮组 */}
      <div className="settings-card">
        <div className="knowledge-actions">
          <button className="settings-action-btn" onClick={handleAddFile} disabled={reindexing}>
            <FileScan size={16} />
            <span>{text.addFile}</span>
          </button>
          <button className="settings-action-btn" onClick={handleAddFolder} disabled={reindexing}>
            <FolderOpen size={16} />
            <span>{text.addFolder}</span>
          </button>
          <button
            className="settings-action-btn secondary"
            onClick={handleReindexAll}
            disabled={reindexing}
          >
            <RefreshCw size={16} className={reindexing ? "spin" : ""} />
            <span>
              {reindexing && reindexProgress
                ? text.reindexAllProgress
                    .replace("{current}", String(reindexProgress.current))
                    .replace("{total}", String(reindexProgress.total))
                : text.reindexAll}
            </span>
          </button>
        </div>
      </div>
    </div>
  );
};
