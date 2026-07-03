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
import { BookOpen, Database, FolderOpen, FileScan, Trash2, RefreshCw, Search, Check, X } from "../common/IconMap";
import { useSettingsStore } from "../../store/settingsStore";
import { ipcApi } from "../../services/ipcApi";

// ============================================================
// 本地化文本
// ============================================================

const KNOWLEDGE_TEXT = {
  "zh-CN": {
    title: "知识库",
    desc: "管理本地知识库，自动为 AI 助手提供工作簿结构、字段含义等上下文信息。",
    enableTitle: "启用知识库",
    enableHint: "开启后 AI 助手会自动检索相关知识并注入对话上下文。",
    sourcesTitle: "知识来源",
    sourcesDesc: "已索引的文件列表，AI 在对话时会自动参考这些知识。",
    noSources: "暂无知识来源。点击下方按钮添加文件或文件夹。",
    sourceCol: "来源文件",
    typeCol: "类型",
    entriesCol: "条目数",
    indexedCol: "索引时间",
    actionsCol: "操作",
    addFile: "添加文件",
    addFolder: "添加文件夹",
    reindexAll: "重建全部索引",
    indexing: "索引中...",
    deleteSource: "删除",
    deleteConfirm: "确定要删除该来源的索引吗？",
    reindexConfirm: "确定要重建全部索引吗？这可能需要一些时间。",
    reindexAllProgress: "正在重建全部索引 ({current}/{total})...",
    indexFileSuccess: "文件索引完成，共 {count} 条",
    indexFolderSuccess: "文件夹索引完成",
    deleteSuccess: "已删除索引",
    error: "操作失败",
    workbook: "工作簿",
    document: "文档",
    note: "笔记",
    agents_md: "项目知识",
  },
  "en-US": {
    title: "Knowledge Base",
    desc: "Manage local knowledge base that provides workbook structure, field meanings, and other context to the AI assistant.",
    enableTitle: "Enable Knowledge Base",
    enableHint: "When enabled, the AI automatically retrieves relevant knowledge and injects it into conversation context.",
    sourcesTitle: "Knowledge Sources",
    sourcesDesc: "Indexed files that the AI references during conversations.",
    noSources: "No knowledge sources yet. Click below to add files or folders.",
    sourceCol: "Source File",
    typeCol: "Type",
    entriesCol: "Entries",
    indexedCol: "Indexed At",
    actionsCol: "Actions",
    addFile: "Add File",
    addFolder: "Add Folder",
    reindexAll: "Rebuild All Indexes",
    indexing: "Indexing...",
    deleteSource: "Delete",
    deleteConfirm: "Delete this source?",
    reindexConfirm: "Rebuild all indexes? This may take some time.",
    reindexAllProgress: "Rebuilding all indexes ({current}/{total})...",
    indexFileSuccess: "File indexed, {count} entries created",
    indexFolderSuccess: "Folder indexed",
    deleteSuccess: "Index deleted",
    error: "Operation failed",
    workbook: "Workbook",
    document: "Document",
    note: "Note",
    agents_md: "Project Knowledge",
  },
} as const;

// ============================================================
// 类型
// ============================================================

interface KnowledgeSource {
  sourcePath: string;
  sourceName: string;
  sourceType: string;
  entryCount: number;
  firstIndexed: number;
  lastIndexed: number;
  fileHash: string;
}

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
  const [reindexProgress, setReindexProgress] = useState<{ current: number; total: number } | null>(null);
  const [successMsg, setSuccessMsg] = useState<string | null>(null);

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

  // ── 格式化时间 ──
  const formatTime = (ts: number): string => {
    const d = new Date(ts);
    return d.toLocaleString(language === "zh-CN" ? "zh-CN" : "en-US");
  };

  // ── 类型标签 ──
  const typeLabel = (type: string): string => {
    const map: Record<string, string> = {
      workbook: text.workbook,
      document: text.document,
      note: text.note,
      agents_md: text.agents_md,
      xlsx: text.workbook,
      csv: text.document,
      md: text.agents_md,
      txt: text.document,
    };
    return map[type] || type;
  };

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
        setSuccessMsg(text.indexFolderSuccess);
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
        setSuccessMsg(text.indexFolderSuccess);
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
          <button className="settings-dismiss-btn" onClick={() => setError(null)}>×</button>
        </div>
      )}

      {/* 启用/禁用开关 */}
      <div className="settings-card">
        <div className="settings-switch-row">
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
        <div className="settings-card-header">
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
          <table className="knowledge-source-table">
            <thead>
              <tr>
                <th>{text.sourceCol}</th>
                <th>{text.typeCol}</th>
                <th>{text.entriesCol}</th>
                <th>{text.indexedCol}</th>
                <th>{text.actionsCol}</th>
              </tr>
            </thead>
            <tbody>
              {sources.map((s) => (
                <tr key={s.sourcePath}>
                  <td className="source-name" title={s.sourcePath}>
                    {s.sourceName}
                  </td>
                  <td>
                    <span className="type-badge">{typeLabel(s.sourceType)}</span>
                  </td>
                  <td>{s.entryCount}</td>
                  <td className="indexed-time">{formatTime(s.lastIndexed)}</td>
                  <td>
                    <button
                      className="settings-action-btn icon-btn danger"
                      onClick={() => handleDelete(s.sourcePath)}
                      title={text.deleteSource}
                    >
                      <Trash2 size={14} />
                    </button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
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
            <span>{reindexing && reindexProgress
              ? text.reindexAllProgress
                .replace("{current}", String(reindexProgress.current))
                .replace("{total}", String(reindexProgress.total))
              : text.reindexAll
            }</span>
          </button>
        </div>
      </div>
    </div>
  );
};
