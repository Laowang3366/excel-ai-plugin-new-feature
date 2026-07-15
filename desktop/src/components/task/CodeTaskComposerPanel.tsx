/**
 * 代码答疑 — 任务编排面板
 *
 * 对齐桌面端任务面板的 code 字段：
 * 1. 数据源选区（多选，chip 展示）
 * 2. 参考样例选区（单选区）
 * 3. 输出/操作锚点（单选区）
 * 4. 运行环境（WPS / Office）
 * 5. 首选语言（自动 / JS / VBA / Python）
 * 6. 代码需求说明（textarea）
 */

import React, { useState, useCallback, useEffect } from "react";
import { Code, Ruler, X } from "../common/IconMap";
import { ipcApi } from "../../services/ipcApi";
import { pickExcelRange } from "../../utils/chatHelpers";
import {
  buildCodeTaskPayload,
  getHostEnvironmentLabel,
  normalizeHostEnvironment,
  type HostEnvironment,
  type PreferredLanguage,
  type ReferenceSampleMode,
} from "../../utils/taskComposerPayloads";
import {
  CODE_TASK_HOST_OPTIONS,
  CODE_TASK_LANGUAGE_OPTIONS,
  type CodeTaskDraft,
} from "./codeTaskComposerModel";

export type { CodeTaskDraft } from "./codeTaskComposerModel";

interface CodeTaskComposerPanelProps {
  onSubmit: (payload: string) => void;
  onClose: () => void;
  embedded?: boolean;
  draft?: CodeTaskDraft;
  onDraftChange?: (draft: CodeTaskDraft) => void;
}

export const CodeTaskComposerPanel: React.FC<CodeTaskComposerPanelProps> = ({
  onSubmit,
  onClose,
  embedded = false,
  draft,
  onDraftChange,
}) => {
  const [dataSourceRanges, setDataSourceRanges] = useState<string[]>(draft?.dataSourceRanges ?? []);
  const [dataSourceInput, setDataSourceInput] = useState(draft?.dataSourceInput ?? "");
  const [referenceSampleRange, setReferenceSampleRange] = useState(
    draft?.referenceSampleRange ?? "",
  );
  const [referenceSampleMode, setReferenceSampleMode] = useState<ReferenceSampleMode>(
    draft?.referenceSampleMode ?? "partial",
  );
  const [outputRange, setOutputRange] = useState(draft?.outputRange ?? "");
  const [hostEnvironment, setHostEnvironment] = useState<HostEnvironment>(
    draft?.hostEnvironment ?? "unknown",
  );
  const [preferredLanguage, setPreferredLanguage] = useState<PreferredLanguage>(
    draft?.preferredLanguage ?? "auto",
  );
  const [task, setTask] = useState(draft?.task ?? "");

  useEffect(() => {
    onDraftChange?.({
      dataSourceRanges,
      dataSourceInput,
      referenceSampleRange,
      referenceSampleMode,
      outputRange,
      hostEnvironment,
      preferredLanguage,
      task,
    });
  }, [
    dataSourceRanges,
    dataSourceInput,
    referenceSampleRange,
    referenceSampleMode,
    outputRange,
    hostEnvironment,
    preferredLanguage,
    task,
    onDraftChange,
  ]);

  const refreshHostEnvironment = useCallback(async () => {
    const status = await ipcApi.excel.detectStatus();
    setHostEnvironment(normalizeHostEnvironment(status));
  }, []);

  useEffect(() => {
    refreshHostEnvironment().catch(() => setHostEnvironment("unknown"));
  }, [refreshHostEnvironment]);

  // 从 Excel 当前选中区域读取
  const pickCurrentSelection = useCallback(async (field: "datasource" | "reference" | "output") => {
    try {
      const rangeStr = await pickExcelRange();
      if (rangeStr) {
        if (field === "datasource") {
          setDataSourceRanges((prev) => (prev.includes(rangeStr) ? prev : [...prev, rangeStr]));
        } else if (field === "reference") {
          setReferenceSampleRange(rangeStr);
        } else {
          setOutputRange(rangeStr);
        }
      } else {
        alert("未获取到选区，请确认已在 Excel/WPS 中选中了单元格");
      }
    } catch (err: any) {
      alert(`获取选区失败: ${err.message || "请确认 Excel/WPS 已打开并选中了单元格"}`);
    }
  }, []);

  const removeDataSourceRange = useCallback((index: number) => {
    setDataSourceRanges((prev) => prev.filter((_, i) => i !== index));
  }, []);

  // 组装提交内容 — 对齐 buildCodeAssistanceAgentDraft 的 displayText 格式
  const handleSubmit = () => {
    onSubmit(
      buildCodeTaskPayload({
        dataSourceRanges,
        dataSourceInput,
        referenceSampleRange,
        referenceSampleMode,
        outputRange,
        hostEnvironment,
        preferredLanguage,
        task,
      }),
    );
  };

  return (
    <div className="task-composer-panel">
      {!embedded && (
        <div className="task-composer-title">
          <Code size={16} /> 代码答疑
          <button className="task-close-btn" onClick={onClose}>
            <X size={14} />
          </button>
        </div>
      )}

      {/* 运行环境 + 首选语言（并排） */}
      <div className={`task-field-row${embedded ? " task-field-row--stacked" : ""}`}>
        <div className="task-field task-field--half">
          <label className="task-field-label">运行环境</label>
          <div className="task-select-group">
            {CODE_TASK_HOST_OPTIONS.map((opt) => (
              <button
                key={opt.value}
                className={`task-select-btn ${hostEnvironment === opt.value ? "active" : ""}`}
                onClick={() => setHostEnvironment(opt.value)}
              >
                {opt.label}
              </button>
            ))}
          </div>
          <button className="task-link-btn task-link-btn--inline" onClick={refreshHostEnvironment}>
            同步当前连接环境：{getHostEnvironmentLabel(hostEnvironment)}
          </button>
        </div>
        <div className="task-field task-field--half">
          <label className="task-field-label">首选语言</label>
          <div className="task-select-group">
            {CODE_TASK_LANGUAGE_OPTIONS.map((opt) => (
              <button
                key={opt.value}
                className={`task-select-btn ${preferredLanguage === opt.value ? "active" : ""}`}
                onClick={() => setPreferredLanguage(opt.value)}
              >
                {opt.label}
              </button>
            ))}
          </div>
        </div>
      </div>

      {/* 数据源选区 */}
      <div className="task-field">
        <label className="task-field-label">数据源选区</label>
        <div className="range-input-row">
          <input
            className="task-field-input"
            value={dataSourceInput}
            onChange={(e) => setDataSourceInput(e.target.value)}
            placeholder="如 Sheet1!A1:F24"
          />
          <button className="btn-pick-range" onClick={() => pickCurrentSelection("datasource")}>
            <Ruler size={13} /> 选区
          </button>
        </div>
        {dataSourceRanges.length > 0 && (
          <div className="range-chips">
            {dataSourceRanges.map((range, i) => (
              <span key={i} className="range-chip">
                {range}
                <button className="range-chip-remove" onClick={() => removeDataSourceRange(i)}>
                  <X size={10} />
                </button>
              </span>
            ))}
          </div>
        )}
      </div>

      {/* 参考样例选区 */}
      <div className="task-field">
        <label className="task-field-label">参考样例选区</label>
        <div className="range-input-row">
          <input
            className="task-field-input"
            value={referenceSampleRange}
            onChange={(e) => setReferenceSampleRange(e.target.value)}
            placeholder="如 Sheet1!H2:H6"
          />
          <button className="btn-pick-range" onClick={() => pickCurrentSelection("reference")}>
            <Ruler size={13} /> 选区
          </button>
        </div>
      </div>

      <div className="task-field">
        <label className="task-field-label">样例类型</label>
        <div className="task-select-group">
          <button
            className={`task-select-btn ${referenceSampleMode === "partial" ? "active" : ""}`}
            onClick={() => setReferenceSampleMode("partial")}
          >
            部分样例
          </button>
          <button
            className={`task-select-btn ${referenceSampleMode === "complete" ? "active" : ""}`}
            onClick={() => setReferenceSampleMode("complete")}
          >
            完整样例
          </button>
        </div>
      </div>

      {/* 输出/操作锚点 */}
      <div className="task-field">
        <label className="task-field-label">输出/操作锚点</label>
        <div className="range-input-row">
          <input
            className="task-field-input"
            value={outputRange}
            onChange={(e) => setOutputRange(e.target.value)}
            placeholder="留空则由 Agent 自主选择"
          />
          <button className="btn-pick-range" onClick={() => pickCurrentSelection("output")}>
            <Ruler size={13} /> 选区
          </button>
        </div>
      </div>

      {/* 代码需求说明 */}
      <div className="task-field">
        <label className="task-field-label">代码需求说明</label>
        <textarea
          className="task-field-textarea"
          value={task}
          onChange={(e) => setTask(e.target.value)}
          placeholder="例如：编写 VBA 宏，将选中区域的数据按部门汇总到新工作表"
        />
      </div>

      <button className="task-submit-btn" onClick={handleSubmit}>
        填入输入框并发送
      </button>
    </div>
  );
};
