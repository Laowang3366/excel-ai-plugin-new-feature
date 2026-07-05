/**
 * 生成公式 — 任务编排面板
 *
 * 对齐桌面端任务面板的 formula 字段：
 * 1. 数据源选区（多选，chip 展示）
 * 2. 答案参考样例（单选区）
 * 3. 答案填入锚点/选区
 * 4. 需求说明（textarea）
 */

import React, { useState, useCallback, useEffect } from "react";
import { Hash, Ruler, X } from "../common/IconMap";
import { ipcApi } from "../../services/ipcApi";
import { pickExcelRange } from "../../utils/chatHelpers";
import {
  buildFormulaTaskPayload,
  getHostEnvironmentLabel,
  normalizeHostEnvironment,
  type HostEnvironment,
  type ReferenceSampleMode,
} from "../../utils/taskComposerPayloads";

export interface FormulaTaskDraft {
  dataSourceRanges: string[];
  dataSourceInput: string;
  referenceSampleRange: string;
  referenceSampleMode: ReferenceSampleMode;
  outputRange: string;
  hostEnvironment: HostEnvironment;
  task: string;
}

interface FormulaTaskComposerPanelProps {
  onSubmit: (payload: string) => void;
  onClose: () => void;
  embedded?: boolean;
  draft?: FormulaTaskDraft;
  onDraftChange?: (draft: FormulaTaskDraft) => void;
}

export const FormulaTaskComposerPanel: React.FC<FormulaTaskComposerPanelProps> = ({
  onSubmit,
  onClose,
  embedded = false,
  draft,
  onDraftChange,
}) => {
  const [dataSourceRanges, setDataSourceRanges] = useState<string[]>(draft?.dataSourceRanges ?? []);
  const [dataSourceInput, setDataSourceInput] = useState(draft?.dataSourceInput ?? "");
  const [referenceSampleRange, setReferenceSampleRange] = useState(draft?.referenceSampleRange ?? "");
  const [referenceSampleMode, setReferenceSampleMode] = useState<ReferenceSampleMode>(draft?.referenceSampleMode ?? "partial");
  const [outputRange, setOutputRange] = useState(draft?.outputRange ?? "");
  const [hostEnvironment, setHostEnvironment] = useState<HostEnvironment>(draft?.hostEnvironment ?? "unknown");
  const [task, setTask] = useState(draft?.task ?? "");

  useEffect(() => {
    onDraftChange?.({
      dataSourceRanges,
      dataSourceInput,
      referenceSampleRange,
      referenceSampleMode,
      outputRange,
      hostEnvironment,
      task,
    });
  }, [
    dataSourceRanges,
    dataSourceInput,
    referenceSampleRange,
    referenceSampleMode,
    outputRange,
    hostEnvironment,
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
          setDataSourceRanges((prev) =>
            prev.includes(rangeStr) ? prev : [...prev, rangeStr]
          );
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

  // 组装提交内容
  const handleSubmit = () => {
    onSubmit(buildFormulaTaskPayload({
      dataSourceRanges,
      dataSourceInput,
      referenceSampleRange,
      referenceSampleMode,
      outputRange,
      hostEnvironment,
      task,
    }));
  };

  return (
    <div className="task-composer-panel">
      {!embedded && (
        <div className="task-composer-title">
          <Hash size={16} /> 生成公式
          <button
            className="task-close-btn"
            onClick={onClose}
          >
            <X size={14} />
          </button>
        </div>
      )}

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

      {/* 答案参考样例 */}
      <div className="task-field">
        <label className="task-field-label">答案参考样例</label>
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

      <div className="task-field">
        <label className="task-field-label">当前环境</label>
        <div className="task-select-group">
          <button
            className={`task-select-btn ${hostEnvironment === "wps" ? "active" : ""}`}
            onClick={() => setHostEnvironment("wps")}
          >
            WPS
          </button>
          <button
            className={`task-select-btn ${hostEnvironment === "microsoft_excel" ? "active" : ""}`}
            onClick={() => setHostEnvironment("microsoft_excel")}
          >
            Microsoft Excel
          </button>
        </div>
        <button className="task-link-btn task-link-btn--inline" onClick={refreshHostEnvironment}>
          同步当前连接环境：{getHostEnvironmentLabel(hostEnvironment)}
        </button>
      </div>

      {/* 答案填入锚点/选区 */}
      <div className="task-field">
        <label className="task-field-label">答案填入锚点/选区</label>
        <div className="range-input-row">
          <input
            className="task-field-input"
            value={outputRange}
            onChange={(e) => setOutputRange(e.target.value)}
            placeholder="留空则由 Agent 选择空白区域"
          />
          <button className="btn-pick-range" onClick={() => pickCurrentSelection("output")}>
            <Ruler size={13} /> 选区
          </button>
        </div>
      </div>

      {/* 需求说明 */}
      <div className="task-field">
        <label className="task-field-label">需求说明</label>
        <textarea
          className="task-field-textarea"
          value={task}
          onChange={(e) => setTask(e.target.value)}
          placeholder="例如：根据产品线、区域和销售额计算提成金额"
        />
      </div>

      <button className="task-submit-btn" onClick={handleSubmit}>
        填入输入框并发送
      </button>
    </div>
  );
};
