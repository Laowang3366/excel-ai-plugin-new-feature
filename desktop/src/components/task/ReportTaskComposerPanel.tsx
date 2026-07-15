import React, { useCallback, useEffect, useState } from "react";
import { FileBarChart, FolderOpen, Ruler, X } from "../common/IconMap";
import { ipcApi } from "../../services/ipcApi";
import { pickExcelRange } from "../../utils/chatHelpers";
import { buildReportTaskPayload, type ReportOutputFormat } from "../../utils/taskComposerPayloads";

export interface ReportTaskDraft {
  range: string;
  task: string;
  outputFormat: ReportOutputFormat;
  storagePath: string;
}

interface ReportTaskComposerPanelProps {
  onSubmit: (payload: string) => void;
  onClose: () => void;
  embedded?: boolean;
  draft?: ReportTaskDraft;
  onDraftChange?: (draft: ReportTaskDraft) => void;
}

const OUTPUT_FORMAT_OPTIONS: Array<{ value: ReportOutputFormat; label: string }> = [
  { value: "excel", label: "Excel" },
  { value: "word", label: "Word 文档" },
  { value: "ppt", label: "PPT" },
];

export const ReportTaskComposerPanel: React.FC<ReportTaskComposerPanelProps> = ({
  onSubmit,
  onClose,
  embedded = false,
  draft,
  onDraftChange,
}) => {
  const [range, setRange] = useState(draft?.range ?? "");
  const [task, setTask] = useState(draft?.task ?? "");
  const [outputFormat, setOutputFormat] = useState<ReportOutputFormat>(
    draft?.outputFormat ?? "excel",
  );
  const [storagePath, setStoragePath] = useState(draft?.storagePath ?? "桌面");

  useEffect(() => {
    onDraftChange?.({
      range,
      task,
      outputFormat,
      storagePath,
    });
  }, [range, task, outputFormat, storagePath, onDraftChange]);

  const pickCurrentSelection = useCallback(async () => {
    const nextRange = await pickExcelRange();
    if (nextRange) {
      setRange(nextRange);
    } else {
      alert("未获取到选区，请确认已在 Excel/WPS 中选中了单元格");
    }
  }, []);

  const pickStoragePath = useCallback(async () => {
    const result = await ipcApi.dialog.openFolder();
    if (!result.canceled && result.filePaths[0]) {
      setStoragePath(result.filePaths[0]);
    }
  }, []);

  const handleSubmit = () => {
    onSubmit(
      buildReportTaskPayload({
        range,
        task,
        outputFormat,
        storagePath,
      }),
    );
  };

  return (
    <div className="task-composer-panel">
      {!embedded && (
        <div className="task-composer-title">
          <FileBarChart size={16} /> 报告生成
          <button className="task-close-btn" onClick={onClose}>
            <X size={14} />
          </button>
        </div>
      )}

      <div className="task-field">
        <label className="task-field-label">输出格式</label>
        <div className="task-select-group">
          {OUTPUT_FORMAT_OPTIONS.map((option) => (
            <button
              key={option.value}
              className={`task-select-btn ${outputFormat === option.value ? "active" : ""}`}
              onClick={() => setOutputFormat(option.value)}
            >
              {option.label}
            </button>
          ))}
        </div>
      </div>

      <div className="task-field">
        <label className="task-field-label">数据源选区</label>
        <div className="range-input-row">
          <input
            className="task-field-input"
            value={range}
            onChange={(event) => setRange(event.target.value)}
            placeholder="如 Sheet1!A1:F24"
          />
          <button className="btn-pick-range" onClick={pickCurrentSelection}>
            <Ruler size={13} /> 选区
          </button>
        </div>
      </div>

      {outputFormat !== "excel" && (
        <div className="task-field">
          <label className="task-field-label">存储路径</label>
          <div className="range-input-row">
            <input
              className="task-field-input"
              value={storagePath}
              onChange={(event) => setStoragePath(event.target.value)}
              placeholder="默认桌面"
            />
            <button className="btn-pick-range" onClick={pickStoragePath}>
              <FolderOpen size={13} /> 选择
            </button>
          </div>
        </div>
      )}

      <div className="task-field">
        <label className="task-field-label">报告需求</label>
        <textarea
          className="task-field-textarea"
          value={task}
          onChange={(event) => setTask(event.target.value)}
          placeholder="例如：根据选区生成经营分析报告，包含趋势、异常和建议"
        />
      </div>

      <button className="task-submit-btn" onClick={handleSubmit}>
        填入输入框并发送
      </button>
    </div>
  );
};
