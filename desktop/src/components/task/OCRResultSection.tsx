import React from "react";
import { AlertTriangle, Ruler } from "../common/IconMap";
import type { OcrResult } from "./ocrTaskResultHelpers";

interface OCRResultSectionProps {
  result: OcrResult;
  fieldNames: string[];
  selectedFields: string[];
  outputRange: string;
  previewRows: string[][];
  canWriteResult: boolean;
  writeStatus: string;
  onSelectAllFields: () => void;
  onDeselectAllFields: () => void;
  onToggleField: (field: string) => void;
  onOutputRangeChange: (range: string) => void;
  onPickRange: () => void;
  onWriteToSheet: () => void;
}

export const OCRResultSection: React.FC<OCRResultSectionProps> = ({
  result,
  fieldNames,
  selectedFields,
  outputRange,
  previewRows,
  canWriteResult,
  writeStatus,
  onSelectAllFields,
  onDeselectAllFields,
  onToggleField,
  onOutputRangeChange,
  onPickRange,
  onWriteToSheet,
}) => {
  return (
    <>
      {result.errors.length > 0 && (
        <div className="ocr-errors">
          {result.errors.map((error, index) => (
            <p key={index} style={{ color: "var(--danger)", fontSize: 12 }}>
              <AlertTriangle size={12} style={{ verticalAlign: "middle", marginRight: 4 }} />{" "}
              {error}
            </p>
          ))}
        </div>
      )}

      {fieldNames.length > 0 && (
        <div className="task-field">
          <div className="task-field-label-row">
            <label className="task-field-label">选择字段</label>
            <div className="task-field-actions">
              <button className="task-link-btn" onClick={onSelectAllFields}>
                全选
              </button>
              <button className="task-link-btn" onClick={onDeselectAllFields}>
                全不选
              </button>
            </div>
          </div>
          <div className="ocr-field-grid">
            {fieldNames.map((field) => (
              <label key={field} className="ocr-field-checkbox">
                <input
                  type="checkbox"
                  checked={selectedFields.includes(field)}
                  onChange={() => onToggleField(field)}
                />
                <span>{field}</span>
              </label>
            ))}
          </div>
        </div>
      )}

      <div className="task-field">
        <label className="task-field-label">填入起始位置</label>
        <div className="range-input-row">
          <input
            className="task-field-input"
            value={outputRange}
            onChange={(event) => onOutputRangeChange(event.target.value)}
            placeholder="如 Sheet1!A1"
          />
          <button className="btn-pick-range" onClick={onPickRange}>
            <Ruler size={13} /> 选区
          </button>
        </div>
      </div>

      {previewRows.length > 0 && (
        <div className="task-field">
          <label className="task-field-label">
            预览（前 {Math.min(previewRows.length, 20)} 行）
          </label>
          <div className="ocr-preview-table-wrapper">
            <table className="ocr-preview-table">
              <thead>
                <tr>
                  {previewRows[0]?.map((cell, index) => (
                    <th key={index}>{cell}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {previewRows.slice(1).map((row, rowIndex) => (
                  <tr key={rowIndex}>
                    {row.map((cell, cellIndex) => (
                      <td key={cellIndex}>{cell}</td>
                    ))}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      )}

      <button
        className="task-submit-btn"
        disabled={!canWriteResult}
        onClick={onWriteToSheet}
        style={{ marginTop: 8 }}
      >
        写入单元格
      </button>
      {writeStatus && (
        <p style={{ marginTop: 8, fontSize: 12, color: "var(--text-muted)" }}>{writeStatus}</p>
      )}
    </>
  );
};
