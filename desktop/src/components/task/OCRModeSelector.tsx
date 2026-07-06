import React from "react";

interface OCRModeSelectorProps {
  ocrMode: "image" | "invoice";
  onSelectImageMode: () => void;
  onSelectInvoiceMode: () => void;
}

export const OCRModeSelector: React.FC<OCRModeSelectorProps> = ({
  ocrMode,
  onSelectImageMode,
  onSelectInvoiceMode,
}) => {
  return (
    <div className="task-field">
      <label className="task-field-label">识别模式</label>
      <div className="task-select-group">
        <button
          className={`task-select-btn ${ocrMode === "image" ? "active" : ""}`}
          onClick={onSelectImageMode}
        >
          通用 OCR
        </button>
        <button
          className={`task-select-btn ${ocrMode === "invoice" ? "active" : ""}`}
          onClick={onSelectInvoiceMode}
        >
          发票识别
        </button>
      </div>
    </div>
  );
};
