import React from "react";

interface OCRRecognizeButtonProps {
  ocrMode: "image" | "invoice";
  fileCount: number;
  recognizing: boolean;
  onRecognize: () => void;
}

export const OCRRecognizeButton: React.FC<OCRRecognizeButtonProps> = ({
  ocrMode,
  fileCount,
  recognizing,
  onRecognize,
}) => {
  return (
    <button
      className="task-submit-btn"
      disabled={fileCount === 0 || recognizing}
      onClick={onRecognize}
    >
      {recognizing
        ? (ocrMode === "invoice" ? "识别并提取字段中..." : "识别中...")
        : "开始识别"}
    </button>
  );
};
