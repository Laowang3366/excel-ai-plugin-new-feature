import React from "react";
import { FileText, Image, Paperclip, X } from "../common/IconMap";
import { ACCEPTED_OCR_TYPES } from "./ocrTaskFileHelpers";

interface OCRFileUploadSectionProps {
  ocrMode: "image" | "invoice";
  files: File[];
  dragOver: boolean;
  fileInputRef: React.RefObject<HTMLInputElement>;
  onFileChange: (event: React.ChangeEvent<HTMLInputElement>) => void;
  onDrop: (event: React.DragEvent) => void;
  onDragOverChange: (dragOver: boolean) => void;
  onRemoveFile: (index: number) => void;
}

export const OCRFileUploadSection: React.FC<OCRFileUploadSectionProps> = ({
  ocrMode,
  files,
  dragOver,
  fileInputRef,
  onFileChange,
  onDrop,
  onDragOverChange,
  onRemoveFile,
}) => {
  return (
    <div className="task-field">
      <label className="task-field-label">
        上传文件（{ocrMode === "invoice" ? "最多 10 个" : "1 个"}）
      </label>
      <div
        className={`ocr-drop-zone ${dragOver ? "drag-over" : ""}`}
        onDragOver={(event) => {
          event.preventDefault();
          onDragOverChange(true);
        }}
        onDragLeave={() => onDragOverChange(false)}
        onDrop={onDrop}
        onClick={() => fileInputRef.current?.click()}
      >
        <input
          ref={fileInputRef}
          type="file"
          accept={ACCEPTED_OCR_TYPES}
          multiple={ocrMode === "invoice"}
          onChange={onFileChange}
          style={{ display: "none" }}
        />
        <div className="ocr-drop-content">
          <span style={{ fontSize: 24, opacity: 0.5 }}><Paperclip size={24} /></span>
          <p>
            {dragOver
              ? "松开以上传"
              : "拖拽文件到此处，或点击选择"}
          </p>
          <p style={{ fontSize: 11, color: "var(--text-faint)" }}>
            支持 PNG、JPG、WebP、BMP、PDF
          </p>
        </div>
      </div>

      {files.length > 0 && (
        <div className="ocr-file-list">
          {files.map((file, index) => (
            <div key={index} className="ocr-file-item">
              <span className="ocr-file-icon">
                {file.type === "application/pdf" ? <FileText size={14} /> : <Image size={14} />}
              </span>
              <span className="ocr-file-name">{file.name}</span>
              <button
                className="ocr-file-remove"
                onClick={() => onRemoveFile(index)}
              >
                <X size={10} />
              </button>
            </div>
          ))}
        </div>
      )}
    </div>
  );
};
