/**
 * OCR / 发票识别 — 任务编排面板
 *
 * 对齐桌面端任务面板的 OCR 字段：
 * 1. 识别模式（通用 OCR / 发票识别）
 * 2. 文件上传（拖拽/点击，支持图片和 PDF）
 * 3. 识别按钮 → 调用当前模型 API 的视觉能力
 * 4. 字段选择（识别结果出来后，勾选要填入的字段）
 * 5. 目标单元格（填入起始位置）
 * 6. 预览表格 + "写入单元格" 按钮
 *
 * 注意：OCR 流程在功能面板内静默完成，识别结果回填字段选择与预览区；
 * 用户确认后再写入工作表，不向聊天区发送内部任务提示词。
 */

import React, { useState, useRef, useCallback, useEffect } from "react";
import { FileScan, X } from "../common/IconMap";
import { ipcApi } from "../../services/ipcApi";
import { pickExcelRange } from "../../utils/chatHelpers";
import { OCRFileUploadSection } from "./OCRFileUploadSection";
import { OCRModeSelector } from "./OCRModeSelector";
import { OCRResultSection } from "./OCRResultSection";
import {
  buildOcrPreviewRows,
  buildOcrWriteValues,
  canWriteOcrResult,
  extractOcrFieldNames,
  type OcrResult,
} from "./ocrTaskResultHelpers";
import {
  isAcceptedOcrFile,
  isLikelyInvoiceFile,
  resolveOcrFilePaths,
  resolveWriteTarget,
} from "./ocrTaskFileHelpers";

export {
  buildOcrPreviewRows,
  buildOcrWriteValues,
  canWriteOcrResult,
  extractOcrFieldNames,
};
export type { OcrInvoiceItem, OcrResult } from "./ocrTaskResultHelpers";

export type OcrMode = "image" | "invoice";

export interface OCRTaskDraft {
  ocrMode: OcrMode;
  files: File[];
  recognizing: boolean;
  result: OcrResult | null;
  selectedFields: string[];
  outputRange: string;
  dragOver: boolean;
}

interface OCRTaskComposerPanelProps {
  onSubmit: (payload: string) => void;
  onClose: () => void;
  embedded?: boolean;
  draft?: OCRTaskDraft;
  onDraftChange?: (draft: OCRTaskDraft) => void;
}

export const OCRTaskComposerPanel: React.FC<OCRTaskComposerPanelProps> = ({
  onSubmit,
  onClose,
  embedded = false,
  draft,
  onDraftChange,
}) => {
  const [ocrMode, setOcrMode] = useState<OcrMode>(draft?.ocrMode ?? "image");
  const [files, setFiles] = useState<File[]>(draft?.files ?? []);
  const [recognizing, setRecognizing] = useState(draft?.recognizing ?? false);
  const [result, setResult] = useState<OcrResult | null>(draft?.result ?? null);
  const [selectedFields, setSelectedFields] = useState<string[]>(draft?.selectedFields ?? []);
  const [outputRange, setOutputRange] = useState(draft?.outputRange ?? "");
  const [dragOver, setDragOver] = useState(draft?.dragOver ?? false);
  const [writeStatus, setWriteStatus] = useState<string>("");
  const fileInputRef = useRef<HTMLInputElement>(null);

  const addOcrFiles = useCallback((newFiles: File[]) => {
    const acceptedFiles = newFiles.filter(isAcceptedOcrFile);
    if (acceptedFiles.length === 0) return;
    const nextMode: OcrMode = acceptedFiles.some(isLikelyInvoiceFile) ? "invoice" : ocrMode;
    const nextMaxFiles = nextMode === "invoice" ? 10 : 1;
    if (nextMode !== ocrMode) {
      setOcrMode(nextMode);
    }
    setFiles((prev) => [...prev, ...acceptedFiles].slice(0, nextMaxFiles));
    setResult(null);
    setWriteStatus("");
  }, [ocrMode]);

  useEffect(() => {
    onDraftChange?.({
      ocrMode,
      files,
      recognizing,
      result,
      selectedFields,
      outputRange,
      dragOver,
    });
  }, [
    ocrMode,
    files,
    recognizing,
    result,
    selectedFields,
    outputRange,
    dragOver,
    onDraftChange,
  ]);

  // ---- 文件处理 ----

  const handleFileChange = useCallback(
    (e: React.ChangeEvent<HTMLInputElement>) => {
      addOcrFiles(Array.from(e.target.files || []));
      // 重置 input 以便再次选择同一文件
      if (fileInputRef.current) fileInputRef.current.value = "";
    },
    [addOcrFiles]
  );

  const handleDrop = useCallback(
    (e: React.DragEvent) => {
      e.preventDefault();
      setDragOver(false);
      addOcrFiles(Array.from(e.dataTransfer.files));
    },
    [addOcrFiles]
  );

  useEffect(() => {
    const handleWindowPaste = (event: ClipboardEvent) => {
      const clipboardFiles = Array.from(event.clipboardData?.files || []);
      const itemFiles = Array.from(event.clipboardData?.items || [])
        .filter((item) => item.kind === "file")
        .map((item) => item.getAsFile())
        .filter(Boolean) as File[];
      const pastedFiles = [...clipboardFiles, ...itemFiles];
      if (!pastedFiles.some(isAcceptedOcrFile)) return;
      event.preventDefault();
      addOcrFiles(pastedFiles);
    };

    window.addEventListener("paste", handleWindowPaste);
    return () => window.removeEventListener("paste", handleWindowPaste);
  }, [addOcrFiles]);

  const removeFile = useCallback((index: number) => {
    setFiles((prev) => prev.filter((_, i) => i !== index));
    setWriteStatus("");
  }, []);

  const selectImageMode = useCallback(() => {
    setOcrMode("image");
    setFiles([]);
    setResult(null);
  }, []);

  const selectInvoiceMode = useCallback(() => {
    setOcrMode("invoice");
    setResult(null);
  }, []);

  // ---- 识别 ----

  const handleRecognize = useCallback(async () => {
    if (files.length === 0) return;
    setRecognizing(true);
    setResult(null);
    setSelectedFields([]);

    try {
      const effectiveOcrMode: OcrMode = ocrMode === "invoice" || files.some(isLikelyInvoiceFile)
        ? "invoice"
        : "image";
      const apiResult = await ipcApi.ocr.recognize(
        effectiveOcrMode,
        await resolveOcrFilePaths(files)
      );
      if (apiResult) {
        setResult(apiResult as OcrResult);
        const fieldNames = extractOcrFieldNames(apiResult as OcrResult);
        setSelectedFields(fieldNames);
        setRecognizing(false);
        return;
      }

      throw new Error("OCR 视觉识别通道不可用");
    } catch (err: any) {
      setResult({
        kind: ocrMode,
        text: "",
        rows: [],
        fields: {},
        invoices: [],
        errors: [err?.message || "识别失败"],
      });
    } finally {
      setRecognizing(false);
    }
  }, [files, ocrMode]);

  // ---- 字段选择 ----

  const toggleField = useCallback((field: string) => {
    setSelectedFields((prev) =>
      prev.includes(field) ? prev.filter((f) => f !== field) : [...prev, field]
    );
  }, []);

  const selectAllFields = useCallback(() => {
    if (!result) return;
    setSelectedFields(extractOcrFieldNames(result));
  }, [result]);

  const deselectAllFields = useCallback(() => {
    setSelectedFields([]);
  }, []);

  // ---- 写入工作表 ----

  const handleWriteToSheet = useCallback(async () => {
    if (!result) return;

    const values = buildOcrWriteValues(result, selectedFields);
    if (values.length === 0 || values.every((row) => row.length === 0)) {
      setWriteStatus("没有可写入的数据");
      return;
    }

    setWriteStatus("正在写入...");
    try {
      const target = await resolveWriteTarget(outputRange);
      if (!target.sheetName || !target.range) {
        throw new Error("未获取到目标单元格，请先在 Excel/WPS 中选中起始单元格");
      }
      const writeResult = await ipcApi.excel.writeRange(target.sheetName, target.range, values);
      if (!writeResult.success) {
        throw new Error(writeResult.error || "写入失败");
      }
      setOutputRange(`${target.sheetName}!${target.range}`);
      setWriteStatus(`已写入 ${values.length} 行 x ${Math.max(...values.map((row) => row.length))} 列`);
    } catch (err: any) {
      setWriteStatus(err?.message || "写入失败");
    }
  }, [result, selectedFields, outputRange]);

  const handlePickRange = useCallback(async () => {
    const range = await pickExcelRange();
    if (range) setOutputRange(range);
    else alert("未获取到选区，请确认已在 Excel/WPS 中选中了单元格");
  }, []);

  // ---- 渲染 ----

  const fieldNames = result ? extractOcrFieldNames(result) : [];
  const previewRows = result ? buildOcrPreviewRows(result, selectedFields).slice(0, 20) : [];
  const canWriteResult = result ? canWriteOcrResult(result, selectedFields) : false;

  return (
    <div className="task-composer-panel">
      {!embedded && (
        <div className="task-composer-title">
          <FileScan size={16} /> OCR / 发票识别
          <button
            className="task-close-btn"
            onClick={onClose}
          >
            <X size={14} />
          </button>
        </div>
      )}

      <OCRModeSelector
        ocrMode={ocrMode}
        onSelectImageMode={selectImageMode}
        onSelectInvoiceMode={selectInvoiceMode}
      />

      <OCRFileUploadSection
        ocrMode={ocrMode}
        files={files}
        dragOver={dragOver}
        fileInputRef={fileInputRef}
        onFileChange={handleFileChange}
        onDrop={handleDrop}
        onDragOverChange={setDragOver}
        onRemoveFile={removeFile}
      />

      {/* 识别按钮 */}
      <button
        className="task-submit-btn"
        disabled={files.length === 0 || recognizing}
        onClick={handleRecognize}
      >
        {recognizing
          ? (ocrMode === "invoice" ? "识别并提取字段中..." : "识别中...")
          : "开始识别"}
      </button>

      {result && (
        <OCRResultSection
          result={result}
          fieldNames={fieldNames}
          selectedFields={selectedFields}
          outputRange={outputRange}
          previewRows={previewRows}
          canWriteResult={canWriteResult}
          writeStatus={writeStatus}
          onSelectAllFields={selectAllFields}
          onDeselectAllFields={deselectAllFields}
          onToggleField={toggleField}
          onOutputRangeChange={setOutputRange}
          onPickRange={handlePickRange}
          onWriteToSheet={handleWriteToSheet}
        />
      )}
    </div>
  );
};
