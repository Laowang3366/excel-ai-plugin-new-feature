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
import { FileScan, Paperclip, FileText, Image, AlertTriangle, Ruler, X } from "../common/IconMap";
import { ipcApi } from "../../services/ipcApi";
import { pickExcelRange } from "../../utils/chatHelpers";

export type OcrMode = "image" | "invoice";

/** 单张发票识别结果 */
export interface OcrInvoiceItem {
  filename: string;
  text: string;
  fields: Record<string, string>;
  rows: string[][];
  error?: string;
}

/** OCR 识别结果 */
export interface OcrResult {
  kind: "image" | "invoice";
  text: string;
  rows: string[][];
  fields: Record<string, string>;
  invoices: OcrInvoiceItem[];
  errors: string[];
}

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

const ACCEPTED_TYPES = "image/png, image/jpeg, image/webp, image/bmp, application/pdf";
const ACCEPTED_OCR_MIME_TYPES = new Set(ACCEPTED_TYPES.split(", "));

function readFileAsBase64(file: File): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => {
      const result = reader.result as string;
      const commaIndex = result.indexOf(",");
      resolve(commaIndex >= 0 ? result.slice(commaIndex + 1) : result);
    };
    reader.onerror = () => reject(reader.error);
    reader.readAsDataURL(file);
  });
}

async function resolveOcrFilePaths(files: File[]): Promise<string[]> {
  const paths: string[] = [];
  for (const file of files) {
    const existingPath = (file as File & { path?: string }).path;
    if (existingPath) {
      paths.push(existingPath);
      continue;
    }

    const suffix = file.name.includes(".")
      ? `.${file.name.split(".").pop()}`
      : file.type === "application/pdf"
      ? ".pdf"
      : ".png";
    const result = await ipcApi.file.writeTempFile({
      prefix: "ocr",
      suffix,
      data: await readFileAsBase64(file),
    });
    if (!result.success || !result.filePath) {
      throw new Error(result.error || `无法读取 OCR 文件: ${file.name}`);
    }
    paths.push(result.filePath);
  }
  return paths;
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

      {/* 识别模式 */}
      <div className="task-field">
        <label className="task-field-label">识别模式</label>
        <div className="task-select-group">
          <button
            className={`task-select-btn ${ocrMode === "image" ? "active" : ""}`}
            onClick={() => {
              setOcrMode("image");
              setFiles([]);
              setResult(null);
            }}
          >
            通用 OCR
          </button>
          <button
            className={`task-select-btn ${ocrMode === "invoice" ? "active" : ""}`}
            onClick={() => {
              setOcrMode("invoice");
              setResult(null);
            }}
          >
            发票识别
          </button>
        </div>
      </div>

      {/* 文件上传 */}
      <div className="task-field">
        <label className="task-field-label">
          上传文件（{ocrMode === "invoice" ? "最多 10 个" : "1 个"}）
        </label>
        <div
          className={`ocr-drop-zone ${dragOver ? "drag-over" : ""}`}
          onDragOver={(e) => {
            e.preventDefault();
            setDragOver(true);
          }}
          onDragLeave={() => setDragOver(false)}
          onDrop={handleDrop}
          onClick={() => fileInputRef.current?.click()}
        >
          <input
            ref={fileInputRef}
            type="file"
            accept={ACCEPTED_TYPES}
            multiple={ocrMode === "invoice"}
            onChange={handleFileChange}
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

        {/* 已选文件列表 */}
        {files.length > 0 && (
          <div className="ocr-file-list">
            {files.map((f, i) => (
              <div key={i} className="ocr-file-item">
                <span className="ocr-file-icon">
                  {f.type === "application/pdf" ? <FileText size={14} /> : <Image size={14} />}
                </span>
                <span className="ocr-file-name">{f.name}</span>
                <button
                  className="ocr-file-remove"
                  onClick={() => removeFile(i)}
                >
                  <X size={10} />
                </button>
              </div>
            ))}
          </div>
        )}
      </div>

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

      {/* 识别结果 */}
      {result && (
        <>
          {result.errors.length > 0 && (
            <div className="ocr-errors">
              {result.errors.map((err, i) => (
                <p key={i} style={{ color: "var(--danger)", fontSize: 12 }}>
                  <AlertTriangle size={12} style={{ verticalAlign: "middle", marginRight: 4 }} /> {err}
                </p>
              ))}
            </div>
          )}

          {/* 字段选择 */}
          {fieldNames.length > 0 && (
            <div className="task-field">
              <div className="task-field-label-row">
                <label className="task-field-label">选择字段</label>
                <div className="task-field-actions">
                  <button className="task-link-btn" onClick={selectAllFields}>
                    全选
                  </button>
                  <button className="task-link-btn" onClick={deselectAllFields}>
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
                      onChange={() => toggleField(field)}
                    />
                    <span>{field}</span>
                  </label>
                ))}
              </div>
            </div>
          )}

          {/* 目标单元格 */}
          <div className="task-field">
            <label className="task-field-label">填入起始位置</label>
            <div className="range-input-row">
              <input
                className="task-field-input"
                value={outputRange}
                onChange={(e) => setOutputRange(e.target.value)}
                placeholder="如 Sheet1!A1"
              />
              <button
                className="btn-pick-range"
                onClick={async () => {
                  const range = await pickExcelRange();
                  if (range) setOutputRange(range);
                  else alert("未获取到选区，请确认已在 Excel/WPS 中选中了单元格");
                }}
              >
                <Ruler size={13} /> 选区
              </button>
            </div>
          </div>

          {/* 预览表格 */}
          {previewRows.length > 0 && (
            <div className="task-field">
              <label className="task-field-label">
                预览（前 {Math.min(previewRows.length, 20)} 行）
              </label>
              <div className="ocr-preview-table-wrapper">
                <table className="ocr-preview-table">
                  <thead>
                    <tr>
                      {previewRows[0]?.map((cell, i) => (
                        <th key={i}>{cell}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {previewRows.slice(1).map((row, ri) => (
                      <tr key={ri}>
                        {row.map((cell, ci) => (
                          <td key={ci}>{cell}</td>
                        ))}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          )}

          {/* 写入按钮 */}
          <button
            className="task-submit-btn"
            disabled={!canWriteResult}
            onClick={handleWriteToSheet}
            style={{ marginTop: 8 }}
          >
            写入单元格
          </button>
          {writeStatus && (
            <p style={{ marginTop: 8, fontSize: 12, color: "var(--text-muted)" }}>
              {writeStatus}
            </p>
          )}
        </>
      )}
    </div>
  );
};

async function resolveWriteTarget(input: string): Promise<{ sheetName: string; range: string }> {
  const trimmed = input.trim();
  if (trimmed) {
    const parsed = parseSheetRange(trimmed);
    if (parsed.sheetName && parsed.range) return parsed;
    const selection = await ipcApi.excel.getSelectionAddress();
    return { sheetName: selection.sheetName, range: parsed.range || trimmed };
  }

  const selection = await ipcApi.excel.getSelectionAddress();
  return { sheetName: selection.sheetName, range: selection.address };
}

function parseSheetRange(value: string): { sheetName: string; range: string } {
  const bangIndex = value.lastIndexOf("!");
  if (bangIndex < 0) return { sheetName: "", range: value };
  const rawSheetName = value.slice(0, bangIndex).trim();
  return {
    sheetName: rawSheetName.replace(/^'(.*)'$/, "$1"),
    range: value.slice(bangIndex + 1).trim(),
  };
}

export function extractOcrFieldNames(result: OcrResult): string[] {
  if (result.kind === "invoice" && result.invoices.length > 0) {
    return Array.from(new Set(result.invoices.flatMap((invoice) => Object.keys(invoice.fields || {}))));
  }
  if (result.rows.length > 0) return result.rows[0];
  if (result.fields && Object.keys(result.fields).length > 0) return Object.keys(result.fields);
  if (result.invoices.length > 0) {
    return Array.from(new Set(result.invoices.flatMap((invoice) => Object.keys(invoice.fields || {}))));
  }
  return [];
}

export function canWriteOcrResult(result: OcrResult, selectedFields: string[]): boolean {
  return selectedFields.length > 0 ||
    (extractOcrFieldNames(result).length === 0 && Boolean(result.text.trim()));
}

export function buildOcrWriteValues(result: OcrResult, selectedFields: string[]): string[][] {
  if (selectedFields.length === 0) {
    return result.text ? [[result.text]] : [];
  }

  if (result.kind === "invoice" && result.invoices.length > 0) {
    const rows = result.invoices.map((invoice) =>
      selectedFields.map((field) => stringifyCell(invoice.fields[field] ?? ""))
    );
    return [selectedFields, ...rows];
  }

  if (result.rows.length > 0) {
    return filterRowsForFields(result.rows, selectedFields);
  }

  if (Object.keys(result.fields).length > 0) {
    return [
      selectedFields,
      selectedFields.map((field) => stringifyCell(result.fields[field] ?? "")),
    ];
  }

  return [[result.text || ""]];
}

export function buildOcrPreviewRows(result: OcrResult, selectedFields: string[]): string[][] {
  const fields = selectedFields.length > 0 ? selectedFields : extractOcrFieldNames(result);
  const values = buildOcrWriteValues(result, fields);
  if (values.length > 0) return values;
  return result.rows.length > 0 ? filterRowsForFields(result.rows, fields) : [];
}

function stringifyCell(value: unknown): string {
  return value === null || value === undefined ? "" : String(value);
}

function isAcceptedOcrFile(file: File): boolean {
  const name = file.name.toLowerCase();
  return ACCEPTED_OCR_MIME_TYPES.has(file.type) ||
    name.endsWith(".pdf") ||
    name.endsWith(".png") ||
    name.endsWith(".jpg") ||
    name.endsWith(".jpeg") ||
    name.endsWith(".webp") ||
    name.endsWith(".bmp");
}

function isLikelyInvoiceFile(file: File): boolean {
  return /发票|invoice|fapiao|票据/i.test(file.name);
}

/** 根据选中字段过滤行 */
function filterRowsForFields(
  rows: string[][],
  selectedFields: string[]
): string[][] {
  if (rows.length === 0 || selectedFields.length === 0) return rows;
  const header = rows[0];
  const colIndices = selectedFields
    .map((f) => header.indexOf(f))
    .filter((i) => i >= 0);
  if (colIndices.length === 0) return rows;
  return rows.map((row) => colIndices.map((i) => row[i] ?? ""));
}
