import { useCallback, useMemo, useState } from "react";
import type { AgentContentPart } from "@shared/agent";
import {
  buildOcrTaskPayload,
  buildOcrWriteValues,
  canWriteOcrResult,
  extractOcrFieldNames,
  isAcceptedOcrFile,
  isImageOcrFile,
  isLikelyInvoiceFile,
  isPdfOcrFile,
  mimeTypeForFile,
  parseOcrAssistantResult,
  parseSheetRangeAddress,
  readFileAsBase64,
  sanitizeOcrUiText,
  type OcrResult,
} from "@shared/tasks";
import { ChatApprovalCard } from "../ChatApprovalCard";
import { useTaskSubmit } from "./useTaskSubmit";
import { useSelectionAddress } from "./useSelectionAddress";
import { TaskSubmitStatus } from "./TaskSubmitStatus";

export function OcrTaskPanel() {
  const {
    submit,
    runTool,
    busy,
    error,
    lastResult,
    setError,
    setLastResult,
    adapter,
    view,
    approve,
    reject,
  } = useTaskSubmit();
  const { readSelection, busy: picking, error: pickError } =
    useSelectionAddress(adapter);
  const [mode, setMode] = useState<"image" | "invoice">("image");
  const [files, setFiles] = useState<File[]>([]);
  const [task, setTask] = useState("");
  const [outputRange, setOutputRange] = useState("");
  const [localNote, setLocalNote] = useState<string | undefined>();
  const [rawPreview, setRawPreview] = useState<string | undefined>();
  const [structured, setStructured] = useState<OcrResult | null>(null);
  const [parseNote, setParseNote] = useState<string | undefined>();
  const [selectedFields, setSelectedFields] = useState<string[]>([]);
  const [writing, setWriting] = useState(false);

  const fieldNames = useMemo(
    () => (structured ? extractOcrFieldNames(structured) : []),
    [structured],
  );

  const addFiles = useCallback((list: FileList | File[]) => {
    const arr = Array.from(list).filter(isAcceptedOcrFile);
    if (arr.length === 0) {
      setLocalNote("仅支持图片或 PDF 文件");
      return;
    }
    const nextMode: "image" | "invoice" = arr.some(isLikelyInvoiceFile)
      ? "invoice"
      : mode;
    if (nextMode !== mode) setMode(nextMode);
    const max = nextMode === "invoice" ? 10 : 1;
    setFiles((prev) => [...prev, ...arr].slice(0, max));
    setLocalNote(undefined);
  }, [mode]);

  const applyAssistantText = (assistantText: string | undefined) => {
    const text = assistantText ?? "";
    const parsed = parseOcrAssistantResult(text);
    if (parsed.ok) {
      setStructured(parsed.result);
      setRawPreview(
        sanitizeOcrUiText(
          [parsed.narrative, parsed.result.text].filter(Boolean).join("\n\n"),
        ),
      );
      const names = extractOcrFieldNames(parsed.result);
      setSelectedFields(names);
      setParseNote(
        parsed.narrative
          ? "已解析结构化 OCR 结果"
          : "已解析结构化 OCR 结果（无额外说明）",
      );
    } else {
      setStructured(null);
      setSelectedFields([]);
      setRawPreview(sanitizeOcrUiText(parsed.rawText));
      setParseNote(parsed.reason);
    }
  };

  const onRecognize = async () => {
    setLocalNote(undefined);
    setError(undefined);
    setParseNote(undefined);
    setStructured(null);
    setRawPreview(undefined);
    setSelectedFields([]);
    if (files.length === 0) {
      setLocalNote("请先上传图片");
      return;
    }
    const pdfs = files.filter(isPdfOcrFile);
    const images = files.filter(isImageOcrFile);
    if (pdfs.length > 0 && images.length === 0) {
      setLocalNote(
        "PDF 在浏览器加载项中无可靠解析通道（无 ocr.parseDocument / 本地引擎）：typed unsupported。请上传图片或改用桌面端。",
      );
      return;
    }
    if (pdfs.length > 0) {
      setLocalNote("已忽略 PDF（unsupported）；仅将图片送入多模态识别。");
    }
    if (images.length === 0) {
      setLocalNote("没有可识别的图片");
      return;
    }

    const contentParts: AgentContentPart[] = [];
    for (const file of images) {
      const base64 = await readFileAsBase64(file);
      contentParts.push({
        type: "image",
        mimeType: mimeTypeForFile(file),
        base64,
        fileName: file.name,
      });
    }

    const payload = buildOcrTaskPayload({
      mode,
      fileNames: images.map((f) => f.name),
      task,
      outputRange,
    });
    // Never put base64/API keys into UI result strings — only file names in payload text.
    const outcome = await submit(payload, { contentParts });
    if (outcome.accepted) {
      applyAssistantText(outcome.assistantText);
    }
  };

  const toggleField = (name: string) => {
    setSelectedFields((prev) =>
      prev.includes(name) ? prev.filter((x) => x !== name) : [...prev, name],
    );
  };

  const resolveWriteTarget = async (): Promise<
    { sheetName: string; range: string } | null
  > => {
    const direct = parseSheetRangeAddress(outputRange);
    if (direct) return direct;
    if (outputRange.trim()) {
      // bare A1 — try current selection sheet
      if (!adapter) {
        setError("宿主未就绪");
        return null;
      }
      const sel = await adapter.getSelection();
      if (!sel.ok) {
        setError(sel.reason || "无法读取选区以补全工作表名");
        return null;
      }
      const bare = outputRange.trim();
      if (/!/.test(bare)) {
        setError("写入地址格式无效，请使用 工作表!A1");
        return null;
      }
      return { sheetName: sel.data.sheetName, range: bare };
    }
    setError("请填写写入锚点或点击「读取选区」");
    return null;
  };

  const writeValues = async (values: string[][], label: string) => {
    if (!values.length) {
      setError("没有可写入的内容");
      return;
    }
    const target = await resolveWriteTarget();
    if (!target) return;
    setWriting(true);
    setError(undefined);
    try {
      const result = await runTool(
        "range.write",
        {
          sheetName: target.sheetName,
          range: target.range,
          values,
          verify: true,
        },
        { toolCallId: `ocr-write-${label}` },
      );
      if (result.ok) {
        setLastResult(`已写入 ${target.sheetName}!${target.range}（${label}）`);
      }
    } finally {
      setWriting(false);
    }
  };

  const onWriteFullText = async () => {
    if (!structured) {
      if (rawPreview?.trim()) {
        await writeValues([[rawPreview]], "整段文本");
        return;
      }
      setError("没有可写入的识别文本");
      return;
    }
    const text = structured.text.trim() || rawPreview?.trim() || "";
    if (!text) {
      setError("整段文本为空");
      return;
    }
    await writeValues([[text]], "整段文本");
  };

  const onWriteSelectedFields = async () => {
    if (!structured) {
      setError("无结构化结果，无法按字段写入；请使用整段文本写入");
      return;
    }
    if (!canWriteOcrResult(structured, selectedFields)) {
      setError("请至少勾选一个字段，或改用整段文本写入");
      return;
    }
    const values = buildOcrWriteValues(structured, selectedFields);
    if (!values.length) {
      setError("所选字段无内容");
      return;
    }
    await writeValues(values, "所选字段");
  };

  const uiBusy = busy || writing;

  return (
    <section className="card task-panel">
      <h2>OCR 识别</h2>
      <p className="muted">
        浏览器上传图片，经当前模型供应商多模态识别；结果可预览，写入经审批边界。无
        API Key / Base64 展示。
      </p>
      <div className="task-field">
        <label className="task-check">
          <input
            type="radio"
            name="ocr-mode"
            checked={mode === "image"}
            onChange={() => setMode("image")}
          />
          通用图片（单文件）
        </label>
        <label className="task-check">
          <input
            type="radio"
            name="ocr-mode"
            checked={mode === "invoice"}
            onChange={() => setMode("invoice")}
          />
          发票（最多 10 张图）
        </label>
      </div>
      <div className="task-field">
        <label>
          上传文件
          <input
            type="file"
            accept="image/*,.pdf,application/pdf"
            multiple={mode === "invoice"}
            onChange={(e) => {
              if (e.target.files) addFiles(e.target.files);
              e.target.value = "";
            }}
          />
        </label>
        {files.length > 0 && (
          <ul className="task-file-list">
            {files.map((f) => (
              <li key={`${f.name}-${f.size}`}>
                {f.name}
                {isPdfOcrFile(f) ? "（PDF：可能 unsupported）" : ""}
                <button
                  type="button"
                  className="task-chip"
                  onClick={() =>
                    setFiles((prev) => prev.filter((x) => x !== f))
                  }
                >
                  移除
                </button>
              </li>
            ))}
          </ul>
        )}
      </div>
      <div className="task-field">
        <label>
          写入锚点（写入前必填；可用读取选区）
          <div className="row">
            <input
              value={outputRange}
              onChange={(e) => setOutputRange(e.target.value)}
              placeholder="Sheet1!A1"
              aria-label="OCR 写入锚点"
            />
            <button
              type="button"
              disabled={picking || !adapter}
              onClick={async () => {
                const addr = await readSelection();
                if (addr) setOutputRange(addr);
              }}
            >
              读取选区
            </button>
          </div>
        </label>
      </div>
      <div className="task-field">
        <label>
          需求说明
          <textarea
            rows={2}
            value={task}
            onChange={(e) => setTask(e.target.value)}
            placeholder="抽取字段 / 原样转录…"
          />
        </label>
      </div>
      <button
        type="button"
        disabled={uiBusy || !adapter || !view.canSend}
        onClick={() => void onRecognize()}
      >
        识别并提交到 AI
      </button>

      {(rawPreview != null || structured) && (
        <div className="task-field ocr-preview" data-testid="ocr-preview">
          <h3>识别结果预览</h3>
          {parseNote && <p className="muted">{sanitizeOcrUiText(parseNote)}</p>}
          {rawPreview != null && (
            <div className="task-field">
              <label>
                整段文本
                <pre className="ocr-preview-text" aria-label="OCR 整段文本">
                  {rawPreview || "（空）"}
                </pre>
              </label>
            </div>
          )}
          {structured && fieldNames.length > 0 && (
            <div className="task-field">
              <span>可选字段</span>
              <ul className="task-chip-list">
                {fieldNames.map((name) => (
                  <li key={name}>
                    <label className="task-check">
                      <input
                        type="checkbox"
                        checked={selectedFields.includes(name)}
                        onChange={() => toggleField(name)}
                      />
                      {name}
                    </label>
                  </li>
                ))}
              </ul>
            </div>
          )}
          {structured && structured.rows.length > 0 && (
            <div className="task-field">
              <span>表格预览</span>
              <div className="ocr-table-wrap">
                <table className="ocr-preview-table">
                  <tbody>
                    {structured.rows.slice(0, 12).map((row, i) => (
                      <tr key={i}>
                        {row.slice(0, 8).map((cell, j) => (
                          <td key={j}>{sanitizeOcrUiText(cell)}</td>
                        ))}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          )}
          {structured &&
            structured.invoices.length > 0 &&
            structured.invoices.map((inv, idx) => (
              <div className="task-field" key={`${inv.filename}-${idx}`}>
                <span>
                  发票 {idx + 1}
                  {inv.filename ? ` · ${inv.filename}` : ""}
                </span>
                <pre className="ocr-preview-text">
                  {sanitizeOcrUiText(
                    Object.entries(inv.fields)
                      .map(([k, v]) => `${k}: ${v}`)
                      .join("\n") || inv.text || "（无字段）",
                  )}
                </pre>
              </div>
            ))}
          <div className="row chat-actions">
            <button
              type="button"
              disabled={uiBusy || !adapter}
              onClick={() => void onWriteFullText()}
            >
              写入整段文本
            </button>
            <button
              type="button"
              disabled={
                uiBusy ||
                !adapter ||
                !structured ||
                selectedFields.length === 0
              }
              onClick={() => void onWriteSelectedFields()}
            >
              写入所选字段
            </button>
          </div>
        </div>
      )}

      {view.pendingApproval && (
        <ChatApprovalCard
          request={view.pendingApproval}
          disabled={!view.canApprove || view.status === "stopping"}
          onApprove={(id) => {
            approve(id);
          }}
          onReject={(id) => {
            reject(id);
          }}
        />
      )}

      <TaskSubmitStatus
        busy={uiBusy}
        error={error || pickError || localNote}
        lastResult={lastResult}
      />
    </section>
  );
}
