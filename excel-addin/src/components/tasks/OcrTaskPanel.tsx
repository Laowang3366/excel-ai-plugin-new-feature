import { useCallback, useState } from "react";
import type { AgentContentPart } from "@shared/agent";
import {
  buildOcrTaskPayload,
  isAcceptedOcrFile,
  isImageOcrFile,
  isLikelyInvoiceFile,
  isPdfOcrFile,
  mimeTypeForFile,
  readFileAsBase64,
} from "@shared/tasks";
import { useTaskSubmit } from "./useTaskSubmit";
import { useSelectionAddress } from "./useSelectionAddress";
import { TaskSubmitStatus } from "./TaskSubmitStatus";

export function OcrTaskPanel() {
  const { submit, busy, error, lastResult, setError, adapter, view } =
    useTaskSubmit();
  const { readSelection, busy: picking, error: pickError } =
    useSelectionAddress(adapter);
  const [mode, setMode] = useState<"image" | "invoice">("image");
  const [files, setFiles] = useState<File[]>([]);
  const [task, setTask] = useState("");
  const [outputRange, setOutputRange] = useState("");
  const [localNote, setLocalNote] = useState<string | undefined>();

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

  const onRecognize = async () => {
    setLocalNote(undefined);
    setError(undefined);
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
    await submit(payload, { contentParts });
  };

  return (
    <section className="card task-panel">
      <h2>OCR 识别</h2>
      <p className="muted">
        浏览器上传图片，经当前模型供应商多模态识别；写入工作表走 host 与审批。无 API Key 展示。
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
          写入锚点（可选，识别后可由 Agent range.write）
          <div className="row">
            <input
              value={outputRange}
              onChange={(e) => setOutputRange(e.target.value)}
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
        disabled={busy || !adapter || !view.canSend}
        onClick={() => void onRecognize()}
      >
        识别并提交到 AI
      </button>
      <TaskSubmitStatus
        busy={busy}
        error={error || pickError || localNote}
        lastResult={lastResult}
      />
    </section>
  );
}
