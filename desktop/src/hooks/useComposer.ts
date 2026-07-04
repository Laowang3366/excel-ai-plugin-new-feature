/**
 * useComposer — 输入框状态与交互逻辑
 *
 * 从 ChatPage.tsx 提取，管理：
 * - 输入文本、附件、拖拽状态
 * - 附件选择（文件/图片/文件夹）
 * - 发送/恢复消息
 * - Popover 开关与外部点击关闭
 * - 文件夹文件列表弹窗开关
 */

import { useState, useCallback, useEffect, useRef } from "react";
import { useChatStore } from "../store/chatStore";
import { useSettingsStore } from "../store/settingsStore";
import { ipcApi } from "../services/ipcApi";
import { readFileAsBase64 } from "../utils/fileBase64";
import type { AttachedFile } from "../electronApi";

/** MIME 类型到文件扩展名映射 */
const MIME_TO_EXT: Record<string, string> = {
  "image/png": ".png",
  "image/jpeg": ".jpg",
  "image/webp": ".webp",
  "image/bmp": ".bmp",
  "image/gif": ".gif",
};

const IMAGE_EXTENSIONS = new Set([".png", ".jpg", ".jpeg", ".webp", ".bmp", ".gif"]);

function getFileExtension(file: File): string {
  const nameExt = file.name.includes(".") ? `.${file.name.split(".").pop()?.toLowerCase()}` : "";
  if (nameExt && nameExt !== ".undefined") return nameExt;
  return MIME_TO_EXT[file.type] || "";
}

function getAttachmentFileType(file: File): AttachedFile["fileType"] {
  const ext = getFileExtension(file);
  return IMAGE_EXTENSIONS.has(ext) || file.type.startsWith("image/") ? "image" : "document";
}

function getLocalPathForFile(file: File): string {
  const electronPath = (file as File & { path?: string }).path;
  if (electronPath) return electronPath;
  try {
    return ipcApi.file.getPathForFile?.(file) || "";
  } catch {
    return "";
  }
}

async function fileToTemporaryAttachment(file: File): Promise<AttachedFile | null> {
  try {
    const suffix = getFileExtension(file) || (file.type === "application/pdf" ? ".pdf" : ".bin");
    const base64 = await readFileAsBase64(file);
    const result = await ipcApi.file.writeTempFile({
      prefix: getAttachmentFileType(file) === "image" ? "image" : "attachment",
      suffix,
      data: base64,
    });
    if (!result.success || !result.filePath) return null;
    return {
      filePath: result.filePath,
      fileName: file.name || result.filePath.split(/[\\/]/).pop() || "attachment",
      fileType: getAttachmentFileType(file),
      size: file.size,
    };
  } catch {
    return null;
  }
}

export async function resolveDroppedFiles(files: File[]): Promise<AttachedFile[]> {
  const resolved: AttachedFile[] = [];
  for (const file of files) {
    const filePath = getLocalPathForFile(file);
    if (filePath) {
      resolved.push({
        filePath,
        fileName: file.name || filePath.split(/[\\/]/).pop() || filePath,
        fileType: getAttachmentFileType(file),
        size: file.size,
      });
      continue;
    }

    const temporaryAttachment = await fileToTemporaryAttachment(file);
    if (temporaryAttachment) {
      resolved.push(temporaryAttachment);
    }
  }
  return resolved;
}

interface ComposerDraft {
  inputText: string;
  attachedFiles: AttachedFile[];
}

const EMPTY_COMPOSER_DRAFT: ComposerDraft = {
  inputText: "",
  attachedFiles: [],
};

export function useComposer(draftKey = "new") {
  const {
    isStreaming,
    activeThreadId,
    turnStatus,
    lastInterruptContext,
    sendMessage,
    resumeFromInterruption,
    pendingComposerFiles,
    consumePendingFiles,
  } = useChatStore();

  const [inputText, setInputText] = useState("");
  const [showAttachPopover, setShowAttachPopover] = useState(false);
  const [showPermissionPopover, setShowPermissionPopover] = useState(false);
  const [showThinkingPopover, setShowThinkingPopover] = useState(false);
  const [attachedFiles, setAttachedFiles] = useState<AttachedFile[]>([]);
  const [composerDragOver, setComposerDragOver] = useState(false);
  const [showFolderFileList, setShowFolderFileList] = useState(false);
  const [showComposerFolderList, setShowComposerFolderList] = useState(false);
  const [composerDrafts, setComposerDrafts] = useState<Record<string, ComposerDraft>>({});
  const textareaRef = useRef<HTMLTextAreaElement>(null);
  const draftKeyRef = useRef(draftKey);
  const inputTextRef = useRef(inputText);
  const attachedFilesRef = useRef(attachedFiles);
  const composerDraftsRef = useRef(composerDrafts);

  useEffect(() => {
    inputTextRef.current = inputText;
    attachedFilesRef.current = attachedFiles;
  }, [inputText, attachedFiles]);

  useEffect(() => {
    composerDraftsRef.current = composerDrafts;
  }, [composerDrafts]);

  useEffect(() => {
    if (draftKeyRef.current === draftKey) return;

    const previousKey = draftKeyRef.current;
    const previousDraft = {
      inputText: inputTextRef.current,
      attachedFiles: attachedFilesRef.current,
    };
    const nextDraft = composerDraftsRef.current[draftKey] ?? EMPTY_COMPOSER_DRAFT;

    setComposerDrafts((prev) => ({
      ...prev,
      [previousKey]: previousDraft,
    }));
    draftKeyRef.current = draftKey;
    setInputText(nextDraft.inputText);
    setAttachedFiles(nextDraft.attachedFiles);
    setShowAttachPopover(false);
    setShowPermissionPopover(false);
    setShowThinkingPopover(false);
    setComposerDragOver(false);
  }, [draftKey]);

  useEffect(() => {
    setComposerDrafts((prev) => ({
      ...prev,
      [draftKeyRef.current]: {
        inputText,
        attachedFiles,
      },
    }));
  }, [inputText, attachedFiles]);

  // 发送消息
  const handleSend = useCallback(() => {
    const trimmedInput = inputText.trim();
    if (!trimmedInput && attachedFiles.length === 0) return;

    const content = trimmedInput;
    const attachments = attachedFiles.length > 0 ? attachedFiles : undefined;

    if (turnStatus === "interrupted" && lastInterruptContext) {
      resumeFromInterruption(content, attachments);
    } else {
      sendMessage(content, attachments);
    }
    if (isStreaming && !activeThreadId) {
      return;
    }
    setInputText("");
    setAttachedFiles([]);
  }, [inputText, attachedFiles, isStreaming, activeThreadId, turnStatus, lastInterruptContext, sendMessage, resumeFromInterruption]);

  // 文件选择
  const handleOpenFile = useCallback(async () => {
    setShowAttachPopover(false);
    try {
      const result = await ipcApi.dialog.openFile();
      if (!result.canceled && result.filePaths.length > 0) {
        const newFiles: AttachedFile[] = result.filePaths.map((fp) => ({
          filePath: fp,
          fileName: fp.split(/[\\/]/).pop() || fp,
          fileType: "document" as const,
        }));
        setAttachedFiles((prev) => [...prev, ...newFiles]);
      }
    } catch { /* 静默失败 */ }
  }, []);

  const handleOpenImage = useCallback(async () => {
    setShowAttachPopover(false);
    try {
      const result = await ipcApi.dialog.openImage();
      if (!result.canceled && result.filePaths.length > 0) {
        const newFiles: AttachedFile[] = result.filePaths.map((fp) => ({
          filePath: fp,
          fileName: fp.split(/[\\/]/).pop() || fp,
          fileType: "image" as const,
        }));
        setAttachedFiles((prev) => [...prev, ...newFiles]);
      }
    } catch { /* 静默失败 */ }
  }, []);

  // 添加文件夹：选择文件夹后，将可操作的 Office 文件加入附件 + 固定到侧边栏
  const handleOpenFolder = useCallback(async () => {
    setShowAttachPopover(false);
    try {
      const result = await ipcApi.dialog.openFolder();
      if (!result.canceled && result.filePaths.length > 0) {
        const folderPath = result.filePaths[0];
        const files = await ipcApi.folder.listFiles(folderPath);
        if (files.length > 0) {
          const newFiles: AttachedFile[] = files.map((f) => ({
            filePath: f.filePath,
            fileName: f.fileName,
            fileType: "document" as const,
            size: f.size,
          }));
          setAttachedFiles((prev) => {
            const existing = new Set(prev.map((f2) => f2.filePath));
            const deduped = newFiles.filter((f2) => !existing.has(f2.filePath));
            return [...prev, ...deduped];
          });
        }
        // 同时固定到侧边栏
        const { addPinnedFolder } = useSettingsStore.getState();
        const folderName = folderPath.split(/[\\/]/).pop() || folderPath;
        addPinnedFolder({ path: folderPath, name: folderName, addedAt: Date.now() });
      }
    } catch { /* 静默失败 */ }
  }, []);

  const removeAttachedFile = useCallback((index: number) => {
    setAttachedFiles((prev) => prev.filter((_, i) => i !== index));
  }, []);

  // 消费从侧边栏文件夹推入的待添加文件
  useEffect(() => {
    if (pendingComposerFiles.length > 0) {
      const files = consumePendingFiles();
      if (files.length > 0) {
        setAttachedFiles((prev) => {
          const existing = new Set(prev.map((f) => f.filePath));
          const newFiles = files.filter((f) => !existing.has(f.filePath));
          return [...prev, ...newFiles];
        });
      }
    }
  }, [pendingComposerFiles, consumePendingFiles]);

  // 拖拽文件到输入框
  const handleComposerDragOver = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    if (e.dataTransfer.types.includes("Files")) {
      setComposerDragOver(true);
    }
  }, []);

  const handleComposerDragLeave = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    const rect = (e.currentTarget as HTMLElement).getBoundingClientRect();
    if (
      e.clientX <= rect.left || e.clientX >= rect.right ||
      e.clientY <= rect.top || e.clientY >= rect.bottom
    ) {
      setComposerDragOver(false);
    }
  }, []);

  const handleComposerDrop = useCallback(async (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    setComposerDragOver(false);

    const files = Array.from(e.dataTransfer.files);
    if (files.length === 0) return;

    const newFiles = await resolveDroppedFiles(files);

    if (newFiles.length > 0) {
      setAttachedFiles((prev) => {
        const existing = new Set(prev.map((f) => f.filePath));
        return [...prev, ...newFiles.filter((f) => !existing.has(f.filePath))];
      });
    }
  }, []);

  // 处理键盘事件
  const handleKeyDown = useCallback((e: React.KeyboardEvent) => {
    if (e.key === "Enter" && !e.shiftKey) {
      e.preventDefault();
      handleSend();
    }
  }, [handleSend]);

  // textarea 自动高度
  const handleTextareaChange = useCallback((e: React.ChangeEvent<HTMLTextAreaElement>) => {
    setInputText(e.target.value);
    const el = e.target;
    el.style.height = "auto";
    el.style.height = Math.min(el.scrollHeight, 220) + "px";
  }, []);

  // 粘贴文件/图片
  const handlePaste = useCallback(async (e: React.ClipboardEvent) => {
    const items = Array.from(e.clipboardData.items);
    const files = items
      .filter(item => item.kind === "file")
      .map(item => item.getAsFile())
      .filter(Boolean) as File[];
    if (files.length === 0) return;

    const newFiles: AttachedFile[] = [];

    for (const f of files) {
      const filePath = getLocalPathForFile(f);

      if (filePath) {
        // 文件系统的文件 → 直接使用路径
        newFiles.push({
          filePath,
          fileName: f.name || filePath.split(/[\\/]/).pop() || filePath,
          fileType: getAttachmentFileType(f),
          size: f.size,
        });
      } else {
        // 截图或虚拟文件 → 读取为 base64，写入临时文件
        const temporaryAttachment = await fileToTemporaryAttachment(f);
        if (temporaryAttachment) newFiles.push(temporaryAttachment);
      }
    }

    if (newFiles.length > 0) {
      e.preventDefault();
      setAttachedFiles((prev) => {
        const existing = new Set(prev.map((f) => f.filePath));
        return [...prev, ...newFiles.filter((f) => !existing.has(f.filePath))];
      });
    }
  }, []);

  // 点击外部关闭 popover
  useEffect(() => {
    if (!showAttachPopover && !showPermissionPopover && !showThinkingPopover) return;
    const handleClickOutside = (e: MouseEvent) => {
      const target = e.target as HTMLElement;
      if (target.closest(".composer-popover-wrapper")) return;
      setShowAttachPopover(false);
      setShowPermissionPopover(false);
      setShowThinkingPopover(false);
    };
    document.addEventListener("click", handleClickOutside);
    return () => {
      document.removeEventListener("click", handleClickOutside);
    };
  }, [showAttachPopover, showPermissionPopover, showThinkingPopover]);

  // 点击外部关闭文件夹文件列表弹窗
  useEffect(() => {
    if (!showFolderFileList && !showComposerFolderList) return;
    const handleClickOutside = (e: MouseEvent) => {
      const target = e.target as HTMLElement;
      if (target.closest(".chat-folder-badge-wrapper") || target.closest(".composer-folder-context-wrapper") || target.closest(".folder-file-popover")) return;
      setShowFolderFileList(false);
      setShowComposerFolderList(false);
    };
    document.addEventListener("click", handleClickOutside);
    return () => {
      document.removeEventListener("click", handleClickOutside);
    };
  }, [showFolderFileList, showComposerFolderList]);

  const hasInput = inputText.trim().length > 0 || attachedFiles.length > 0;

  return {
    // 状态
    inputText,
    setInputText,
    attachedFiles,
    composerDragOver,
    showAttachPopover,
    setShowAttachPopover,
    showPermissionPopover,
    setShowPermissionPopover,
    showThinkingPopover,
    setShowThinkingPopover,
    showFolderFileList,
    setShowFolderFileList,
    showComposerFolderList,
    setShowComposerFolderList,
    textareaRef,
    hasInput,
    isStreaming,

    // 回调
    handleSend,
    handleOpenFile,
    handleOpenImage,
    handleOpenFolder,
    removeAttachedFile,
    handleComposerDragOver,
    handleComposerDragLeave,
    handleComposerDrop,
    handleKeyDown,
    handleTextareaChange,
    handlePaste,
  };
}
