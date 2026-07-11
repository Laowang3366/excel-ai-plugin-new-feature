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
import { resolveDroppedFiles } from "./composerAttachmentFiles";
import { useDocumentDismiss } from "./useDocumentDismiss";
import type { FileAttachment } from "../electronApi";

export { resolveDroppedFiles } from "./composerAttachmentFiles";

const COMPOSER_POPOVER_IGNORE_SELECTORS = [".composer-popover-wrapper"];
const COMPOSER_FOLDER_IGNORE_SELECTORS = [
  ".chat-folder-badge-wrapper",
  ".composer-folder-context-wrapper",
  ".folder-file-popover",
];

interface ComposerDraft {
  inputText: string;
  attachedFiles: FileAttachment[];
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
  const [attachedFiles, setAttachedFiles] = useState<FileAttachment[]>([]);
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
    if (!trimmedInput && attachedFiles.length === 0) return Promise.resolve(null);

    const content = trimmedInput;
    const attachments = attachedFiles.length > 0 ? attachedFiles : undefined;

    const sendPromise =
      turnStatus === "interrupted" && lastInterruptContext
        ? resumeFromInterruption(content, attachments).then(() => activeThreadId)
        : sendMessage(content, attachments);
    if (isStreaming && !activeThreadId) {
      return sendPromise;
    }
    setInputText("");
    setAttachedFiles([]);
    return sendPromise;
  }, [
    inputText,
    attachedFiles,
    isStreaming,
    activeThreadId,
    turnStatus,
    lastInterruptContext,
    sendMessage,
    resumeFromInterruption,
  ]);

  // 文件选择
  const handleOpenFile = useCallback(async () => {
    setShowAttachPopover(false);
    try {
      const result = await ipcApi.dialog.openFile();
      if (!result.canceled && result.filePaths.length > 0) {
        const newFiles: FileAttachment[] = result.filePaths.map((fp) => ({
          filePath: fp,
          fileName: fp.split(/[\\/]/).pop() || fp,
          fileType: "document" as const,
        }));
        setAttachedFiles((prev) => [...prev, ...newFiles]);
      }
    } catch {
      /* 静默失败 */
    }
  }, []);

  const handleOpenImage = useCallback(async () => {
    setShowAttachPopover(false);
    try {
      const result = await ipcApi.dialog.openImage();
      if (!result.canceled && result.filePaths.length > 0) {
        const newFiles: FileAttachment[] = result.filePaths.map((fp) => ({
          filePath: fp,
          fileName: fp.split(/[\\/]/).pop() || fp,
          fileType: "image" as const,
        }));
        setAttachedFiles((prev) => [...prev, ...newFiles]);
      }
    } catch {
      /* 静默失败 */
    }
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
          const newFiles: FileAttachment[] = files.map((f) => ({
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
    } catch {
      /* 静默失败 */
    }
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
      e.clientX <= rect.left ||
      e.clientX >= rect.right ||
      e.clientY <= rect.top ||
      e.clientY >= rect.bottom
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
      .filter((item) => item.kind === "file")
      .map((item) => item.getAsFile())
      .filter(Boolean) as File[];
    if (files.length === 0) return;

    const newFiles = await resolveDroppedFiles(files);

    if (newFiles.length > 0) {
      e.preventDefault();
      setAttachedFiles((prev) => {
        const existing = new Set(prev.map((f) => f.filePath));
        return [...prev, ...newFiles.filter((f) => !existing.has(f.filePath))];
      });
    }
  }, []);

  const closeComposerPopovers = useCallback(() => {
    setShowAttachPopover(false);
    setShowPermissionPopover(false);
    setShowThinkingPopover(false);
  }, []);

  useDocumentDismiss({
    active: showAttachPopover || showPermissionPopover || showThinkingPopover,
    closeOnEscape: false,
    ignoreSelectors: COMPOSER_POPOVER_IGNORE_SELECTORS,
    onDismiss: closeComposerPopovers,
  });

  const closeComposerFolderPopovers = useCallback(() => {
    setShowFolderFileList(false);
    setShowComposerFolderList(false);
  }, []);

  useDocumentDismiss({
    active: showFolderFileList || showComposerFolderList,
    closeOnEscape: false,
    ignoreSelectors: COMPOSER_FOLDER_IGNORE_SELECTORS,
    onDismiss: closeComposerFolderPopovers,
  });

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
    handleTextareaChange,
    handlePaste,
  };
}
