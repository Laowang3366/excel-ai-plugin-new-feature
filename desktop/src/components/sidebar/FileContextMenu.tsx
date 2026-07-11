/**
 * FileContextMenu — 侧边栏文件夹文件的右键菜单
 *
 * 功能：
 * - 置顶 / 取消置顶
 * - 重命名
 * - 删除（移到回收站）
 * - 打开（系统默认应用）
 * - 复制路径
 * - 在文件管理器中显示
 */

import type { FolderFileInfo } from "../../electronApi";
import { getAppText } from "../../i18n";
import type { AppLanguage } from "../../store/settingsStore";
import {
  Pin,
  Trash2,
  Zap,
  Copy,
  FolderOpen,
} from "../common/IconMap";

export interface FileContextMenuState {
  file: FolderFileInfo;
  x: number;
  y: number;
  isPinned: boolean;
}

interface FileContextMenuProps {
  state: FileContextMenuState;
  language: AppLanguage;
  onClose: () => void;
  onDelete: (filePath: string) => void;
  onOpen: (filePath: string) => void;
  onCopyPath: (filePath: string) => void;
  onRevealInExplorer: (filePath: string) => void;
  onPinFile: (filePath: string) => void;
}

export function FileContextMenu({
  state,
  language,
  onClose,
  onDelete,
  onOpen,
  onCopyPath,
  onRevealInExplorer,
  onPinFile,
}: FileContextMenuProps) {
  const text = getAppText(language);

  return (
    <div
      className="thread-context-menu"
      style={{ left: state.x, top: state.y }}
      onClick={(e) => e.stopPropagation()}
      onContextMenu={(e) => e.preventDefault()}
    >
      <button
        className="thread-context-menu-item"
        onClick={() => { onPinFile(state.file.filePath); onClose(); }}
      >
        <Pin size={14} />
        <span>{state.isPinned ? text.sidebar.unpin : text.sidebar.pinToTop}</span>
      </button>
      <button
        className="thread-context-menu-item"
        onClick={() => { onOpen(state.file.filePath); onClose(); }}
      >
        <Zap size={14} />
        <span>{text.sidebar.openFile}</span>
      </button>
      <button
        className="thread-context-menu-item"
        onClick={() => { onCopyPath(state.file.filePath); onClose(); }}
      >
        <Copy size={14} />
        <span>{text.sidebar.copyPath}</span>
      </button>
      <button
        className="thread-context-menu-item"
        onClick={() => { onRevealInExplorer(state.file.filePath); onClose(); }}
      >
        <FolderOpen size={14} />
        <span>{text.sidebar.revealInExplorer}</span>
      </button>
      <div className="thread-context-menu-divider" />
      <button
        className="thread-context-menu-item danger"
        onClick={() => { onDelete(state.file.filePath); onClose(); }}
      >
        <Trash2 size={14} />
        <span>{text.sidebar.deleteFile}</span>
      </button>
    </div>
  );
}
