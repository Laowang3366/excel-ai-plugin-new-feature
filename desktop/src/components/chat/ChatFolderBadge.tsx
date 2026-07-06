import React from "react";
import type { FolderFileInfo } from "../../electronApi";
import type { PinnedFolder } from "../../store/settingsStore";
import { formatFileSize } from "../../utils/fileSize";
import { FileSpreadsheet, FolderOpen } from "../common/IconMap";

interface ChatFolderBadgeText {
  folderFileCount: (count: number) => string;
}

interface ChatFolderBadgeProps {
  folder: PinnedFolder;
  files: FolderFileInfo[];
  open: boolean;
  text: ChatFolderBadgeText;
  onToggleOpen: () => void;
  onClose: () => void;
  onAddFilesToComposer: (files: Array<{
    filePath: string;
    fileName: string;
    fileType: "document";
    size?: number;
  }>) => void;
  onHideBadge: () => void;
}

export const ChatFolderBadge: React.FC<ChatFolderBadgeProps> = ({
  folder,
  files,
  open,
  text,
  onToggleOpen,
  onClose,
  onAddFilesToComposer,
  onHideBadge,
}) => {
  return (
    <div className="chat-folder-badge-wrapper">
      <button
        className="chat-folder-badge"
        title={folder.path}
        onClick={onToggleOpen}
      >
        <FolderOpen size={13} />
        <span className="chat-folder-name">{folder.name}</span>
        {files.length > 0 && (
          <span className="chat-folder-file-count">
            {text.folderFileCount(files.length)}
          </span>
        )}
      </button>
      {open && files.length > 0 && (
        <div className="folder-file-popover" onClick={(event) => event.stopPropagation()}>
          <div className="folder-file-popover-title">{folder.name}</div>
          {files.map((file) => (
            <button
              key={file.filePath}
              className="folder-file-popover-item clickable"
              title={file.filePath}
              onClick={() => {
                onAddFilesToComposer([{
                  filePath: file.filePath,
                  fileName: file.fileName,
                  fileType: "document",
                  size: file.size,
                }]);
                onClose();
                onHideBadge();
              }}
            >
              <FileSpreadsheet size={13} />
              <span className="folder-file-popover-name">{file.fileName}</span>
              <span className="folder-file-popover-size">{formatFileSize(file.size)}</span>
            </button>
          ))}
        </div>
      )}
    </div>
  );
};
