import { useMemo } from "react";
import type { FolderFileInfo } from "../electronApi";
import type {
  FolderSectionActions,
  FolderSectionFileMenuApi,
  FolderSectionThreadActions,
} from "../components/sidebar/FolderSection";
import type { FileContextMenuState } from "../components/sidebar/FileContextMenu";

interface UseSidebarSectionActionsInput {
  handleToggleFolder: (path: string) => void;
  handleCreateFolderThread: (path: string) => void;
  removePinnedFolder: (path: string) => void;
  handleSwitchThread: (threadId: string) => void;
  handleThreadContextMenu: (e: React.MouseEvent, threadId: string, inFolder?: boolean) => void;
  fileContextMenu: FileContextMenuState | null;
  handleAddFile: (file: FolderFileInfo) => void;
  handleFileContextMenu: (e: React.MouseEvent, file: FolderFileInfo, isPinned: boolean) => void;
  closeFileContextMenu: () => void;
  handleTrashFile: (filePath: string) => void;
  handleOpenFile: (filePath: string) => void;
  handleCopyPath: (filePath: string) => void;
  handleRevealInExplorer: (filePath: string) => void;
  handlePinFile: (filePath: string) => void;
}

export function useSidebarSectionActions(input: UseSidebarSectionActionsInput): {
  folderActions: FolderSectionActions;
  folderThreadActions: FolderSectionThreadActions;
  folderFileMenuApi: FolderSectionFileMenuApi;
} {
  const {
    handleToggleFolder,
    handleCreateFolderThread,
    removePinnedFolder,
    handleSwitchThread,
    handleThreadContextMenu,
    fileContextMenu,
    handleAddFile,
    handleFileContextMenu,
    closeFileContextMenu,
    handleTrashFile,
    handleOpenFile,
    handleCopyPath,
    handleRevealInExplorer,
    handlePinFile,
  } = input;

  const folderActions = useMemo<FolderSectionActions>(
    () => ({
      toggle: handleToggleFolder,
      createThread: handleCreateFolderThread,
      remove: removePinnedFolder,
    }),
    [handleCreateFolderThread, handleToggleFolder, removePinnedFolder],
  );

  const folderThreadActions = useMemo<FolderSectionThreadActions>(
    () => ({
      switchThread: handleSwitchThread,
      openContextMenu: handleThreadContextMenu,
    }),
    [handleSwitchThread, handleThreadContextMenu],
  );

  const folderFileMenuApi = useMemo<FolderSectionFileMenuApi>(
    () => ({
      state: fileContextMenu,
      addFile: handleAddFile,
      openContextMenu: handleFileContextMenu,
      close: closeFileContextMenu,
      trashFile: handleTrashFile,
      openFile: handleOpenFile,
      copyPath: handleCopyPath,
      revealInExplorer: handleRevealInExplorer,
      pinFile: handlePinFile,
    }),
    [
      closeFileContextMenu,
      fileContextMenu,
      handleAddFile,
      handleCopyPath,
      handleFileContextMenu,
      handleOpenFile,
      handlePinFile,
      handleRevealInExplorer,
      handleTrashFile,
    ],
  );

  return { folderActions, folderThreadActions, folderFileMenuApi };
}
