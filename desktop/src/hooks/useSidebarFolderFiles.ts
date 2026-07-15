import { useCallback, useEffect, useState } from "react";
import type { MouseEvent } from "react";
import type { FolderFileInfo } from "../electronApi";
import type { PinnedFolder } from "../store/settingsStore";
import { ipcApi } from "../services/ipcApi";
import type { FileContextMenuState } from "../components/sidebar/FileContextMenu";

interface UseSidebarFolderFilesParams {
  pinnedFolders: PinnedFolder[];
  searchOpen: boolean;
  addPinnedFolder: (folder: PinnedFolder) => void;
  updatePinnedFolder: (folderPath: string, patch: Partial<PinnedFolder>) => void;
  addFilesToComposer: (
    files: Array<{
      filePath: string;
      fileName: string;
      fileType: "document";
      size?: number;
    }>,
  ) => void;
  onOpenFileMenu: () => void;
}

export function useSidebarFolderFiles({
  pinnedFolders,
  searchOpen,
  addPinnedFolder,
  updatePinnedFolder,
  addFilesToComposer,
  onOpenFileMenu,
}: UseSidebarFolderFilesParams) {
  const [folderFiles, setFolderFiles] = useState<Record<string, FolderFileInfo[]>>({});
  const [expandedFolders, setExpandedFolders] = useState<Record<string, boolean>>({});
  const [fileContextMenu, setFileContextMenu] = useState<FileContextMenuState | null>(null);

  const handleAddFolder = useCallback(async () => {
    try {
      const result = await ipcApi.dialog.openFolder();
      if (!result.canceled && result.filePaths.length > 0) {
        const folderPath = result.filePaths[0];
        const folderName = folderPath.split(/[\\/]/).pop() || folderPath;
        addPinnedFolder({ path: folderPath, name: folderName, addedAt: Date.now() });
        setExpandedFolders((prev) => ({ ...prev, [folderPath]: true }));
        const files = await ipcApi.folder.listFiles(folderPath);
        setFolderFiles((prev) => ({ ...prev, [folderPath]: files }));
      }
    } catch {
      // Folder selection and refresh failures are non-blocking UI actions.
    }
  }, [addPinnedFolder]);

  const handleToggleFolder = useCallback(
    async (folderPath: string) => {
      setExpandedFolders((prev) => {
        const next = !prev[folderPath];
        if (next && !folderFiles[folderPath]) {
          ipcApi.folder.listFiles(folderPath).then((files) => {
            setFolderFiles((prev2) => ({ ...prev2, [folderPath]: files }));
          });
        }
        return { ...prev, [folderPath]: next };
      });
    },
    [folderFiles],
  );

  useEffect(() => {
    if (!searchOpen) return;
    const missingPaths = pinnedFolders
      .map((folder) => folder.path)
      .filter((folderPath) => !folderFiles[folderPath]);
    if (missingPaths.length === 0) return;
    let cancelled = false;
    ipcApi.folder
      .listFilesBatch(missingPaths)
      .then((filesByFolder) => {
        if (cancelled) return;
        setFolderFiles((prev) => {
          const next = { ...prev };
          missingPaths.forEach((folderPath) => {
            next[folderPath] = filesByFolder[folderPath] || [];
          });
          return next;
        });
      })
      .catch(() => {
        if (cancelled) return;
        setFolderFiles((prev) => {
          const next = { ...prev };
          missingPaths.forEach((folderPath) => {
            next[folderPath] = [];
          });
          return next;
        });
      });
    return () => {
      cancelled = true;
    };
  }, [folderFiles, pinnedFolders, searchOpen]);

  const handleAddFile = useCallback(
    (file: FolderFileInfo) => {
      addFilesToComposer([
        {
          filePath: file.filePath,
          fileName: file.fileName,
          fileType: "document",
          size: file.size,
        },
      ]);
    },
    [addFilesToComposer],
  );

  const handleFileContextMenu = useCallback(
    (event: MouseEvent, file: FolderFileInfo, isPinned: boolean) => {
      event.preventDefault();
      event.stopPropagation();
      onOpenFileMenu();
      setFileContextMenu({ file, x: event.clientX, y: event.clientY, isPinned });
    },
    [onOpenFileMenu],
  );

  const refreshFolderContainingFile = useCallback(
    async (filePath: string) => {
      const folderPath = Object.keys(folderFiles).find((fp) =>
        folderFiles[fp].some((file) => file.filePath === filePath),
      );
      if (!folderPath) return;
      const files = await ipcApi.folder.listFiles(folderPath);
      setFolderFiles((prev) => ({ ...prev, [folderPath]: files }));
    },
    [folderFiles],
  );

  const handleTrashFile = useCallback(
    async (filePath: string) => {
      setFileContextMenu(null);
      await ipcApi.file.trashFile(filePath);
      await refreshFolderContainingFile(filePath);
    },
    [refreshFolderContainingFile],
  );

  const handleOpenFile = useCallback(async (filePath: string) => {
    setFileContextMenu(null);
    await ipcApi.file.openFile(filePath);
  }, []);

  const handleCopyPath = useCallback(async (filePath: string) => {
    setFileContextMenu(null);
    await ipcApi.file.copyPath(filePath);
  }, []);

  const handleRevealInExplorer = useCallback(async (filePath: string) => {
    setFileContextMenu(null);
    await ipcApi.file.revealInExplorer(filePath);
  }, []);

  const handlePinFile = useCallback(
    async (filePath: string) => {
      setFileContextMenu(null);
      const folderPath = Object.keys(folderFiles).find((fp) =>
        folderFiles[fp].some((file) => file.filePath === filePath),
      );
      const folder = folderPath
        ? pinnedFolders.find((item) => item.path === folderPath)
        : undefined;
      if (!folder || !folderPath) return;
      const pinned = folder.pinnedFiles || [];
      const nextPinned = pinned.includes(filePath)
        ? pinned.filter((item) => item !== filePath)
        : [...pinned, filePath];
      updatePinnedFolder(folderPath, { pinnedFiles: nextPinned });
    },
    [folderFiles, pinnedFolders, updatePinnedFolder],
  );

  const closeFileContextMenu = useCallback(() => {
    setFileContextMenu(null);
  }, []);

  return {
    folderFiles,
    expandedFolders,
    fileContextMenu,
    handleAddFolder,
    handleToggleFolder,
    handleAddFile,
    handleFileContextMenu,
    closeFileContextMenu,
    handleTrashFile,
    handleOpenFile,
    handleCopyPath,
    handleRevealInExplorer,
    handlePinFile,
  };
}
