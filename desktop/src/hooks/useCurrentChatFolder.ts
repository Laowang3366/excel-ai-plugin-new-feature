import { useEffect, useMemo, useState } from "react";

import type { FolderFileInfo } from "../electronApi";
import { ipcApi } from "../services/ipcApi";
import type { PinnedFolder } from "../store/settingsStore";

export function useCurrentChatFolder(
  currentFolderId: string | null | undefined,
  pinnedFolders: PinnedFolder[],
) {
  const [badgeHidden, setBadgeHidden] = useState(false);
  const [files, setFiles] = useState<FolderFileInfo[]>([]);
  const folder = useMemo(
    () => pinnedFolders.find((item) => item.path === currentFolderId),
    [currentFolderId, pinnedFolders],
  );

  useEffect(() => {
    setBadgeHidden(false);
  }, [currentFolderId]);

  useEffect(() => {
    let cancelled = false;
    if (!currentFolderId) {
      setFiles([]);
      return;
    }

    ipcApi.folder
      .listFiles(currentFolderId)
      .then((nextFiles) => {
        if (!cancelled) setFiles(nextFiles);
      })
      .catch(() => {
        if (!cancelled) setFiles([]);
      });

    return () => {
      cancelled = true;
    };
  }, [currentFolderId]);

  return {
    badgeHidden,
    files,
    folder,
    hideBadge: () => setBadgeHidden(true),
  };
}
