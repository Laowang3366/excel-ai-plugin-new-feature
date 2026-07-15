import { useCallback, useState } from "react";
import type { MouseEvent } from "react";
import type { ContextMenuState } from "../components/sidebar/ThreadContextMenu";

interface UseSidebarThreadContextMenuParams {
  deleteThread: (threadId: string) => Promise<void>;
  moveThreadToFolder: (threadId: string, folderId?: string) => Promise<void>;
}

export function useSidebarThreadContextMenu({
  deleteThread,
  moveThreadToFolder,
}: UseSidebarThreadContextMenuParams) {
  const [contextMenu, setContextMenu] = useState<ContextMenuState | null>(null);

  const openThreadContextMenu = useCallback(
    (event: MouseEvent, threadId: string, inFolder?: boolean) => {
      event.preventDefault();
      event.stopPropagation();
      setContextMenu({ threadId, x: event.clientX, y: event.clientY, confirming: false, inFolder });
    },
    [],
  );

  const handleConfirmDelete = useCallback(
    async (threadId: string) => {
      setContextMenu(null);
      await deleteThread(threadId);
    },
    [deleteThread],
  );

  const handleMoveToFolder = useCallback(
    async (threadId: string, folderId?: string) => {
      setContextMenu(null);
      await moveThreadToFolder(threadId, folderId);
    },
    [moveThreadToFolder],
  );

  const closeContextMenu = useCallback(() => {
    setContextMenu(null);
  }, []);

  const setContextMoveMenu = useCallback((show: boolean) => {
    setContextMenu((menu) => (menu ? { ...menu, showMoveMenu: show } : menu));
  }, []);

  const setContextConfirming = useCallback((confirming: boolean) => {
    setContextMenu((menu) => (menu ? { ...menu, confirming } : menu));
  }, []);

  return {
    contextMenu,
    openThreadContextMenu,
    handleConfirmDelete,
    handleMoveToFolder,
    closeContextMenu,
    setContextMoveMenu,
    setContextConfirming,
  };
}
