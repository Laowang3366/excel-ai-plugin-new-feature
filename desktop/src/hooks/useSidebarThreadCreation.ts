import { useCallback, useState } from "react";

export function useSidebarThreadCreation(createNewThread: (folderPath?: string) => Promise<void>) {
  const [creatingFolderThread, setCreatingFolderThread] = useState<string | null>(null);
  const [creatingNewThread, setCreatingNewThread] = useState(false);

  const handleCreateFolderThread = useCallback(
    async (folderPath: string) => {
      setCreatingFolderThread(folderPath);
      try {
        await createNewThread(folderPath);
      } finally {
        setTimeout(() => setCreatingFolderThread(null), 300);
      }
    },
    [createNewThread],
  );

  const handleCreateNewThread = useCallback(async () => {
    setCreatingNewThread(true);
    try {
      await createNewThread();
    } finally {
      setTimeout(() => setCreatingNewThread(false), 300);
    }
  }, [createNewThread]);

  return {
    creatingFolderThread,
    creatingNewThread,
    handleCreateFolderThread,
    handleCreateNewThread,
  };
}
