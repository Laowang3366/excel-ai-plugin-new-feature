import { useMemo } from "react";
import type { ThreadMetadata, FolderFileInfo } from "../electronApi";
import type { AppLanguage, PinnedFolder } from "../store/settingsStore";
import { buildSidebarDerivedLists, type SidebarSortMode } from "../utils/sidebarHelpers";

interface UseSidebarDerivedListsInput {
  threads: ThreadMetadata[];
  pinnedFolders: PinnedFolder[];
  folderFiles: Record<string, FolderFileInfo[]>;
  projectSortMode: SidebarSortMode;
  conversationSortMode: SidebarSortMode;
  language: AppLanguage;
  hasSearchQuery: boolean;
}

export function useSidebarDerivedLists(input: UseSidebarDerivedListsInput) {
  const {
    threads,
    pinnedFolders,
    folderFiles,
    projectSortMode,
    conversationSortMode,
    language,
    hasSearchQuery,
  } = input;

  return useMemo(
    () =>
      buildSidebarDerivedLists({
        threads,
        pinnedFolders,
        folderFiles,
        projectSortMode,
        conversationSortMode,
        language,
        hasSearchQuery,
      }),
    [
      conversationSortMode,
      folderFiles,
      hasSearchQuery,
      language,
      pinnedFolders,
      projectSortMode,
      threads,
    ],
  );
}
