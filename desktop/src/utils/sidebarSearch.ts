import type { FolderFileInfo, ThreadMetadata } from "../electronApi";
import type { PinnedFolder } from "../store/settingsStore";

export interface SidebarSearchAction {
  id: string;
  label: string;
}

export interface SidebarSearchThreadResult {
  thread: ThreadMetadata;
  folder?: PinnedFolder;
}

export interface SidebarSearchFileResult {
  file: FolderFileInfo;
  folder: PinnedFolder;
}

export interface SidebarSearchResults {
  threads: SidebarSearchThreadResult[];
  files: SidebarSearchFileResult[];
  actions: SidebarSearchAction[];
}

interface BuildSidebarSearchResultsInput {
  query: string;
  threads: ThreadMetadata[];
  folders: PinnedFolder[];
  folderFiles: Record<string, FolderFileInfo[]>;
  actions: SidebarSearchAction[];
  limitPerGroup?: number;
}

function includesQuery(values: Array<string | undefined | null>, query: string): boolean {
  if (!query) return true;
  return values.some((value) => value?.toLowerCase().includes(query));
}

export function buildSidebarSearchResults({
  query,
  threads,
  folders,
  folderFiles,
  actions,
  limitPerGroup = 6,
}: BuildSidebarSearchResultsInput): SidebarSearchResults {
  const normalizedQuery = query.trim().toLowerCase();
  const folderByPath = new Map(folders.map((folder) => [folder.path, folder]));

  const threadResults = threads
    .map((thread) => ({ thread, folder: thread.folderId ? folderByPath.get(thread.folderId) : undefined }))
    .filter(({ thread, folder }) =>
      includesQuery([thread.name, thread.preview, folder?.name, folder?.path], normalizedQuery)
    )
    .sort((left, right) => right.thread.updatedAt - left.thread.updatedAt)
    .slice(0, limitPerGroup);

  const fileResults = folders
    .flatMap((folder) =>
      (folderFiles[folder.path] || []).map((file) => ({
        file,
        folder,
      }))
    )
    .filter(({ file }) => includesQuery([file.fileName, file.filePath], normalizedQuery))
    .sort((left, right) => right.file.lastModified - left.file.lastModified)
    .slice(0, limitPerGroup);

  const actionResults = actions
    .filter((action) => includesQuery([action.label], normalizedQuery))
    .slice(0, limitPerGroup);

  return {
    threads: threadResults,
    files: fileResults,
    actions: actionResults,
  };
}
