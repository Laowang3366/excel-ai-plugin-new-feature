import folderPrompt from "../templates/runtime/folder.zh-CN.md?raw";
import { appendPromptSections, renderPromptTemplate } from "../promptComposer";

export interface FolderFileItem {
  fileName: string;
  filePath: string;
  size: number;
}

export function appendFolderContext(
  systemPrompt: string,
  folderPath: string,
  folderName: string,
  files: FolderFileItem[],
): string {
  if (!folderPath) return systemPrompt;

  const fileInfoLines =
    files.length > 0
      ? files
          .map((f, i) => {
            const sizeStr = f.size > 0 ? ` (${(f.size / 1024).toFixed(0)}KB)` : "";
            return `${i + 1}. ${f.fileName}${sizeStr}`;
          })
          .join("\n")
      : "（文件夹内暂无可直接操作的 Office 文件）";

  const folderContext = renderPromptTemplate(folderPrompt, {
    FOLDER_NAME: folderName,
    FOLDER_PATH: folderPath,
    FILE_LIST: fileInfoLines,
  });

  return appendPromptSections(systemPrompt, [{ key: "folder-context", content: folderContext }]);
}
