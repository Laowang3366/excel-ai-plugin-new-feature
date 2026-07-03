/**
 * 文件夹上下文提示词：把当前工作文件夹和 Office 文件列表追加到系统提示词。
 *
 * 关联模块：
 * - ../systemPrompt.ts: 组装完整系统提示词。
 * - ../systemPrompt.test.ts: 校验关键提示词内容不丢失。
 */

// ============================================================
// 文件夹上下文注入
// ============================================================

/** 文件夹内文件信息（与前端 FolderFileInfo 对齐） */
export interface FolderFileItem {
  fileName: string;
  filePath: string;
  size: number;
}

/**
 * 将当前工作文件夹上下文追加到系统提示词末尾
 *
 * 让 AI 知道用户正在哪个文件夹中工作，以及该文件夹中有哪些 Office 文件。
 * AI 可以按扩展名使用 workbook.open 或 office.action.* 操作这些文件，
 * 无需用户手动指定路径。
 */
export function appendFolderContext(
  systemPrompt: string,
  folderPath: string,
  folderName: string,
  files: FolderFileItem[]
): string {
  if (!folderPath) return systemPrompt;

  const fileInfoLines = files.length > 0
    ? files.map((f, i) => {
        const sizeStr = f.size > 0 ? ` (${(f.size / 1024).toFixed(0)}KB)` : "";
        return `${i + 1}. ${f.fileName}${sizeStr}`;
      }).join("\n")
    : "（文件夹内暂无可直接操作的 Office 文件）";

  const folderContext = `## 当前工作文件夹

用户正在以下文件夹中工作，你可以直接操作该文件夹内的文件。

- 文件夹名称：${folderName}
- 文件夹路径：${folderPath}
- 文件夹内 Office 文件：
${fileInfoLines}

文件级读取/编辑优先使用 office.action.*，需要实时操控当前应用窗口时再按扩展名使用 workbook.open、word.open 或 presentation.open，无需询问用户文件位置。
**创建新文件时，默认保存到此文件夹路径下**，除非用户明确指定了其他位置。
例如创建新工作簿、Word 文档或 PowerPoint 演示文稿时，filePath 应使用此文件夹路径 + 文件名。`;

  return systemPrompt + "\n\n" + folderContext;
}
