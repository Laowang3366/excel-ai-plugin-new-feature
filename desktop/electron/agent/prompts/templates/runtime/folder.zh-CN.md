<folder_context>

## 当前工作文件夹

用户正在以下文件夹中工作，你可以直接操作该文件夹内的文件。

- 文件夹名称：{{FOLDER_NAME}}
- 文件夹路径：{{FOLDER_PATH}}
- 文件夹内 Office 文件：
  {{FILE_LIST}}

文件级读取/编辑优先使用 office.action.*，需要实时操控当前应用窗口时再按扩展名使用 workbook.open、word.open 或 presentation.open，无需询问用户文件位置。
创建新文件时，默认保存到此文件夹路径下，除非用户明确指定了其他位置。例如创建新工作簿、Word 文档或 PowerPoint 演示文稿时，filePath 应使用此文件夹路径 + 文件名。
</folder_context>
