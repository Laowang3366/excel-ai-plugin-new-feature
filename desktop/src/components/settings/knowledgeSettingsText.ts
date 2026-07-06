export interface KnowledgeSource {
  sourcePath: string;
  sourceName: string;
  sourceType: string;
  entryCount: number;
  firstIndexed: number;
  lastIndexed: number;
  fileHash: string;
}

export interface KnowledgeIndexResultSummary {
  success?: boolean;
  entryCount?: number;
}

export const KNOWLEDGE_TEXT = {
  "zh-CN": {
    title: "知识库",
    desc: "管理本地知识库，自动为 AI 助手提供工作簿结构、字段含义等上下文信息。",
    enableTitle: "启用知识库",
    enableHint: "开启后 AI 助手会自动检索相关知识并注入对话上下文。",
    sourcesTitle: "知识来源",
    sourcesDesc: "已索引的文件列表，AI 在对话时会自动参考这些知识。",
    noSources: "暂无知识来源。点击下方按钮添加文件或文件夹。",
    sourceStats: "已索引 {files} 个来源，共 {entries} 条知识",
    sourcePathLabel: "路径",
    sourceCol: "来源文件",
    typeCol: "类型",
    entriesCol: "条目数",
    indexedCol: "索引时间",
    actionsCol: "操作",
    addFile: "添加文件",
    addFolder: "添加文件夹",
    reindexAll: "重建全部索引",
    indexing: "索引中...",
    deleteSource: "删除",
    deleteConfirm: "确定要删除该来源的索引吗？",
    reindexConfirm: "确定要重建全部索引吗？这可能需要一些时间。",
    reindexAllProgress: "正在重建全部索引 ({current}/{total})...",
    indexFileSuccess: "文件索引完成，共 {count} 条",
    indexFolderSuccess: "文件夹索引完成，成功 {success} 个，失败 {failed} 个，共 {count} 条",
    deleteSuccess: "已删除索引",
    error: "操作失败",
    workbook: "工作簿",
    document: "文档",
    json: "JSON",
    word: "Word",
    presentation: "演示文稿",
    note: "笔记",
    agents_md: "项目知识",
  },
  "en-US": {
    title: "Knowledge Base",
    desc: "Manage local knowledge base that provides workbook structure, field meanings, and other context to the AI assistant.",
    enableTitle: "Enable Knowledge Base",
    enableHint: "When enabled, the AI automatically retrieves relevant knowledge and injects it into conversation context.",
    sourcesTitle: "Knowledge Sources",
    sourcesDesc: "Indexed files that the AI references during conversations.",
    noSources: "No knowledge sources yet. Click below to add files or folders.",
    sourceStats: "{files} sources indexed, {entries} knowledge entries",
    sourcePathLabel: "Path",
    sourceCol: "Source File",
    typeCol: "Type",
    entriesCol: "Entries",
    indexedCol: "Indexed At",
    actionsCol: "Actions",
    addFile: "Add File",
    addFolder: "Add Folder",
    reindexAll: "Rebuild All Indexes",
    indexing: "Indexing...",
    deleteSource: "Delete",
    deleteConfirm: "Delete this source?",
    reindexConfirm: "Rebuild all indexes? This may take some time.",
    reindexAllProgress: "Rebuilding all indexes ({current}/{total})...",
    indexFileSuccess: "File indexed, {count} entries created",
    indexFolderSuccess: "Folder indexed: {success} succeeded, {failed} failed, {count} entries",
    deleteSuccess: "Index deleted",
    error: "Operation failed",
    workbook: "Workbook",
    document: "Document",
    json: "JSON",
    word: "Word",
    presentation: "Presentation",
    note: "Note",
    agents_md: "Project Knowledge",
  },
} as const;

export type KnowledgeSettingsLanguage = keyof typeof KNOWLEDGE_TEXT;
export type KnowledgeSettingsText = (typeof KNOWLEDGE_TEXT)[KnowledgeSettingsLanguage];

export function formatKnowledgeSourceStats(
  text: KnowledgeSettingsText,
  sourceCount: number,
  totalEntries: number,
): string {
  return text.sourceStats
    .replace("{files}", String(sourceCount))
    .replace("{entries}", String(totalEntries));
}

export function formatKnowledgeFolderIndexSuccess(
  text: KnowledgeSettingsText,
  results: KnowledgeIndexResultSummary[],
): string {
  const success = results.filter((item) => item?.success).length;
  const failed = results.length - success;
  const count = results.reduce((sum, item) => sum + (item?.entryCount || 0), 0);
  return text.indexFolderSuccess
    .replace("{success}", String(success))
    .replace("{failed}", String(failed))
    .replace("{count}", String(count));
}

export function formatKnowledgeTime(timestamp: number, language: KnowledgeSettingsLanguage): string {
  return new Date(timestamp).toLocaleString(language === "zh-CN" ? "zh-CN" : "en-US");
}

export function getKnowledgeSourceTypeLabel(text: KnowledgeSettingsText, type: string): string {
  const map: Record<string, string> = {
    workbook: text.workbook,
    document: text.document,
    json: text.json,
    docx: text.word,
    pptx: text.presentation,
    note: text.note,
    agents_md: text.agents_md,
    xlsx: text.workbook,
    xlsm: text.workbook,
    csv: text.document,
    md: text.agents_md,
    txt: text.document,
  };
  return map[type] || type;
}
