/**
 * 工作簿笔记存储
 *
 * 自动将 workbook.inspect 结果中的关键信息索引为知识条目，
 * 实现跨会话的知识沉淀。
 *
 * 触发时机：
 * - 每次 workbook.inspect 工具返回结果后
 * - 提取表结构（sheet 名、列数、行数、表头）
 * - 以 work book:${filePath} 为来源存储
 */

import { randomUUID } from "crypto";
import type { KnowledgeEntry, KnowledgeSource } from "./types";
import { SqliteStore } from "./sqliteStore";

// ============================================================
// WorkbookNotesStore
// ============================================================

export class WorkbookNotesStore {
  private store: SqliteStore;

  constructor(store: SqliteStore) {
    this.store = store;
  }

  /**
   * 从 workbook.inspect 结果中提取笔记并索引
   *
   * @param filePath - 工作簿文件路径
   * @param fileName - 文件名
   * @param inspectResult - workbook.inspect 的返回结果
   */
  indexFromInspect(
    filePath: string,
    fileName: string,
    inspectResult: Record<string, unknown>
  ): void {
    const now = Date.now();

    // 提取工作表信息
    const sheets = inspectResult.sheets as
      | Array<{
          name: string;
          rowCount?: number;
          colCount?: number;
          usedRange?: string;
        }>
      | undefined;

    if (!sheets || sheets.length === 0) return;

    const entries: KnowledgeEntry[] = [];
    const sourcePath = filePath;

    // 先删除该工作簿的旧笔记
    this.store.deleteSource(sourcePath);

    // 为每个 sheet 创建条目
    for (const sheet of sheets) {
      const sheetName = sheet.name || "未知";
      const rowCount = sheet.rowCount || 0;
      const colCount = sheet.colCount || 0;
      const usedRange = sheet.usedRange || "";

      // 构建内容
      const contentLines: string[] = [
        `工作簿: ${fileName}`,
        `工作表: ${sheetName}`,
        `行数: ${rowCount}`,
        `列数: ${colCount}`,
      ];
      if (usedRange) {
        contentLines.push(`使用区域: ${usedRange}`);
      }

      entries.push({
        id: randomUUID(),
        source: "workbook",
        sourcePath,
        sourceName: fileName,
        sourceType: "xlsx",
        chunkIndex: 0,
        content: contentLines.join("\n"),
        metadata: {
          sheetName,
          rowCount,
          colCount,
          usedRange,
          isAutoIndexed: true,
        },
        embedding: null, // 不嵌入，后续在索引流程中统一嵌入
        indexedAt: now,
        tokenCount: Math.ceil(contentLines.join("\n").length / 3.5),
      });
    }

    // 以 workbook path 为来源存储
    if (entries.length > 0) {
      this.store.bulkInsert(entries);
      this.store.upsertSource({
        sourcePath,
        sourceName: fileName,
        sourceType: "xlsx",
        entryCount: entries.length,
        firstIndexed: now,
        lastIndexed: now,
        fileHash: `auto-${now}`, // 自动索引标记
      });
    }
  }

  /**
   * 删除指定工作簿的笔记
   */
  deleteWorkbookNotes(filePath: string): void {
    this.store.deleteSource(filePath);
  }

  /**
   * 强制为工作簿笔记生成嵌入向量
   *
   * 由 KnowledgeIndexer 在后续流程中统一调用。
   */
  getEntriesForEmbedding(filePath: string): KnowledgeEntry[] {
    return this.store.getEntriesBySource(filePath);
  }
}
