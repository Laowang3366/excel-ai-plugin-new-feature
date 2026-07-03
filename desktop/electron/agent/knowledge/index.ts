/**
 * RAG 知识增强层 — 统一导出
 */

export * from "./types";
export { EmbeddingService } from "./embeddingService";
export { SqliteStore } from "./sqliteStore";
export { DocumentParser } from "./documentParser";
export { TextChunker } from "./textChunker";
export { KnowledgeIndexer } from "./knowledgeIndexer";
export { KnowledgeWriter } from "./knowledgeWriter";
export { Retriever } from "./retriever";
export { WorkbookNotesStore } from "./workbookNotesStore";
export {
  setKnowledgeRetriever,
  getKnowledgeRetriever,
  setKnowledgeStore,
  getKnowledgeStore,
  setKnowledgeIndexer,
  getKnowledgeIndexer,
  setKnowledgeWriter,
  getKnowledgeWriter,
  resetKnowledgeRegistry,
  isKnowledgeAvailable,
} from "./knowledgeRegistry";
