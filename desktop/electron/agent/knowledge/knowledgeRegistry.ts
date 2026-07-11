/**
 * RAG 知识库全局注册表
 *
 * 提供全局的 getter/setter，方便 main.ts 初始化 RAG 模块后，
 * 让 buildStreamParams.ts 等模块获取 Retriever 实例进行自动知识注入。
 *
 * 设计理由：避免修改 AgentLoopConfig 的接口签名，
 * 沿用 existing 的模块级 setter 模式（参照 eventForwarder.ts）。
 * resetKnowledgeRegistry 用于测试、数据路径切换和运行期重新初始化时清理旧实例。
 */

import type { Retriever } from "./retriever";
import type { SqliteStore } from "./sqliteStore";
import type { KnowledgeIndexer } from "./knowledgeIndexer";
import type { KnowledgeWriter } from "./knowledgeWriter";

let _retriever: Retriever | null = null;
let _store: SqliteStore | null = null;
let _indexer: KnowledgeIndexer | null = null;
let _writer: KnowledgeWriter | null = null;

/** 注册知识检索器 */
export function setKnowledgeRetriever(retriever: Retriever | null): void {
  _retriever = retriever;
}

/** 获取知识检索器 */
export function getKnowledgeRetriever(): Retriever | null {
  return _retriever;
}

/** 注册 SQLite 存储 */
export function setKnowledgeStore(store: SqliteStore | null): void {
  _store = store;
}

/** 获取 SQLite 存储 */
export function getKnowledgeStore(): SqliteStore | null {
  return _store;
}

/** 注册知识索引器 */
export function setKnowledgeIndexer(indexer: KnowledgeIndexer | null): void {
  _indexer = indexer;
}

/** 获取知识索引器 */
export function getKnowledgeIndexer(): KnowledgeIndexer | null {
  return _indexer;
}

/** 注册知识写入器 */
export function setKnowledgeWriter(writer: KnowledgeWriter | null): void {
  _writer = writer;
}

/** 获取知识写入器 */
export function getKnowledgeWriter(): KnowledgeWriter | null {
  return _writer;
}

/** 清空已注册的知识库实例 */
export function resetKnowledgeRegistry(): void {
  _retriever = null;
  _store = null;
  _indexer = null;
  _writer = null;
}
