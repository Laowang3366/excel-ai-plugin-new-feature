/** 知识库 IPC 注册，以及带路径授权复核的重建流程。 */

import { assertAuthorizedPath, type PathAuthorizer } from "../../main-modules/ipcPathSecurity";
import {
  KnowledgeDeleteInput,
  KnowledgeIndexFileInput,
  KnowledgeIndexFolderInput,
  KnowledgeSearchInput,
  validateInput,
} from "../../shared/ipcSchemas";
import { trustedIpcMain as ipcMain } from "../../shared/trustedIpc";
import {
  getKnowledgeIndexer,
  getKnowledgeRetriever,
  getKnowledgeStore,
} from "../knowledge/knowledgeRegistry";
import type { IndexResult } from "../knowledge/types";
import type { KnowledgeRuntimeState } from "../runtime/knowledgeRuntime";

export interface KnowledgeIpcHandlerDeps {
  ensureKnowledgeRuntime?: () => Promise<KnowledgeRuntimeState>;
  isDataMigrationInProgress?: () => boolean;
  pathAuthorizer: PathAuthorizer;
}

async function ensureKnowledgeRuntimeForIpc(deps: KnowledgeIpcHandlerDeps): Promise<string | null> {
  if (deps.isDataMigrationInProgress?.()) return "数据存储正在迁移，请稍后重试";
  try {
    const runtime = await deps.ensureKnowledgeRuntime?.();
    return runtime?.error || null;
  } catch (error: unknown) {
    return error instanceof Error ? error.message : String(error || "未知错误");
  }
}

function formatKnowledgeUnavailableError(error?: string | null): string {
  return error ? `知识库未初始化：${error}` : "知识库未初始化";
}

function formatError(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}

export async function reindexAuthorizedKnowledgeSources(
  indexer: Pick<ReturnType<typeof getKnowledgeIndexer> & object, "listSources" | "indexFile">,
  pathAuthorizer: PathAuthorizer,
): Promise<IndexResult[]> {
  const results: IndexResult[] = [];
  for (const source of indexer.listSources()) {
    try {
      const sourcePath = assertAuthorizedPath(pathAuthorizer, source.sourcePath);
      results.push(await indexer.indexFile(sourcePath, { skipUnchanged: false }));
    } catch (error) {
      results.push({
        sourcePath: source.sourcePath,
        success: false,
        error: formatError(error),
        entryCount: 0,
        durationMs: 0,
      });
    }
  }
  return results;
}

export function registerKnowledgeIpcHandlers(deps: KnowledgeIpcHandlerDeps): void {
  ipcMain.handle("knowledge:listSources", async () => {
    if (deps.isDataMigrationInProgress?.()) throw new Error("数据存储正在迁移，请稍后重试");
    let store = getKnowledgeStore();
    const initError = !store ? await ensureKnowledgeRuntimeForIpc(deps) : null;
    store = getKnowledgeStore();
    if (!store) throw new Error(formatKnowledgeUnavailableError(initError));
    return store.listSources();
  });

  ipcMain.handle("knowledge:search", async (_event, query: unknown, topK: unknown) => {
    if (deps.isDataMigrationInProgress?.()) {
      return { success: false, error: "数据存储正在迁移，请稍后重试" };
    }
    const validated = validateInput(KnowledgeSearchInput, { query, topK });
    let retriever = getKnowledgeRetriever();
    const initError = !retriever ? await ensureKnowledgeRuntimeForIpc(deps) : null;
    retriever = getKnowledgeRetriever();
    if (!retriever) {
      return { success: false, error: formatKnowledgeUnavailableError(initError) };
    }
    try {
      const results = await retriever.search({
        text: validated.query,
        topK: validated.topK || 5,
      });
      return { success: true, data: results };
    } catch (error) {
      return { success: false, error: formatError(error) };
    }
  });

  ipcMain.handle("knowledge:indexFile", async (_event, filePath: unknown) => {
    if (deps.isDataMigrationInProgress?.()) {
      return { success: false, error: "数据存储正在迁移，请稍后重试" };
    }
    const validated = validateInput(KnowledgeIndexFileInput, { filePath });
    let indexer = getKnowledgeIndexer();
    const initError = !indexer ? await ensureKnowledgeRuntimeForIpc(deps) : null;
    indexer = getKnowledgeIndexer();
    if (!indexer) {
      return { success: false, error: formatKnowledgeUnavailableError(initError) };
    }
    try {
      const sourcePath = assertAuthorizedPath(deps.pathAuthorizer, validated.filePath);
      return await indexer.indexFile(sourcePath);
    } catch (error) {
      return {
        success: false,
        error: formatError(error),
        sourcePath: validated.filePath,
        entryCount: 0,
        durationMs: 0,
      };
    }
  });

  ipcMain.handle("knowledge:indexFolder", async (_event, folderPath: unknown) => {
    if (deps.isDataMigrationInProgress?.()) {
      return { success: false, error: "数据存储正在迁移，请稍后重试" };
    }
    const validated = validateInput(KnowledgeIndexFolderInput, { folderPath });
    let indexer = getKnowledgeIndexer();
    const initError = !indexer ? await ensureKnowledgeRuntimeForIpc(deps) : null;
    indexer = getKnowledgeIndexer();
    if (!indexer) {
      return { success: false, error: formatKnowledgeUnavailableError(initError) };
    }
    try {
      const folderPath = assertAuthorizedPath(deps.pathAuthorizer, validated.folderPath);
      return await indexer.indexFolder(folderPath);
    } catch (error) {
      return { success: false, error: formatError(error) };
    }
  });

  ipcMain.handle("knowledge:deleteFile", async (_event, sourcePath: unknown) => {
    if (deps.isDataMigrationInProgress?.()) {
      return { success: false, error: "数据存储正在迁移，请稍后重试" };
    }
    const validated = validateInput(KnowledgeDeleteInput, { sourcePath });
    let indexer = getKnowledgeIndexer();
    const initError = !indexer ? await ensureKnowledgeRuntimeForIpc(deps) : null;
    indexer = getKnowledgeIndexer();
    if (!indexer) {
      return { success: false, error: formatKnowledgeUnavailableError(initError) };
    }
    try {
      const authorizedSourcePath = assertAuthorizedPath(deps.pathAuthorizer, validated.sourcePath);
      await indexer.deleteSource(authorizedSourcePath);
      return { success: true };
    } catch (error) {
      return { success: false, error: formatError(error) };
    }
  });

  ipcMain.handle("knowledge:reindexAll", async () => {
    if (deps.isDataMigrationInProgress?.()) {
      return { success: false, error: "数据存储正在迁移，请稍后重试" };
    }
    let indexer = getKnowledgeIndexer();
    const initError = !indexer ? await ensureKnowledgeRuntimeForIpc(deps) : null;
    indexer = getKnowledgeIndexer();
    if (!indexer) {
      return { success: false, error: formatKnowledgeUnavailableError(initError) };
    }
    try {
      const results = await reindexAuthorizedKnowledgeSources(indexer, deps.pathAuthorizer);
      return { success: true, results };
    } catch (error) {
      return { success: false, error: formatError(error) };
    }
  });
}
