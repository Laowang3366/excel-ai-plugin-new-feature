/**
 * 工具执行器装配入口
 *
 * 组合各领域执行器，保持 createToolExecutors 对外签名不变。
 */

import path from "node:path";

import type { ToolExecutor } from "../../shared/types";
import type {
  ExcelConnectionBridge,
  ExcelWorkbookBridge,
  ExcelVbaBridge,
  WpsJsaBridge,
  ExcelUiBridge,
} from "../contracts/excel";
import type {
  WordDocumentBridge,
  PresentationBridge,
  OfficeActionBridge,
  OfficeDocumentManagerBridge,
} from "../contracts/office";
import type { Retriever } from "../../knowledge/retriever";
import type { LongTermMemoryStore } from "../../memory/longTerm/memoryStore";
import { addExcelExecutors } from "./excelExecutors";
import { addFileExecutors } from "./fileExecutors";
import { addKnowledgeExecutors } from "./knowledgeExecutors";
import { addMemoryExecutors } from "./memoryExecutors";
import { addOfficeExecutors } from "./officeExecutors";
import { addOcrExecutors } from "./ocrExecutors";
import { addWebSearchExecutors } from "./webSearchExecutors";
import { getToolNameAliases } from "../registry/toolDefinitions";

export interface ToolExecutorRuntimeDeps {
  getMineruApiToken?: () => string;
  isRemoteDataProcessingEnabled?: () => boolean;
  officeDocumentBridge?: OfficeDocumentManagerBridge;
  officeAutomationRoot?: string;
}

/**
 * 根据桥接接口创建工具执行器映射。
 */
export function createToolExecutors(
  workbookBridge: ExcelWorkbookBridge,
  vbaBridge: ExcelVbaBridge,
  jsaBridge: WpsJsaBridge,
  uiBridge: ExcelUiBridge,
  sessionFolderPath?: string,
  knowledgeRetriever?: Retriever,
  wordBridge?: WordDocumentBridge,
  presentationBridge?: PresentationBridge,
  officeActionBridge?: OfficeActionBridge,
  memoryStore?: LongTermMemoryStore,
  runtimeDeps: ToolExecutorRuntimeDeps = {}
): Map<string, ToolExecutor> {
  const executors = new Map<string, ToolExecutor>();

  addExcelExecutors(executors, { workbookBridge, vbaBridge, jsaBridge, uiBridge });
  addFileExecutors(executors, { sessionFolderPath });
  addKnowledgeExecutors(executors, { knowledgeRetriever });
  addWebSearchExecutors(executors, {
    isRemoteDataProcessingEnabled: runtimeDeps.isRemoteDataProcessingEnabled,
  });
  addOcrExecutors(executors, {
    getMineruApiToken: runtimeDeps.getMineruApiToken,
    isRemoteDataProcessingEnabled: runtimeDeps.isRemoteDataProcessingEnabled,
  });
  addMemoryExecutors(executors, { memoryStore });
  addOfficeExecutors(executors, {
    excelBridge: workbookBridge as ExcelConnectionBridge,
    wordBridge,
    presentationBridge,
    officeActionBridge,
    officeDocumentBridge: runtimeDeps.officeDocumentBridge,
    workflowRoot: runtimeDeps.officeAutomationRoot ? path.join(runtimeDeps.officeAutomationRoot, "workflows") : undefined,
    transactionRoot: runtimeDeps.officeAutomationRoot ? path.join(runtimeDeps.officeAutomationRoot, "transactions") : undefined,
  });
  addExecutorAliases(executors);

  return executors;
}

function addExecutorAliases(executors: Map<string, ToolExecutor>): void {
  const entries = Array.from(executors.entries());
  for (const [toolName, executor] of entries) {
    for (const alias of getToolNameAliases(toolName)) {
      if (!executors.has(alias)) {
        executors.set(alias, { ...executor, name: alias });
      }
    }
  }
}
