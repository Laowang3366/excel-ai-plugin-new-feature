/**
 * 工具执行器装配入口
 *
 * 组合各领域执行器，保持 createToolExecutors 对外签名不变。
 */

import type { ToolExecutor } from "../../shared/types";
import type {
  ExcelConnectionBridge,
  ExcelWorkbookBridge,
  ExcelVbaBridge,
  ExcelScriptBridge,
  ExcelUiBridge,
} from "../contracts/excel";
import type {
  WordDocumentBridge,
  PresentationBridge,
  OfficeScriptBridge,
  OfficeActionBridge,
} from "../contracts/office";
import type { Retriever } from "../../knowledge/retriever";
import type { LongTermMemoryStore } from "../../memory/longTerm/memoryStore";
import { addExcelExecutors } from "./excelExecutors";
import { addFileExecutors } from "./fileExecutors";
import { addKnowledgeExecutors } from "./knowledgeExecutors";
import { addMemoryExecutors } from "./memoryExecutors";
import { addOfficeExecutors } from "./officeExecutors";
import { addOcrExecutors } from "./ocrExecutors";
import { addPythonExecutors } from "./pythonExecutor";
import { addShellExecutors, executeShellCommand } from "./shellExecutor";
import { addWebSearchExecutors } from "./webSearchExecutors";
import type { ShellCommandResult } from "./shellExecutor";
import { getToolNameAliases } from "../registry/toolDefinitions";

export { executeShellCommand };
export type { ShellCommandResult };

export interface ToolExecutorRuntimeDeps {
  getMineruApiToken?: () => string;
}

/**
 * 根据桥接接口创建工具执行器映射。
 */
export function createToolExecutors(
  workbookBridge: ExcelWorkbookBridge,
  vbaBridge: ExcelVbaBridge,
  scriptBridge: ExcelScriptBridge,
  uiBridge: ExcelUiBridge,
  sessionFolderPath?: string,
  knowledgeRetriever?: Retriever,
  wordBridge?: WordDocumentBridge,
  presentationBridge?: PresentationBridge,
  officeScriptBridge?: OfficeScriptBridge,
  officeActionBridge?: OfficeActionBridge,
  memoryStore?: LongTermMemoryStore,
  runtimeDeps: ToolExecutorRuntimeDeps = {}
): Map<string, ToolExecutor> {
  const executors = new Map<string, ToolExecutor>();

  addExcelExecutors(executors, { workbookBridge, vbaBridge, scriptBridge, uiBridge });
  addFileExecutors(executors, { sessionFolderPath });
  addShellExecutors(executors);
  addPythonExecutors(executors);
  addKnowledgeExecutors(executors, { knowledgeRetriever });
  addWebSearchExecutors(executors);
  addOcrExecutors(executors, { getMineruApiToken: runtimeDeps.getMineruApiToken });
  addMemoryExecutors(executors, { memoryStore });
  addOfficeExecutors(executors, {
    excelBridge: workbookBridge as ExcelConnectionBridge,
    wordBridge,
    presentationBridge,
    officeScriptBridge,
    officeActionBridge,
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
