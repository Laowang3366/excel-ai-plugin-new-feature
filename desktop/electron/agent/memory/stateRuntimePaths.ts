import * as fs from "fs";
import * as path from "path";

import type {
  ResolvedRuntimePaths,
  RuntimeDatabasePaths,
  RuntimeDbName,
} from "./stateRuntimeTypes";

/**
 * StateRuntime 路径管理。
 *
 * 关联模块：
 * - stateRuntimeStore.ts: 初始化四库前解析路径并执行旧 state-runtime.db 迁移。
 */
export function defaultStateRuntimeRoot(): string {
  const appData = process.env.APPDATA || path.join(process.env.USERPROFILE || "C:\\", "AppData", "Roaming");
  return path.join(appData, "excel-ai-assistant", "sessions", "state-runtime");
}

export function runtimeDbNames(): RuntimeDbName[] {
  return ["state", "logs", "goals", "memories"];
}

export function resolveRuntimeDatabasePaths(runtimeRoot: string): ResolvedRuntimePaths {
  if (runtimeRoot === ":memory:") {
    return {
      dbPaths: {
        state: ":memory:",
        logs: ":memory:",
        goals: ":memory:",
        memories: ":memory:",
      },
    };
  }

  const isLegacyFilePath = runtimeRoot.toLowerCase().endsWith(".db");
  const root = isLegacyFilePath
    ? path.join(path.dirname(runtimeRoot), path.basename(runtimeRoot, ".db"))
    : runtimeRoot;
  const legacyStateDbPath = isLegacyFilePath ? runtimeRoot : `${runtimeRoot}.db`;
  return {
    dbPaths: {
      state: path.join(root, "state.db"),
      logs: path.join(root, "logs.db"),
      goals: path.join(root, "goals.db"),
      memories: path.join(root, "memories.db"),
    },
    legacyStateDbPath,
  };
}

export function isMemoryRuntime(paths: RuntimeDatabasePaths): boolean {
  return runtimeDbNames().every((name) => paths[name] === ":memory:");
}

export async function migrateLegacyStateDbIfNeeded(
  stateDbPath: string,
  legacyStateDbPath?: string
): Promise<void> {
  if (!legacyStateDbPath || legacyStateDbPath === stateDbPath) return;
  if (fs.existsSync(stateDbPath) || !fs.existsSync(legacyStateDbPath)) return;
  await fs.promises.copyFile(legacyStateDbPath, stateDbPath);
}
