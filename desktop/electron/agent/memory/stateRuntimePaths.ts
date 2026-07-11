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
 * - stateRuntimeStore.ts: 初始化四库前解析路径。
 */
export function defaultStateRuntimeRoot(): string {
  const appData =
    process.env.APPDATA || path.join(process.env.USERPROFILE || "C:\\", "AppData", "Roaming");
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

  return {
    dbPaths: {
      state: path.join(runtimeRoot, "state.db"),
      logs: path.join(runtimeRoot, "logs.db"),
      goals: path.join(runtimeRoot, "goals.db"),
      memories: path.join(runtimeRoot, "memories.db"),
    },
  };
}

export function isMemoryRuntime(paths: RuntimeDatabasePaths): boolean {
  return runtimeDbNames().every((name) => paths[name] === ":memory:");
}
