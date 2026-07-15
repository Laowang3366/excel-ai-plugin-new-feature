import { exportUserDataDirectory } from "./userDataExport";

export interface UserDataExportCoordinatorDeps {
  isBusy: () => boolean;
  setBusy: (busy: boolean) => void;
  hasRunningAgent: () => boolean;
  getDataPath: () => string;
  getSanitizedSettings: () => Record<string, unknown>;
  getSessionStore: () => {
    suspendWrites: (reason: string) => void;
    resumeWrites: () => void;
    flushRolloutWrites: () => Promise<void>;
  };
  closeStateRuntime: () => Promise<void>;
  resetKnowledgeRuntime: () => void;
  restoreRuntimes: () => Promise<void>;
}

export interface CoordinatedUserDataExportResult {
  success: boolean;
  exportPath?: string;
  exportedAt?: string;
  categories?: string[];
  error?: string;
}

export async function runUserDataExport(
  targetPath: string,
  deps: UserDataExportCoordinatorDeps,
): Promise<CoordinatedUserDataExportResult> {
  if (deps.isBusy()) {
    return { success: false, error: "数据维护或相关操作正在进行中，请稍后重试" };
  }
  const trimmedPath = targetPath.trim();
  if (!trimmedPath) return { success: false, error: "请选择有效的导出目录" };
  if (deps.hasRunningAgent()) {
    return { success: false, error: "请等待当前会话执行完成或停止后再导出数据" };
  }

  deps.setBusy(true);
  const sessionStore = deps.getSessionStore();
  let result: Awaited<ReturnType<typeof exportUserDataDirectory>> | undefined;
  let operationError: unknown;
  let recoveryError: unknown;
  try {
    sessionStore.suspendWrites("正在导出本地数据，请稍后重试");
    await sessionStore.flushRolloutWrites();
    await deps.closeStateRuntime();
    deps.resetKnowledgeRuntime();
    result = await exportUserDataDirectory({
      sourceDataPath: deps.getDataPath(),
      targetPath: trimmedPath,
      sanitizedSettings: deps.getSanitizedSettings(),
    });
  } catch (error) {
    operationError = error;
  } finally {
    sessionStore.resumeWrites();
    try {
      await deps.restoreRuntimes();
    } catch (error) {
      recoveryError = error;
    }
    deps.setBusy(false);
  }

  if (operationError) {
    return {
      success: false,
      error: operationError instanceof Error ? operationError.message : "导出本地数据失败",
    };
  }
  if (recoveryError || !result) {
    return {
      success: false,
      exportPath: result?.exportPath,
      error: `数据已导出，但恢复本地运行时失败：${recoveryError instanceof Error ? recoveryError.message : "未知错误"}`,
    };
  }
  return { success: true, ...result };
}
