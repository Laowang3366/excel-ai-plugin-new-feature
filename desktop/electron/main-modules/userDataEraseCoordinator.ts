import {
  eraseManagedUserData,
  USER_DATA_ERASE_CONFIRMATION,
  type UserDataEraseReport,
} from "./userDataErase";

export interface UserDataEraseCoordinatorDeps {
  isBusy: () => boolean;
  setBusy: (busy: boolean) => void;
  hasRunningAgent: () => boolean;
  getDataPath: () => string;
  getSessionStore: () => {
    suspendWrites: (reason: string) => void;
    resumeWrites: () => void;
    flushRolloutWrites: () => Promise<void>;
  };
  resetAgents: () => Promise<void>;
  closeStateRuntime: () => Promise<void>;
  resetKnowledgeRuntime: () => void;
  clearSettings: () => void;
  restoreRuntimes: () => Promise<void>;
  eraseManagedData?: (dataPath: string) => Promise<UserDataEraseReport>;
}

export interface CoordinatedUserDataEraseResult extends UserDataEraseReport {
  success: boolean;
  error?: string;
}

export async function runUserDataErase(
  confirmation: string,
  deps: UserDataEraseCoordinatorDeps,
): Promise<CoordinatedUserDataEraseResult> {
  const empty = { erasedCategories: [], errors: [] };
  if (confirmation !== USER_DATA_ERASE_CONFIRMATION) {
    return { success: false, ...empty, error: "确认文字不匹配，未擦除任何数据" };
  }
  if (deps.isBusy()) {
    return { success: false, ...empty, error: "数据维护或相关操作正在进行中" };
  }
  if (deps.hasRunningAgent()) {
    return { success: false, ...empty, error: "请等待当前会话执行完成或停止后再擦除数据" };
  }

  deps.setBusy(true);
  const sessionStore = deps.getSessionStore();
  let report: UserDataEraseReport = empty;
  let operationError: unknown;
  let recoveryError: unknown;
  try {
    sessionStore.suspendWrites("正在擦除本地数据，请稍后重试");
    await sessionStore.flushRolloutWrites();
    await deps.resetAgents();
    await deps.closeStateRuntime();
    deps.resetKnowledgeRuntime();
    deps.clearSettings();
    report = { erasedCategories: ["settings"], errors: [] };
    const managedReport = await (deps.eraseManagedData ?? eraseManagedUserData)(deps.getDataPath());
    report = {
      erasedCategories: [...report.erasedCategories, ...managedReport.erasedCategories],
      errors: managedReport.errors,
    };
  } catch (error) {
    operationError = error;
  } finally {
    try {
      await deps.restoreRuntimes();
    } catch (error) {
      recoveryError = error;
    }
    sessionStore.resumeWrites();
    deps.setBusy(false);
  }

  const errors = [...report.errors];
  if (operationError) errors.push(errorMessage(operationError));
  if (recoveryError) errors.push(`恢复本地运行时失败: ${errorMessage(recoveryError)}`);
  if (errors.length > 0) {
    return {
      success: false,
      erasedCategories: report.erasedCategories,
      errors,
      error: `本地数据擦除未完全完成：${errors.join("；")}`,
    };
  }
  return { success: true, ...report };
}

function errorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}
