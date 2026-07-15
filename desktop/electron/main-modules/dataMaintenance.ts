export const DATA_MAINTENANCE_BUSY_MESSAGE = "数据存储正在维护，请稍后重试";

let activeDataOperations = 0;

export function assertDataMaintenanceAvailable(isBusy?: () => boolean): void {
  if (isBusy?.()) throw new Error(DATA_MAINTENANCE_BUSY_MESSAGE);
}

export function guardDataOperation<TArgs extends unknown[], TResult>(
  isMaintenanceInProgress: (() => boolean) | undefined,
  operation: (...args: TArgs) => Promise<TResult> | TResult,
) {
  return async (...args: TArgs): Promise<TResult> => {
    assertDataMaintenanceAvailable(isMaintenanceInProgress);
    activeDataOperations += 1;
    try {
      return await operation(...args);
    } finally {
      activeDataOperations -= 1;
    }
  };
}

export function hasActiveDataOperations(): boolean {
  return activeDataOperations > 0;
}
