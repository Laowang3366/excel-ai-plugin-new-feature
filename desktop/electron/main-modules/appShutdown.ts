export interface BeforeQuitEventLike {
  preventDefault(): void;
}

export interface AppShutdownControllerOptions {
  cleanup: () => Promise<void>;
  quit: () => void;
  onCleanupError?: (error: unknown) => void;
}

export async function runCleanupSteps(steps: Array<() => Promise<void>>): Promise<void> {
  const errors: unknown[] = [];
  for (const step of steps) {
    try {
      await step();
    } catch (error) {
      errors.push(error);
    }
  }
  if (errors.length > 0) {
    const cleanupError = new Error("应用退出清理未全部完成") as Error & { causes: unknown[] };
    cleanupError.causes = errors;
    throw cleanupError;
  }
}

export function createAppShutdownController(options: AppShutdownControllerOptions) {
  let cleanupStarted = false;
  let cleanupCompleted = false;

  return (event: BeforeQuitEventLike): void => {
    if (cleanupCompleted) return;

    event.preventDefault();
    if (cleanupStarted) return;
    cleanupStarted = true;

    void options.cleanup()
      .catch((error) => options.onCleanupError?.(error))
      .finally(() => {
        cleanupCompleted = true;
        options.quit();
      });
  };
}
