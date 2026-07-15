import { describe, expect, it, vi } from "vitest";

import { createAppShutdownController, runCleanupSteps } from "./appShutdown";

describe("runCleanupSteps", () => {
  it("continues remaining cleanup steps after an earlier failure", async () => {
    const second = vi.fn().mockResolvedValue(undefined);
    await expect(
      runCleanupSteps([vi.fn().mockRejectedValue(new Error("flush failed")), second]),
    ).rejects.toThrow("应用退出清理未全部完成");
    expect(second).toHaveBeenCalledTimes(1);
  });
});

describe("createAppShutdownController", () => {
  it("prevents the first quit and quits only after cleanup finishes", async () => {
    let finishCleanup!: () => void;
    const cleanup = vi.fn(
      () =>
        new Promise<void>((resolve) => {
          finishCleanup = resolve;
        }),
    );
    const quit = vi.fn();
    const preventDefault = vi.fn();
    const handler = createAppShutdownController({ cleanup, quit });

    handler({ preventDefault });
    handler({ preventDefault });

    expect(preventDefault).toHaveBeenCalledTimes(2);
    expect(cleanup).toHaveBeenCalledTimes(1);
    expect(quit).not.toHaveBeenCalled();

    finishCleanup();
    await vi.waitFor(() => expect(quit).toHaveBeenCalledTimes(1));

    handler({ preventDefault });
    expect(preventDefault).toHaveBeenCalledTimes(2);
  });

  it("reports cleanup failure and still completes the second quit", async () => {
    const error = new Error("flush failed");
    const onCleanupError = vi.fn();
    const quit = vi.fn();
    const handler = createAppShutdownController({
      cleanup: vi.fn().mockRejectedValue(error),
      quit,
      onCleanupError,
    });

    handler({ preventDefault: vi.fn() });

    await vi.waitFor(() => expect(quit).toHaveBeenCalledTimes(1));
    expect(onCleanupError).toHaveBeenCalledWith(error);
  });
});
