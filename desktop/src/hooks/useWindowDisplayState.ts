import { useEffect, useState } from "react";

import type { WindowDisplayMode } from "../electronApi";
import { ipcApi } from "../services/ipcApi";
import { logWarn } from "../utils/rendererLogger";

interface WindowDisplayStateOptions {
  autoCompactEnabled: boolean;
  hasConnectedOffice: boolean;
}

function requestLayoutReflow(): void {
  void document.documentElement.offsetWidth;
  window.requestAnimationFrame(() => {
    void document.documentElement.offsetWidth;
  });
}

export function useWindowDisplayState({
  autoCompactEnabled,
  hasConnectedOffice,
}: WindowDisplayStateOptions) {
  const [alwaysOnTop, setAlwaysOnTop] = useState(true);
  const [displayMode, setDisplayMode] = useState<WindowDisplayMode>("normal");

  useEffect(() => {
    ipcApi.window
      .getAlwaysOnTop()
      .then(setAlwaysOnTop)
      .catch(() => {
        logWarn("App", "获取窗口置顶状态失败，使用默认值");
        setAlwaysOnTop(true);
      });
  }, []);

  useEffect(() => {
    let disposed = false;
    let resizeSyncTimer: number | undefined;

    const applyActualMode = (mode: WindowDisplayMode, forceReflow = false) => {
      if (disposed) return;
      setDisplayMode(mode);
      if (forceReflow) requestLayoutReflow();
    };

    const syncActualMode = () => {
      ipcApi.window
        .getDisplayMode()
        .then((mode) => applyActualMode(mode))
        .catch(() => {
          logWarn("App", "获取窗口显示模式失败");
          applyActualMode("normal");
        });
    };

    const handleResize = () => {
      if (resizeSyncTimer !== undefined) window.clearTimeout(resizeSyncTimer);
      resizeSyncTimer = window.setTimeout(syncActualMode, 120);
    };

    const unsubscribeDisplayMode = ipcApi.window.onDisplayModeChanged((mode) => {
      applyActualMode(mode, true);
    });

    syncActualMode();
    window.addEventListener("resize", handleResize);

    return () => {
      disposed = true;
      if (resizeSyncTimer !== undefined) window.clearTimeout(resizeSyncTimer);
      unsubscribeDisplayMode();
      window.removeEventListener("resize", handleResize);
    };
  }, []);

  useEffect(() => {
    if (!autoCompactEnabled || !hasConnectedOffice) return;

    const handleBlur = () => {
      setDisplayMode((currentMode) => {
        if (currentMode !== "normal") return currentMode;
        ipcApi.window
          .setDisplayMode("compact")
          .then(setDisplayMode)
          .catch(() => {
            logWarn("App", "设置紧凑模式失败");
            setDisplayMode(currentMode);
          });
        return "compact";
      });
    };

    window.addEventListener("blur", handleBlur);
    return () => window.removeEventListener("blur", handleBlur);
  }, [autoCompactEnabled, hasConnectedOffice]);

  const toggleAlwaysOnTop = async () => {
    const next = !alwaysOnTop;
    setAlwaysOnTop(next);
    try {
      const actual = await ipcApi.window.setAlwaysOnTop(next);
      if (typeof actual === "boolean") setAlwaysOnTop(actual);
    } catch {
      logWarn("App", "切换窗口置顶失败");
      setAlwaysOnTop(alwaysOnTop);
    }
  };

  const setWindowMode = async (mode: WindowDisplayMode) => {
    setDisplayMode(mode);
    try {
      setDisplayMode(await ipcApi.window.setDisplayMode(mode));
    } catch {
      logWarn("App", "切换窗口模式失败");
      setDisplayMode("normal");
    }
  };

  return {
    alwaysOnTop,
    displayMode,
    toggleAlwaysOnTop,
    toggleCompactMode: () => {
      void setWindowMode(displayMode === "normal" ? "compact" : "normal");
    },
  };
}
