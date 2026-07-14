/**
 * 窗口管理器 — Electron 窗口创建与托盘管理
 *
 * 从 main.ts 提取，包含 createWindow、createTray、showMainWindow、applyWindowTheme
 *
 * 注意：createWindow 不再接收 ref 对象参数，直接返回 BrowserWindow 实例。
 * 调用方负责将返回值赋值给自己的变量。
 */

import { BrowserWindow, Tray, Menu, app, screen } from "electron";
import * as path from "path";
import * as fs from "fs";
import { pathToFileURL } from "url";
import { getSettingsStore, applyWindowOpacity, applyWindowTheme } from "./settingsManager";
import { disableActiveHotPatch, resolveHotPatchPath } from "./hotPatchManager";
import { isAllowedWindowNavigation } from "./windowNavigationPolicy";

let tray: Tray | null = null;
let isQuitting = false;
let displayMode: WindowDisplayMode = "normal";
let normalBounds: Electron.Rectangle | null = null;
let applyingDisplayMode = false;

export type WindowDisplayMode = "normal" | "compact";

const NORMAL_MIN_SIZE = { width: 900, height: 500 };
const COMPACT_SIZE = { width: 420, minWidth: 360, minHeight: 500 };
const WORK_AREA_GAP = 16;
const APP_DISPLAY_NAME = "文格 AI 助手";

export function setIsQuitting(value: boolean): void { isQuitting = value; }
export function getWindowDisplayMode(): WindowDisplayMode { return displayMode; }

function getIconPath(): string {
  // 生产环境: extraResources 将 public/ 复制到 resources/public/
  // 开发环境: public/ 在项目根目录下，__dirname 指向 electron/
  const prodPath = path.join(process.resourcesPath || "", "public", "icon.png");
  if (process.resourcesPath && fs.existsSync(prodPath)) {
    return prodPath;
  }
  return path.join(__dirname, "../public/icon.png");
}

/** 隐藏菜单栏 — 包含标准编辑快捷键（复制/粘贴/剪切/全选/撤销） */
const editMenu = Menu.buildFromTemplate([
  {
    label: "编辑",
    submenu: [
      { role: "undo", label: "撤销" },
      { role: "redo", label: "重做" },
      { type: "separator" },
      { role: "cut", label: "剪切" },
      { role: "copy", label: "复制" },
      { role: "paste", label: "粘贴" },
      { role: "selectAll", label: "全选" },
    ],
  },
]);

export function createWindow(
  /** 当关闭到托盘时需要重建窗口，此回调提供重建能力 */
  recreateWindow?: () => BrowserWindow,
  /** 窗口创建完成后回调，方便调用方保存引用 */
  onCreated?: (mw: BrowserWindow) => void
): BrowserWindow {
  const initialTheme = getSettingsStore().get("theme") === "dark" ? "dark" : "light";
  const mw = new BrowserWindow({
    width: 1200,
    height: 720,
    minWidth: 900,
    minHeight: 500,
    title: APP_DISPLAY_NAME,
    icon: getIconPath(),
    webPreferences: {
      preload: path.join(__dirname, "preload.js"),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: true,
    },
    backgroundColor: initialTheme === "dark" ? "#0f172a" : "#ffffff",
    titleBarStyle: "hidden",
    titleBarOverlay: {
      color: initialTheme === "dark" ? "#0b1220" : "#eef5fb",
      symbolColor: initialTheme === "dark" ? "#f8fafc" : "#111827",
      height: 36,
    },
    alwaysOnTop: true,
    autoHideMenuBar: true,
    resizable: true,
  });

  const developmentUrl = process.env.VITE_DEV_SERVER_URL;
  const bundledIndex = path.join(__dirname, "../dist/index.html");
  const patchedIndex = resolveHotPatchPath("dist/index.html");
  const appUrl = developmentUrl ?? pathToFileURL(patchedIndex ?? bundledIndex).toString();

  mw.webContents.on("will-navigate", (event, targetUrl) => {
    if (!isAllowedWindowNavigation(targetUrl, appUrl)) event.preventDefault();
  });
  mw.webContents.on("will-redirect", (event, targetUrl) => {
    if (!isAllowedWindowNavigation(targetUrl, appUrl)) event.preventDefault();
  });
  mw.webContents.setWindowOpenHandler(() => ({ action: "deny" }));

  if (developmentUrl) {
    void mw.loadURL(developmentUrl);
  } else {
    void mw.loadFile(patchedIndex ?? bundledIndex).catch(async () => {
      if (!patchedIndex) return;
      await disableActiveHotPatch(app.getPath("userData"));
      await mw.loadFile(bundledIndex);
    });
  }

  mw.on("page-title-updated", (event) => {
    event.preventDefault();
    mw?.setTitle(APP_DISPLAY_NAME);
  });

  mw.setMenuBarVisibility(false);
  mw.setMenu(editMenu);
  applyWindowTheme(mw);
  applyWindowOpacity(mw);

  mw.on("close", (event) => {
    if (!isQuitting && getSettingsStore().get("closeToTray") === true) {
      event.preventDefault();
      createTray(mw, recreateWindow);
      mw?.hide();
      mw?.setSkipTaskbar(true);
    }
  });

  mw.on("closed", () => {
    /* 调用方应清理自己的引用 */
  });

  mw.on("resize", () => rememberNormalBounds(mw));
  mw.on("move", () => rememberNormalBounds(mw));

  onCreated?.(mw);
  return mw;
}

export function setWindowDisplayMode(
  mainWindow: BrowserWindow | null,
  mode: WindowDisplayMode
): WindowDisplayMode {
  if (!mainWindow || mainWindow.isDestroyed()) return displayMode;
  if (displayMode === mode) return displayMode;

  applyingDisplayMode = true;
  try {
    if (displayMode === "normal" && mode !== "normal") {
      normalBounds = mainWindow.getBounds();
    }
    if (mode === "normal") {
      applyNormalWindowMode(mainWindow);
    } else {
      applyCompactWindowMode(mainWindow);
    }

    displayMode = mode;
    notifyRendererDisplayModeChanged(mainWindow, displayMode);
    requestRendererLayoutRefresh(mainWindow);
    mainWindow.show();
    mainWindow.focus();
    return displayMode;
  } finally {
    applyingDisplayMode = false;
  }
}

function rememberNormalBounds(mainWindow: BrowserWindow): void {
  if (applyingDisplayMode || displayMode !== "normal" || mainWindow.isDestroyed()) return;
  normalBounds = mainWindow.getBounds();
}

function applyNormalWindowMode(mainWindow: BrowserWindow): void {
  mainWindow.setResizable(true);
  mainWindow.setMinimumSize(NORMAL_MIN_SIZE.width, NORMAL_MIN_SIZE.height);
  mainWindow.setSkipTaskbar(false);
  applyWindowTheme(mainWindow);
  applyWindowOpacity(mainWindow);

  const fallbackBounds = mainWindow.getBounds();
  const nextBounds = normalizeNormalBounds(normalBounds ?? fallbackBounds);
  mainWindow.setBounds(nextBounds, true);
}

function applyCompactWindowMode(mainWindow: BrowserWindow): void {
  mainWindow.setResizable(true);
  mainWindow.setMinimumSize(COMPACT_SIZE.minWidth, COMPACT_SIZE.minHeight);
  mainWindow.setSkipTaskbar(false);
  applyWindowTheme(mainWindow);
  applyWindowOpacity(mainWindow);
  mainWindow.setBounds(getSideDockBounds(mainWindow, COMPACT_SIZE.width), true);
}

function notifyRendererDisplayModeChanged(
  mainWindow: BrowserWindow,
  mode: WindowDisplayMode
): void {
  if (mainWindow.isDestroyed() || mainWindow.webContents.isDestroyed()) return;
  mainWindow.webContents.send("window:displayModeChanged", mode);
}

function requestRendererLayoutRefresh(mainWindow: BrowserWindow): void {
  if (mainWindow.isDestroyed() || mainWindow.webContents.isDestroyed()) return;

  const dispatchResize = () => {
    if (mainWindow.isDestroyed() || mainWindow.webContents.isDestroyed()) return;
    void mainWindow.webContents.executeJavaScript(
      "window.dispatchEvent(new Event('resize')); document.documentElement.offsetWidth;",
      true
    ).catch(() => undefined);
  };

  dispatchResize();
  setTimeout(dispatchResize, 80);
}

function getCurrentWorkArea(mainWindow: BrowserWindow): Electron.Rectangle {
  const currentBounds = mainWindow.getBounds();
  return screen.getDisplayMatching(currentBounds).workArea;
}

function getSideDockBounds(mainWindow: BrowserWindow, width: number): Electron.Rectangle {
  const workArea = getCurrentWorkArea(mainWindow);
  const height = Math.max(
    COMPACT_SIZE.minHeight,
    Math.min(760, workArea.height - WORK_AREA_GAP * 2)
  );
  return {
    width,
    height,
    x: workArea.x + workArea.width - width - WORK_AREA_GAP,
    y: workArea.y + Math.max(WORK_AREA_GAP, Math.round((workArea.height - height) / 2)),
  };
}

function normalizeNormalBounds(bounds: Electron.Rectangle): Electron.Rectangle {
  const display = screen.getDisplayMatching(bounds);
  const workArea = display.workArea;
  const width = Math.max(NORMAL_MIN_SIZE.width, Math.min(bounds.width, workArea.width));
  const height = Math.max(NORMAL_MIN_SIZE.height, Math.min(bounds.height, workArea.height));
  return {
    width,
    height,
    x: clamp(bounds.x, workArea.x, workArea.x + workArea.width - width),
    y: clamp(bounds.y, workArea.y, workArea.y + workArea.height - height),
  };
}

function clamp(value: number, min: number, max: number): number {
  if (max < min) return min;
  return Math.min(Math.max(value, min), max);
}

function createTray(mainWindow: BrowserWindow, recreateWindow?: () => BrowserWindow): void {
  if (tray) return;

  tray = new Tray(getIconPath());
  tray.setToolTip(APP_DISPLAY_NAME);
  const contextMenu = Menu.buildFromTemplate([
    {
      label: "显示主窗口",
      click: () => {
        if (mainWindow.isDestroyed() && recreateWindow) {
          recreateWindow();
          return;
        }
        mainWindow.show();
        mainWindow.setSkipTaskbar(false);
        mainWindow.focus();
      },
    },
    {
      label: "退出",
      click: () => {
        isQuitting = true;
        app.quit();
      },
    },
  ]);
  tray.setContextMenu(contextMenu);
  tray.on("click", () => {
    if (mainWindow.isDestroyed() && recreateWindow) {
      recreateWindow();
      return;
    }
    mainWindow.show();
    mainWindow.setSkipTaskbar(false);
    mainWindow.focus();
  });
}
