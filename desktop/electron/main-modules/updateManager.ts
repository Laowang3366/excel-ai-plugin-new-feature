import { app, net } from "electron";
import { promises as fsp } from "node:fs";
import * as path from "node:path";

import { autoUpdater } from "electron-updater";

import { createLogger } from "../shared/logger";
import {
  activateInstalledHotPatch,
  getActiveHotPatchId,
  installHotPatchArchive,
  sha256File,
} from "./hotPatchManager";
import {
  compareVersions,
  type RemoteUpdateManifest,
  verifyRemoteUpdateManifest,
} from "./updateManifest";

export type UpdateKind = "installer" | "hotPatch";
export type UpdatePhase =
  | "idle"
  | "checking"
  | "up-to-date"
  | "available"
  | "downloading"
  | "downloaded"
  | "applying"
  | "error";

export interface DesktopUpdateState {
  phase: UpdatePhase;
  currentVersion: string;
  availableVersion?: string;
  installerAvailable: boolean;
  hotPatchAvailable: boolean;
  activeHotPatchId?: string;
  downloadedKind?: UpdateKind;
  progress?: number;
  releaseNotes: string[];
  publishedAt?: string;
  error?: string;
}

interface UpdateManagerOptions {
  userDataPath: string;
  notify: (state: DesktopUpdateState) => void;
  prepareToRestart: () => Promise<void>;
  updateBaseUrl?: string;
}

const updateLogger = createLogger("UpdateManager");
const DEFAULT_UPDATE_BASE_URL = "https://plugin.shelelove.top";
const AUTO_CHECK_DELAY_MS = 12_000;
const AUTO_CHECK_INTERVAL_MS = 6 * 60 * 60 * 1_000;

let options: UpdateManagerOptions | null = null;
let remoteManifest: RemoteUpdateManifest | null = null;
let restartInProgress = false;
let autoCheckTimer: NodeJS.Timeout | null = null;
let autoCheckInterval: NodeJS.Timeout | null = null;
let state: DesktopUpdateState = {
  phase: "idle",
  currentVersion: app.getVersion(),
  installerAvailable: false,
  hotPatchAvailable: false,
  releaseNotes: [],
};

function updateBaseUrl(): string {
  return (options?.updateBaseUrl || process.env.WENGE_UPDATE_BASE_URL || DEFAULT_UPDATE_BASE_URL)
    .replace(/\/+$/u, "");
}

function emitState(patch: Partial<DesktopUpdateState>): DesktopUpdateState {
  state = { ...state, ...patch };
  options?.notify(state);
  return state;
}

function publicKeyCandidates(): string[] {
  return [
    path.join(process.resourcesPath || "", "public", "update-public.pem"),
    path.join(process.cwd(), "public", "update-public.pem"),
    path.join(__dirname, "../public/update-public.pem"),
  ];
}

async function loadUpdatePublicKey(): Promise<string> {
  for (const candidate of publicKeyCandidates()) {
    try {
      return await fsp.readFile(candidate, "utf8");
    } catch {
      // Try the next packaged/development location.
    }
  }
  throw new Error("找不到更新签名公钥");
}

function availableState(manifest: RemoteUpdateManifest): Partial<DesktopUpdateState> {
  const activePatchId = options ? getActiveHotPatchId(options.userDataPath, app.getVersion()) : null;
  const installerAvailable = Boolean(
    manifest.installer && compareVersions(manifest.version, app.getVersion()) > 0,
  );
  const hotPatchAvailable = Boolean(
    manifest.hotPatch &&
      manifest.hotPatch.baseVersion === app.getVersion() &&
      manifest.hotPatch.id !== activePatchId,
  );
  return {
    phase: installerAvailable || hotPatchAvailable ? "available" : "up-to-date",
    availableVersion: manifest.version,
    installerAvailable,
    hotPatchAvailable,
    activeHotPatchId: activePatchId ?? undefined,
    downloadedKind: undefined,
    progress: undefined,
    releaseNotes: manifest.releaseNotes,
    publishedAt: manifest.publishedAt,
    error: undefined,
  };
}

export function activatePendingHotPatch(userDataPath: string): string | null {
  return activateInstalledHotPatch(app.getVersion(), userDataPath);
}

export function initializeUpdateManager(input: UpdateManagerOptions): void {
  options = input;
  state = {
    ...state,
    currentVersion: app.getVersion(),
    activeHotPatchId: getActiveHotPatchId(input.userDataPath, app.getVersion()) ?? undefined,
  };
  autoUpdater.autoDownload = false;
  autoUpdater.autoInstallOnAppQuit = false;
  autoUpdater.setFeedURL({ provider: "generic", url: `${updateBaseUrl()}/releases/windows` });
  autoUpdater.on("download-progress", (progress) => {
    emitState({ phase: "downloading", progress: Math.max(0, Math.min(100, progress.percent)) });
  });
  autoUpdater.on("error", (error) => {
    if (state.phase === "downloading") emitState({ phase: "error", error: error.message });
  });

  autoCheckTimer = setTimeout(() => void checkForUpdates(false), AUTO_CHECK_DELAY_MS);
  autoCheckTimer.unref();
  autoCheckInterval = setInterval(() => void checkForUpdates(false), AUTO_CHECK_INTERVAL_MS);
  autoCheckInterval.unref();
}

export function disposeUpdateManager(): void {
  if (autoCheckTimer) clearTimeout(autoCheckTimer);
  if (autoCheckInterval) clearInterval(autoCheckInterval);
  autoCheckTimer = null;
  autoCheckInterval = null;
}

export function getUpdateState(): DesktopUpdateState {
  return state;
}

export function isUpdateRestartInProgress(): boolean {
  return restartInProgress;
}

export async function checkForUpdates(manual = true): Promise<DesktopUpdateState> {
  if (!options) throw new Error("更新管理器尚未初始化");
  if (state.phase === "checking" || state.phase === "downloading") return state;
  emitState({ phase: "checking", error: undefined, progress: undefined });
  try {
    const requestUrl = new URL(`${updateBaseUrl()}/api/v1/updates/check`);
    requestUrl.searchParams.set("version", app.getVersion());
    requestUrl.searchParams.set("platform", process.platform);
    requestUrl.searchParams.set("arch", process.arch);
    const response = await net.fetch(requestUrl.toString(), { method: "GET" });
    if (!response.ok) throw new Error(`更新服务返回 ${response.status}`);
    remoteManifest = verifyRemoteUpdateManifest(await response.json(), await loadUpdatePublicKey());
    return emitState(availableState(remoteManifest));
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    updateLogger.warn("检查更新失败", { manual, message });
    return emitState({ phase: manual ? "error" : "idle", error: manual ? message : undefined });
  }
}

async function downloadFile(
  url: string,
  destination: string,
  expectedSize: number,
  onProgress: (progress: number) => void,
): Promise<void> {
  const parsedUrl = new URL(url);
  const localDevelopmentUrl = parsedUrl.hostname === "127.0.0.1" || parsedUrl.hostname === "localhost";
  if (parsedUrl.protocol !== "https:" && !(process.env.NODE_ENV !== "production" && localDevelopmentUrl)) {
    throw new Error("更新文件必须通过 HTTPS 下载");
  }
  const response = await net.fetch(url, { method: "GET" });
  if (!response.ok || !response.body) throw new Error(`补丁下载失败: ${response.status}`);
  if (new URL(response.url).protocol !== "https:" && !localDevelopmentUrl) {
    throw new Error("补丁下载重定向到了非 HTTPS 地址");
  }
  await fsp.mkdir(path.dirname(destination), { recursive: true });
  const temporary = `${destination}.tmp`;
  const output = await fsp.open(temporary, "w");
  const reader = response.body.getReader();
  let received = 0;
  let downloadError: unknown;
  try {
    while (true) {
      const { done, value } = await reader.read();
      if (done) break;
      if (!value) continue;
      received += value.byteLength;
      if (received > expectedSize) throw new Error("补丁下载大小超过清单声明");
      await output.write(value);
      onProgress((received / expectedSize) * 100);
    }
  } catch (error) {
    downloadError = error;
  } finally {
    await output.close();
  }
  if (downloadError) {
    await fsp.rm(temporary, { force: true });
    throw downloadError;
  }
  if (received !== expectedSize) {
    await fsp.rm(temporary, { force: true });
    throw new Error("补丁下载大小与清单不一致");
  }
  await fsp.rename(temporary, destination);
}

async function downloadHotPatch(): Promise<DesktopUpdateState> {
  if (!options || !remoteManifest?.hotPatch) throw new Error("没有可下载的热补丁");
  const descriptor = remoteManifest.hotPatch;
  const archivePath = path.join(options.userDataPath, "updates", "downloads", `${descriptor.id}.zip`);
  emitState({ phase: "downloading", downloadedKind: undefined, progress: 0, error: undefined });
  await downloadFile(descriptor.url, archivePath, descriptor.size, (progress) => {
    emitState({ phase: "downloading", progress: Math.max(0, Math.min(100, progress)) });
  });
  await installHotPatchArchive({
    archivePath,
    descriptor,
    currentVersion: app.getVersion(),
    userDataPath: options.userDataPath,
  });
  await fsp.rm(archivePath, { force: true });
  return emitState({
    phase: "downloaded",
    hotPatchAvailable: false,
    activeHotPatchId: descriptor.id,
    downloadedKind: "hotPatch",
    progress: 100,
  });
}

async function downloadInstaller(): Promise<DesktopUpdateState> {
  if (!remoteManifest?.installer) throw new Error("没有可下载的安装包更新");
  if (!app.isPackaged) throw new Error("开发模式不能安装整包更新");
  emitState({ phase: "downloading", downloadedKind: undefined, progress: 0, error: undefined });
  const checkResult = await autoUpdater.checkForUpdates();
  if (!checkResult || checkResult.updateInfo.version !== remoteManifest.version) {
    throw new Error("安装包版本与签名清单不一致");
  }
  const downloadedFiles = await autoUpdater.downloadUpdate();
  const downloadedFile = downloadedFiles[0];
  if (!downloadedFile) throw new Error("安装包下载结果为空");
  const actualHash = await sha256File(downloadedFile);
  if (actualHash.toLowerCase() !== remoteManifest.installer.sha256.toLowerCase()) {
    await fsp.rm(downloadedFile, { force: true });
    throw new Error("安装包哈希校验失败");
  }
  return emitState({ phase: "downloaded", downloadedKind: "installer", progress: 100 });
}

export async function downloadUpdate(kind: UpdateKind): Promise<DesktopUpdateState> {
  try {
    return kind === "hotPatch" ? await downloadHotPatch() : await downloadInstaller();
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    updateLogger.error("下载更新失败", { kind, message });
    return emitState({ phase: "error", error: message, progress: undefined });
  }
}

export async function applyDownloadedUpdate(): Promise<DesktopUpdateState> {
  if (!options || state.phase !== "downloaded" || !state.downloadedKind) {
    return emitState({ phase: "error", error: "没有已下载的更新" });
  }
  const kind = state.downloadedKind;
  emitState({ phase: "applying", error: undefined });
  try {
    await options.prepareToRestart();
    restartInProgress = true;
    if (kind === "installer") {
      autoUpdater.quitAndInstall(false, true);
    } else {
      app.relaunch();
      app.quit();
    }
    return state;
  } catch (error) {
    restartInProgress = false;
    const message = error instanceof Error ? error.message : String(error);
    return emitState({ phase: "error", error: message });
  }
}
