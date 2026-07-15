import { useEffect, useMemo, useState } from "react";

import type { DesktopUpdateState, UpdateKind } from "../../electronApi";
import { ipcApi } from "../../services/ipcApi";
import { useSettingsStore } from "../../store/settingsStore";
import { CheckCircle, Download, RefreshCw } from "../common/IconMap";

const UPDATE_TEXT = {
  "zh-CN": {
    title: "软件更新",
    description: "获取新功能、体验改进与问题修复。",
    currentVersion: "当前版本",
    availableVersion: "最新版本",
    activePatch: "当前热补丁",
    check: "检查更新",
    checking: "正在检查",
    downloadInstaller: "下载并安装新版本",
    downloadPatch: "下载热补丁",
    restart: "重启并应用",
    downloading: "正在下载",
    applying: "正在应用更新",
    upToDate: "当前已是最新版本",
    available: "发现可用更新",
    idle: "应用会定期自动检查更新",
    downloaded: "更新已就绪，重启后生效",
    failed: "更新失败",
    releaseNotes: "更新日志",
    noReleaseNotes: "暂无新的功能更新说明。",
    automatic: "安装包更新会在应用内完成覆盖安装；热补丁下载后仅需重启应用。",
  },
  "en-US": {
    title: "Software updates",
    description: "Get new features, experience improvements, and fixes.",
    currentVersion: "Current version",
    availableVersion: "Latest version",
    activePatch: "Active hot patch",
    check: "Check for updates",
    checking: "Checking",
    downloadInstaller: "Download and install",
    downloadPatch: "Download hot patch",
    restart: "Restart and apply",
    downloading: "Downloading",
    applying: "Applying update",
    upToDate: "You are up to date",
    available: "An update is available",
    idle: "The app checks for updates periodically",
    downloaded: "Update ready. Restart to apply it.",
    failed: "Update failed",
    releaseNotes: "What's new",
    noReleaseNotes: "No new feature notes are available.",
    automatic: "Installer updates run in the app. Hot patches only require an app restart.",
  },
} as const;

function statusText(
  state: DesktopUpdateState,
  text: (typeof UPDATE_TEXT)["zh-CN"] | (typeof UPDATE_TEXT)["en-US"],
): string {
  if (state.phase === "checking") return text.checking;
  if (state.phase === "downloading")
    return `${text.downloading} ${Math.round(state.progress ?? 0)}%`;
  if (state.phase === "applying") return text.applying;
  if (state.phase === "downloaded") return text.downloaded;
  if (state.phase === "available") return text.available;
  if (state.phase === "up-to-date") return text.upToDate;
  if (state.phase === "error") return state.error || text.failed;
  return text.idle;
}

export function UpdateSettings() {
  const language = useSettingsStore((store) => store.language);
  const text = UPDATE_TEXT[language];
  const [updateState, setUpdateState] = useState<DesktopUpdateState | null>(null);
  const [busyKind, setBusyKind] = useState<UpdateKind | null>(null);

  useEffect(() => {
    let active = true;
    void ipcApi.update.getState().then((nextState) => {
      if (active) setUpdateState(nextState);
    });
    const unsubscribe = ipcApi.update.onStateChanged((nextState) => {
      if (active) setUpdateState(nextState);
    });
    return () => {
      active = false;
      unsubscribe();
    };
  }, []);

  const isBusy =
    updateState?.phase === "checking" ||
    updateState?.phase === "downloading" ||
    updateState?.phase === "applying";
  const notes = useMemo(() => updateState?.releaseNotes ?? [], [updateState?.releaseNotes]);

  const check = async () => {
    setUpdateState(await ipcApi.update.check(true));
  };

  const download = async (kind: UpdateKind) => {
    setBusyKind(kind);
    try {
      setUpdateState(await ipcApi.update.download(kind));
    } finally {
      setBusyKind(null);
    }
  };

  const apply = async () => {
    setUpdateState(await ipcApi.update.apply());
  };

  return (
    <div className="settings-section-content update-settings">
      <h2>{text.title}</h2>
      <p className="section-desc">{text.description}</p>

      <section className="settings-card update-overview">
        <div className="update-version-row">
          <div>
            <span>{text.currentVersion}</span>
            <strong>{updateState?.currentVersion ?? "..."}</strong>
          </div>
          {updateState?.availableVersion && (
            <div>
              <span>{text.availableVersion}</span>
              <strong>{updateState.availableVersion}</strong>
            </div>
          )}
          {updateState?.activeHotPatchId && (
            <div>
              <span>{text.activePatch}</span>
              <strong>{updateState.activeHotPatchId}</strong>
            </div>
          )}
        </div>

        <div className={`update-status update-status-${updateState?.phase ?? "idle"}`}>
          {updateState?.phase === "up-to-date" ? (
            <CheckCircle size={17} />
          ) : (
            <RefreshCw size={17} className={isBusy ? "spin" : ""} />
          )}
          <span>{updateState ? statusText(updateState, text) : text.idle}</span>
        </div>

        {updateState?.phase === "downloading" && (
          <div className="update-progress" aria-label={text.downloading}>
            <span style={{ width: `${Math.max(0, Math.min(100, updateState.progress ?? 0))}%` }} />
          </div>
        )}

        <div className="update-actions">
          <button
            type="button"
            className="btn-secondary update-check-btn"
            onClick={check}
            disabled={Boolean(isBusy)}
            aria-busy={updateState?.phase === "checking"}
          >
            <RefreshCw
              size={15}
              className={updateState?.phase === "checking" ? "spin" : undefined}
            />
            {updateState?.phase === "checking" ? text.checking : text.check}
          </button>
          {updateState?.hotPatchAvailable && (
            <button
              type="button"
              className="btn-primary"
              onClick={() => download("hotPatch")}
              disabled={Boolean(isBusy)}
            >
              <Download size={15} />
              {busyKind === "hotPatch" ? text.downloading : text.downloadPatch}
            </button>
          )}
          {updateState?.installerAvailable && (
            <button
              type="button"
              className="btn-primary"
              onClick={() => download("installer")}
              disabled={Boolean(isBusy)}
            >
              <Download size={15} />
              {busyKind === "installer" ? text.downloading : text.downloadInstaller}
            </button>
          )}
          {updateState?.phase === "downloaded" && (
            <button type="button" className="btn-primary" onClick={apply}>
              <RefreshCw size={15} />
              {text.restart}
            </button>
          )}
        </div>
        <p className="form-hint">{text.automatic}</p>
      </section>

      <section className="settings-card update-release-notes">
        <h3>{text.releaseNotes}</h3>
        {notes.length > 0 ? (
          <ul>
            {notes.map((note) => (
              <li key={note}>{note}</li>
            ))}
          </ul>
        ) : (
          <p>{text.noReleaseNotes}</p>
        )}
      </section>
    </div>
  );
}
