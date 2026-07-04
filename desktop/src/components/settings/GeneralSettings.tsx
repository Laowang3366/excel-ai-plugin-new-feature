/**
 * 常规设置 — 主题、语言、显示偏好、数据存储
 */

import React, { useEffect, useState } from "react";
import { Copy, Database, FolderOpen, Activity, FileScan, Maximize2 } from "../common/IconMap";
import { useSettingsStore, type AppLanguage, type AppTheme } from "../../store/settingsStore";
import { formatTokensAsK, DEFAULT_CONTEXT_WINDOW } from "../../utils/modelContextWindows";
import { ipcApi } from "../../services/ipcApi";

const GENERAL_TEXT = {
  "zh-CN": {
    title: "常规设置",
    desc: "应用的基本显示和行为偏好。",
    displayTitle: "显示偏好",
    displayDesc: "控制应用界面语言、主题等基础显示方式。",
    language: "界面语言",
    languageZh: "简体中文",
    languageEn: "English",
    languageHint: "切换后会立即保存到本机设置。",
    theme: "主题",
    light: "浅色模式",
    dark: "深色模式",
    themeHint: "深色模式会立即应用到当前窗口。",
    closeBehavior: "窗口关闭行为",
    closeToTray: "关闭窗口时隐藏到系统托盘",
    closeHint: "开启后点击右上角关闭按钮不会退出应用，可从托盘图标恢复或退出。",
    windowAvoidanceTitle: "窗口避让",
    windowAvoidanceDesc: "控制助手在 Office 操作时的占屏方式。",
    officeAutoCompactEnabled: "Office 操作时自动避让",
    officeAutoCompactHint: "当检测到 Office 已连接并且助手窗口失焦时，自动缩为右侧紧凑栏，减少对表格、文档或幻灯片的遮挡。",
    storageTitle: "数据存储",
    storageDesc: "会话历史、模型配置、权限偏好等数据保存在本机目录。",
    storagePath: "数据存储路径",
    loadingPath: "正在读取...",
    unsupportedPath: "当前环境暂不支持读取数据路径",
    readPathFailed: "读取数据路径失败",
    migrateFailed: "迁移数据失败",
    open: "打开",
    copy: "复制",
    copied: "已复制",
    change: "更换",
    migrating: "迁移中...",
    openTitle: "打开数据目录",
    copyTitle: "复制数据路径",
    changeTitle: "更换数据存储目录并迁移数据",
    storageHint: "更换目录会自动复制当前设置和会话历史，旧目录不会被删除。",
    contextTitle: "上下文管理",
    contextDesc: "控制对话上下文的自动压缩行为，避免长对话超出模型处理能力。",
    compactionEnabled: "启用自动压缩",
    compactionEnabledHint: "当上下文使用量达到阈值百分比时，自动压缩历史记录以释放上下文空间。",
    compactionThreshold: "自动压缩阈值",
    compactionThresholdHint: "当上下文使用量达到模型窗口的此百分比时触发自动压缩。较低的值压缩更频繁，较高的值保留更多历史。",
    currentModelContext: "当前模型上下文窗口",
    ocrTitle: "OCR 服务",
    ocrDesc: "配置 MinerU 通用解析接口，用于 OCR 和发票识别。",
    mineruApiToken: "MinerU API Token",
    mineruApiTokenPlaceholder: "粘贴 MinerU 控制台生成的 API Token",
    mineruApiTokenHint: "保存后 OCR 会优先使用 MinerU 通用解析；未配置或解析失败时会回退当前视觉模型。",
    saved: "已保存",
  },
  "en-US": {
    title: "General",
    desc: "Basic display and behavior preferences.",
    displayTitle: "Display",
    displayDesc: "Control language, theme, and basic visual preferences.",
    language: "Language",
    languageZh: "Simplified Chinese",
    languageEn: "English",
    languageHint: "Changes are saved locally immediately.",
    theme: "Theme",
    light: "Light",
    dark: "Dark",
    themeHint: "Dark mode is applied to the current window immediately.",
    closeBehavior: "Window close behavior",
    closeToTray: "Hide to system tray when closing the window",
    closeHint: "When enabled, the close button hides the app. Restore or quit from the tray icon.",
    windowAvoidanceTitle: "Window Avoidance",
    windowAvoidanceDesc: "Control how much screen space the assistant uses while working in Office.",
    officeAutoCompactEnabled: "Auto-avoid while using Office",
    officeAutoCompactHint: "When Office is connected and the assistant loses focus, it shrinks into a right-side compact panel to reduce obstruction.",
    storageTitle: "Data storage",
    storageDesc: "Conversation history, model settings, and preferences are stored locally.",
    storagePath: "Data storage path",
    loadingPath: "Loading...",
    unsupportedPath: "Data path is not available in this environment",
    readPathFailed: "Failed to read data path",
    migrateFailed: "Failed to migrate data",
    open: "Open",
    copy: "Copy",
    copied: "Copied",
    change: "Change",
    migrating: "Migrating...",
    openTitle: "Open data folder",
    copyTitle: "Copy data path",
    changeTitle: "Change data folder and migrate data",
    storageHint: "Changing the folder copies current settings and conversations. The old folder is not deleted.",
    contextTitle: "Context Management",
    contextDesc: "Control auto-compaction of conversation context to avoid exceeding model limits.",
    compactionEnabled: "Enable auto-compaction",
    compactionEnabledHint: "When context usage reaches the threshold percentage, history is automatically compacted to free context space.",
    compactionThreshold: "Auto-compact threshold",
    compactionThresholdHint: "Auto-compaction triggers when context usage reaches this percentage of the model's window. Lower values compact more frequently; higher values retain more history.",
    currentModelContext: "Current model context window",
    ocrTitle: "OCR Service",
    ocrDesc: "Configure MinerU's general parsing API for OCR and invoice recognition.",
    mineruApiToken: "MinerU API Token",
    mineruApiTokenPlaceholder: "Paste the API token generated in MinerU console",
    mineruApiTokenHint: "OCR uses MinerU first after saving. If unset or parsing fails, it falls back to the current vision model.",
    saved: "Saved",
  },
} as const;

export const GeneralSettings: React.FC = () => {
  const {
    language,
    theme,
    closeToTray,
    officeAutoCompactEnabled,
    windowOpacity,
    compactionEnabled,
    autoCompactThresholdPercent,
    providers,
    activeProviderId,
    setLanguage,
    setTheme,
    setCloseToTray,
    setOfficeAutoCompactEnabled,
    setWindowOpacity,
    setCompactionEnabled,
    setAutoCompactThresholdPercent,
    loadSettings,
  } = useSettingsStore();
  const [dataPath, setDataPath] = useState("");
  const [pathError, setPathError] = useState("");
  const [copied, setCopied] = useState(false);
  const [isMigrating, setIsMigrating] = useState(false);
  const [mineruApiToken, setMineruApiToken] = useState("");
  const [mineruSaved, setMineruSaved] = useState(false);
  const text = GENERAL_TEXT[language];
  const windowOpacityPercent = Math.round(windowOpacity * 100);
  const windowOpacityLabel = language === "zh-CN" ? "窗口透明度" : "Window opacity";
  const windowOpacityHint = language === "zh-CN"
    ? "降低后整个助手窗口会半透明，便于查看和操作后方的 Office 内容。"
    : "Lower values make the whole assistant window translucent so Office content behind it remains easier to use.";

  // 获取当前供应商的上下文窗口大小（用户自定义，支持 per-model 覆盖）
  const activeProvider = providers[activeProviderId];
  const currentModel = activeProvider?.model || activeProvider?.defaultModel || "";
  const activeModelConfig = activeProvider?.modelConfigs?.find(m => m.name === currentModel);
  const currentContextWindow = activeModelConfig?.contextWindowSize || activeProvider?.contextWindowSize || DEFAULT_CONTEXT_WINDOW;

  useEffect(() => {
    let canceled = false;

    const loadLocalSettings = async () => {
      try {
        const path = await ipcApi.app.getDataPath();
        if (!canceled) setDataPath(path || text.unsupportedPath);
      } catch (error) {
        if (!canceled) {
          setPathError(error instanceof Error ? error.message : text.readPathFailed);
        }
      }

      const token = await ipcApi.settings.get("mineruApiToken");
      if (!canceled && typeof token === "string") {
        setMineruApiToken(token);
      }
    };

    loadLocalSettings();
    return () => {
      canceled = true;
    };
  }, [text.readPathFailed, text.unsupportedPath]);

  const handleOpenDataPath = async () => {
    if (!dataPath || pathError) return;
    const result = await ipcApi.app.openPath(dataPath);
    if (result) setPathError(result);
  };

  const handleCopyDataPath = async () => {
    if (!dataPath || pathError) return;
    await navigator.clipboard.writeText(dataPath);
    setCopied(true);
    window.setTimeout(() => setCopied(false), 1600);
  };

  const handleChangeDataPath = async () => {
    setPathError("");
    const selection = await ipcApi.app.selectDataPath();
    const selectedPath = selection?.filePaths?.[0];
    if (!selectedPath) return;

    setIsMigrating(true);
    try {
      const result = await ipcApi.app.migrateDataPath(selectedPath);
      if (!result?.success) {
        setPathError(result?.error || text.migrateFailed);
        return;
      }
      setDataPath(result.dataPath || selectedPath);
      await loadSettings();
    } finally {
      setIsMigrating(false);
    }
  };

  const handleMineruTokenChange = async (value: string) => {
    setMineruApiToken(value);
    setMineruSaved(false);
    await ipcApi.settings.set("mineruApiToken", value.trim());
    setMineruSaved(true);
    window.setTimeout(() => setMineruSaved(false), 1600);
  };

  return (
    <div className="settings-section-content">
      <h2>{text.title}</h2>
      <p className="section-desc">{text.desc}</p>

      <div className="settings-card">
        <div className="settings-card-header">
          <div>
            <h3>{text.displayTitle}</h3>
            <p>{text.displayDesc}</p>
          </div>
        </div>

        <div className="form-group">
          <label>{text.language}</label>
          <select
            className="form-input"
            value={language}
            onChange={(event) => setLanguage(event.target.value as AppLanguage)}
          >
            <option value="zh-CN">{text.languageZh}</option>
            <option value="en-US">{text.languageEn}</option>
          </select>
          <span className="form-hint">{text.languageHint}</span>
        </div>

        <div className="form-group">
          <label>{text.theme}</label>
          <select
            className="form-input"
            value={theme}
            onChange={(event) => setTheme(event.target.value as AppTheme)}
          >
            <option value="light">{text.light}</option>
            <option value="dark">{text.dark}</option>
          </select>
          <span className="form-hint">{text.themeHint}</span>
        </div>

        <div className="form-group">
          <label>{text.closeBehavior}</label>
          <label className="settings-switch-row">
            <input
              type="checkbox"
              checked={closeToTray}
              onChange={(event) => setCloseToTray(event.target.checked)}
            />
            <span>{text.closeToTray}</span>
          </label>
          <span className="form-hint">{text.closeHint}</span>
        </div>
      </div>

      <div className="settings-card">
        <div className="settings-card-header">
          <div className="settings-card-title-row">
            <Maximize2 size={16} />
            <h3>{text.windowAvoidanceTitle}</h3>
          </div>
          <p>{text.windowAvoidanceDesc}</p>
        </div>

        <div className="form-group">
          <label className="settings-switch-row">
            <input
              type="checkbox"
              checked={officeAutoCompactEnabled}
              onChange={(event) => setOfficeAutoCompactEnabled(event.target.checked)}
            />
            <span>{text.officeAutoCompactEnabled}</span>
          </label>
          <span className="form-hint">{text.officeAutoCompactHint}</span>
        </div>

        <div className="form-group">
          <label>{windowOpacityLabel}</label>
          <div className="compaction-threshold-row">
            <input
              type="range"
              className="compaction-slider"
              min={55}
              max={100}
              step={5}
              value={windowOpacityPercent}
              onChange={(event) => setWindowOpacity(Number(event.target.value) / 100)}
            />
            <span className="compaction-threshold-value">
              {windowOpacityPercent}%
            </span>
          </div>
          <span className="form-hint">{windowOpacityHint}</span>
        </div>
      </div>

      <div className="settings-card">
        <div className="settings-card-header">
          <div className="settings-card-title-row">
            <FileScan size={16} />
            <h3>{text.ocrTitle}</h3>
          </div>
          <p>{text.ocrDesc}</p>
        </div>

        <div className="form-group">
          <label>{text.mineruApiToken}</label>
          <input
            className="form-input"
            type="password"
            value={mineruApiToken}
            onChange={(event) => handleMineruTokenChange(event.target.value)}
            placeholder={text.mineruApiTokenPlaceholder}
            autoComplete="off"
          />
          <span className="form-hint">
            {mineruSaved ? text.saved : text.mineruApiTokenHint}
          </span>
        </div>
      </div>

      <div className="settings-card">
        <div className="settings-card-header">
          <div className="settings-card-title-row">
            <Database size={16} />
            <h3>{text.storageTitle}</h3>
          </div>
          <p>{text.storageDesc}</p>
        </div>

        <div className="form-group">
          <label>{text.storagePath}</label>
          <div className="storage-path-row">
            <input
              className="form-input storage-path-input"
              value={pathError || dataPath || text.loadingPath}
              readOnly
            />
            <button
              className="settings-action-btn"
              onClick={handleOpenDataPath}
              disabled={!dataPath || Boolean(pathError)}
              title={text.openTitle}
            >
              <FolderOpen size={15} />
              {text.open}
            </button>
            <button
              className="settings-action-btn"
              onClick={handleCopyDataPath}
              disabled={!dataPath || Boolean(pathError)}
              title={text.copyTitle}
            >
              <Copy size={15} />
              {copied ? text.copied : text.copy}
            </button>
            <button
              className="settings-action-btn primary"
              onClick={handleChangeDataPath}
              disabled={isMigrating}
              title={text.changeTitle}
            >
              {isMigrating ? text.migrating : text.change}
            </button>
          </div>
          <span className="form-hint">{text.storageHint}</span>
        </div>
      </div>

      {/* 上下文管理卡片（参考 Codex auto_compact_token_limit） */}
      <div className="settings-card">
        <div className="settings-card-header">
          <div className="settings-card-title-row">
            <Activity size={16} />
            <h3>{text.contextTitle}</h3>
          </div>
          <p>{text.contextDesc}</p>
        </div>

        <div className="form-group">
          <label>{text.compactionEnabled}</label>
          <label className="settings-switch-row">
            <input
              type="checkbox"
              checked={compactionEnabled}
              onChange={(event) => setCompactionEnabled(event.target.checked)}
            />
            <span>{text.compactionEnabled}</span>
          </label>
          <span className="form-hint">{text.compactionEnabledHint}</span>
        </div>

        <div className={`form-group ${!compactionEnabled ? "form-group-disabled" : ""}`}>
          <label>{text.compactionThreshold}</label>
          {/* 显示当前模型上下文窗口 */}
          <div className="compaction-model-info">
            {text.currentModelContext}: <strong>{formatTokensAsK(currentContextWindow)}</strong>
            {currentModel && <span className="compaction-model-name">({currentModel})</span>}
          </div>
          <div className="compaction-threshold-row">
            <input
              type="range"
              className="compaction-slider"
              min={10}
              max={95}
              step={5}
              value={autoCompactThresholdPercent}
              onChange={(event) => setAutoCompactThresholdPercent(Number(event.target.value))}
              disabled={!compactionEnabled}
            />
            <span className="compaction-threshold-value">
              {autoCompactThresholdPercent}%
            </span>
          </div>
          <span className="form-hint">{text.compactionThresholdHint}</span>
        </div>
      </div>
    </div>
  );
};
