/**
 * 常规设置 — 主题、语言、显示偏好、数据存储
 */

import React, { useEffect, useState } from "react";
import { Activity, FileScan, Maximize2 } from "../common/IconMap";
import { useSettingsStore, type AppLanguage, type AppTheme } from "../../store/settingsStore";
import { formatTokensAsK, DEFAULT_CONTEXT_WINDOW } from "../../utils/modelContextWindows";
import { ipcApi } from "../../services/ipcApi";
import { SettingsSliderField, SettingsSwitchField } from "./SettingsFields";
import { GENERAL_TEXT, getWindowOpacityText } from "./generalSettingsText";
import { GeneralSettingsStorageCard } from "./GeneralSettingsStorageCard";

export const GeneralSettings: React.FC = () => {
  const {
    language,
    theme,
    closeToTray,
    officeAutoCompactEnabled,
    dynamicArrayFunctionsEnabled,
    windowOpacity,
    compactionEnabled,
    autoCompactThresholdPercent,
    providers,
    activeProviderId,
    setLanguage,
    setTheme,
    setCloseToTray,
    setOfficeAutoCompactEnabled,
    setDynamicArrayFunctionsEnabled,
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
  const windowOpacityMin = 55;
  const windowOpacityMax = 100;
  const autoCompactThresholdMin = 10;
  const autoCompactThresholdMax = 95;
  const windowOpacityText = getWindowOpacityText(language);
  const windowOpacityLabel = windowOpacityText.label;
  const windowOpacityHint = windowOpacityText.hint;

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

        <SettingsSwitchField
          groupLabel={text.closeBehavior}
          label={text.closeToTray}
          checked={closeToTray}
          onChange={setCloseToTray}
          hint={text.closeHint}
        />
      </div>

      <div className="settings-card">
        <div className="settings-card-header">
          <div className="settings-card-title-row">
            <Maximize2 size={16} />
            <h3>{text.windowAvoidanceTitle}</h3>
          </div>
          <p>{text.windowAvoidanceDesc}</p>
        </div>

        <SettingsSwitchField
          label={text.officeAutoCompactEnabled}
          checked={officeAutoCompactEnabled}
          onChange={setOfficeAutoCompactEnabled}
          hint={text.officeAutoCompactHint}
        />

        <SettingsSliderField
          label={windowOpacityLabel}
          value={windowOpacityPercent}
          min={windowOpacityMin}
          max={windowOpacityMax}
          step={5}
          valueText={`${windowOpacityPercent}%`}
          onChange={(value) => setWindowOpacity(value / 100)}
          hint={windowOpacityHint}
        />

        <SettingsSwitchField
          groupLabel={text.dynamicArrayTitle}
          label={text.dynamicArrayFunctionsEnabled}
          checked={dynamicArrayFunctionsEnabled}
          onChange={setDynamicArrayFunctionsEnabled}
          hint={text.dynamicArrayFunctionsHint}
        />
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

      <GeneralSettingsStorageCard
        language={language}
        dataPath={dataPath}
        pathError={pathError}
        copied={copied}
        isMigrating={isMigrating}
        onOpenDataPath={handleOpenDataPath}
        onCopyDataPath={handleCopyDataPath}
        onChangeDataPath={handleChangeDataPath}
      />

      <div className="settings-card">
        <div className="settings-card-header">
          <div className="settings-card-title-row">
            <Activity size={16} />
            <h3>{text.contextTitle}</h3>
          </div>
          <p>{text.contextDesc}</p>
        </div>

        <SettingsSwitchField
          groupLabel={text.compactionEnabled}
          label={text.compactionEnabled}
          checked={compactionEnabled}
          onChange={setCompactionEnabled}
          hint={text.compactionEnabledHint}
        />

        <SettingsSliderField
          className={!compactionEnabled ? "form-group-disabled" : ""}
          label={text.compactionThreshold}
          value={autoCompactThresholdPercent}
          min={autoCompactThresholdMin}
          max={autoCompactThresholdMax}
          step={5}
          valueText={`${autoCompactThresholdPercent}%`}
          disabled={!compactionEnabled}
          onChange={setAutoCompactThresholdPercent}
          hint={text.compactionThresholdHint}
          info={(
          <div className="compaction-model-info">
            {text.currentModelContext}: <strong>{formatTokensAsK(currentContextWindow)}</strong>
            {currentModel && <span className="compaction-model-name">({currentModel})</span>}
          </div>
          )}
        />
      </div>
    </div>
  );
};
