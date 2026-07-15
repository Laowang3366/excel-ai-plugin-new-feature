/**
 * 常规设置 — 主题、语言、显示偏好、数据存储
 */

import React, { useEffect, useState } from "react";
import { Activity, FileScan, Maximize2, Shield } from "../common/IconMap";
import { useSettingsStore, type AppLanguage, type AppTheme } from "../../store/settingsStore";
import { formatTokensAsK, DEFAULT_CONTEXT_WINDOW } from "../../utils/modelContextWindows";
import { ipcApi } from "../../services/ipcApi";
import { SettingsSliderField, SettingsSwitchField } from "./SettingsFields";
import { GENERAL_TEXT } from "./generalSettingsText";
import { GeneralSettingsStorageCard } from "./GeneralSettingsStorageCard";
import { useGeneralSettingsStorage } from "./useGeneralSettingsStorage";

export const GeneralSettings: React.FC = () => {
  const {
    language,
    theme,
    closeToTray,
    officeAutoCompactEnabled,
    dynamicArrayFunctionsEnabled,
    remoteDataProcessingEnabled,
    compactionEnabled,
    autoCompactThresholdPercent,
    providers,
    activeProviderId,
    setLanguage,
    setTheme,
    setCloseToTray,
    setOfficeAutoCompactEnabled,
    setDynamicArrayFunctionsEnabled,
    setRemoteDataProcessingEnabled,
    setCompactionEnabled,
    setAutoCompactThresholdPercent,
    loadSettings,
  } = useSettingsStore();
  const [mineruApiToken, setMineruApiToken] = useState("");
  const [mineruSaved, setMineruSaved] = useState(false);
  const text = GENERAL_TEXT[language];
  const storage = useGeneralSettingsStorage(text, loadSettings);
  const autoCompactThresholdMin = 10;
  const autoCompactThresholdMax = 95;

  // 获取当前供应商的上下文窗口大小（用户自定义，支持 per-model 覆盖）
  const activeProvider = providers[activeProviderId];
  const currentModel = activeProvider?.model || activeProvider?.defaultModel || "";
  const activeModelConfig = activeProvider?.modelConfigs?.find((m) => m.name === currentModel);
  const currentContextWindow =
    activeModelConfig?.contextWindowSize ||
    activeProvider?.contextWindowSize ||
    DEFAULT_CONTEXT_WINDOW;

  useEffect(() => {
    let canceled = false;

    const loadLocalSettings = async () => {
      const token = await ipcApi.settings.get("mineruApiToken");
      if (!canceled && typeof token === "string") {
        setMineruApiToken(token);
      }
    };

    loadLocalSettings();
    return () => {
      canceled = true;
    };
  }, []);

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
          <div className="settings-card-title-row">
            <Shield size={16} />
            <h3>{text.remoteDataTitle}</h3>
          </div>
          <p>{text.remoteDataDesc}</p>
        </div>

        <SettingsSwitchField
          groupLabel={text.remoteDataTitle}
          label={text.remoteDataProcessingEnabled}
          checked={remoteDataProcessingEnabled}
          onChange={setRemoteDataProcessingEnabled}
          hint={
            remoteDataProcessingEnabled ? text.remoteDataEnabledHint : text.remoteDataDisabledHint
          }
        />
      </div>

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
          <span className="form-hint">{mineruSaved ? text.saved : text.mineruApiTokenHint}</span>
        </div>
      </div>

      <GeneralSettingsStorageCard language={language} {...storage} />

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
          info={
            <div className="compaction-model-info">
              {text.currentModelContext}: <strong>{formatTokensAsK(currentContextWindow)}</strong>
              {currentModel && <span className="compaction-model-name">({currentModel})</span>}
            </div>
          }
        />
      </div>
    </div>
  );
};
