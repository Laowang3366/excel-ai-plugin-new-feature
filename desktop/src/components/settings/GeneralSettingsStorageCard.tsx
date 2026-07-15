import React from "react";
import { USER_DATA_ERASE_CONFIRMATION } from "../../../electron/shared/userDataEraseContract";
import { Copy, Database, FolderOpen } from "../common/IconMap";
import type { GeneralSettingsLanguage } from "./generalSettingsText";
import { GENERAL_TEXT } from "./generalSettingsText";

interface GeneralSettingsStorageCardProps {
  language: GeneralSettingsLanguage;
  dataPath: string;
  pathError: string;
  copied: boolean;
  isMigrating: boolean;
  isExporting: boolean;
  exportMessage: string;
  eraseConfirmation: string;
  isErasing: boolean;
  eraseMessage: string;
  isRotatingKey: boolean;
  rotateKeyMessage: string;
  onOpenDataPath: () => void;
  onCopyDataPath: () => void;
  onChangeDataPath: () => void;
  onExportUserData: () => void;
  onEraseConfirmationChange: (value: string) => void;
  onEraseUserData: () => void;
  onRotateLocalDataKey: () => void;
}

export const GeneralSettingsStorageCard: React.FC<GeneralSettingsStorageCardProps> = ({
  language,
  dataPath,
  pathError,
  copied,
  isMigrating,
  isExporting,
  exportMessage,
  eraseConfirmation,
  isErasing,
  eraseMessage,
  isRotatingKey,
  rotateKeyMessage,
  onOpenDataPath,
  onCopyDataPath,
  onChangeDataPath,
  onExportUserData,
  onEraseConfirmationChange,
  onEraseUserData,
  onRotateLocalDataKey,
}) => {
  const text = GENERAL_TEXT[language];

  return (
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
            onClick={onOpenDataPath}
            disabled={!dataPath || Boolean(pathError)}
            title={text.openTitle}
          >
            <FolderOpen size={15} />
            {text.open}
          </button>
          <button
            className="settings-action-btn"
            onClick={onCopyDataPath}
            disabled={!dataPath || Boolean(pathError)}
            title={text.copyTitle}
          >
            <Copy size={15} />
            {copied ? text.copied : text.copy}
          </button>
          <button
            className="settings-action-btn primary"
            onClick={onChangeDataPath}
            disabled={isMigrating || isExporting || isErasing || isRotatingKey}
            title={text.changeTitle}
          >
            {isMigrating ? text.migrating : text.change}
          </button>
        </div>
        <span className="form-hint">{text.storageHint}</span>
      </div>

      <div className="form-group">
        <label>{text.exportData}</label>
        <button
          className="settings-action-btn"
          onClick={onExportUserData}
          disabled={isMigrating || isExporting || isErasing || isRotatingKey}
          title={text.exportTitle}
        >
          {isExporting ? text.exporting : text.exportData}
        </button>
        <span className="form-hint">{exportMessage || text.exportHint}</span>
      </div>

      <div className="form-group">
        <label>{text.rotateDataKey}</label>
        <button
          className="settings-action-btn"
          onClick={onRotateLocalDataKey}
          disabled={isMigrating || isExporting || isErasing || isRotatingKey}
          title={text.rotateDataKeyTitle}
        >
          {isRotatingKey ? text.rotatingDataKey : text.rotateDataKey}
        </button>
        <span className="form-hint">{rotateKeyMessage || text.rotateDataKeyHint}</span>
      </div>

      <div className="form-group">
        <label>{text.eraseData}</label>
        <input
          className="form-input"
          value={eraseConfirmation}
          onChange={(event) => onEraseConfirmationChange(event.target.value)}
          placeholder={text.eraseConfirmationPlaceholder}
          disabled={isMigrating || isExporting || isErasing || isRotatingKey}
          autoComplete="off"
        />
        <button
          className="settings-action-btn danger"
          onClick={onEraseUserData}
          disabled={
            isMigrating ||
            isExporting ||
            isErasing ||
            isRotatingKey ||
            eraseConfirmation !== USER_DATA_ERASE_CONFIRMATION
          }
          title={text.eraseTitle}
        >
          {isErasing ? text.erasing : text.eraseData}
        </button>
        <span className="form-hint">{eraseMessage || text.eraseHint}</span>
      </div>
    </div>
  );
};
