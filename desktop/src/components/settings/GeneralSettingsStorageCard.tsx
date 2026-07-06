import React from "react";
import { Copy, Database, FolderOpen } from "../common/IconMap";
import type { GeneralSettingsLanguage } from "./generalSettingsText";
import { GENERAL_TEXT } from "./generalSettingsText";

interface GeneralSettingsStorageCardProps {
  language: GeneralSettingsLanguage;
  dataPath: string;
  pathError: string;
  copied: boolean;
  isMigrating: boolean;
  onOpenDataPath: () => void;
  onCopyDataPath: () => void;
  onChangeDataPath: () => void;
}

export const GeneralSettingsStorageCard: React.FC<GeneralSettingsStorageCardProps> = ({
  language,
  dataPath,
  pathError,
  copied,
  isMigrating,
  onOpenDataPath,
  onCopyDataPath,
  onChangeDataPath,
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
            disabled={isMigrating}
            title={text.changeTitle}
          >
            {isMigrating ? text.migrating : text.change}
          </button>
        </div>
        <span className="form-hint">{text.storageHint}</span>
      </div>
    </div>
  );
};
