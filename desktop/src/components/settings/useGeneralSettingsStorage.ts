import { useEffect, useState } from "react";

import { USER_DATA_ERASE_CONFIRMATION } from "../../../electron/shared/userDataEraseContract";
import { ipcApi } from "../../services/ipcApi";
import type { GENERAL_TEXT } from "./generalSettingsText";

type GeneralSettingsText = (typeof GENERAL_TEXT)[keyof typeof GENERAL_TEXT];

export function useGeneralSettingsStorage(
  text: GeneralSettingsText,
  loadSettings: () => Promise<void>,
) {
  const [dataPath, setDataPath] = useState("");
  const [pathError, setPathError] = useState("");
  const [copied, setCopied] = useState(false);
  const [isMigrating, setIsMigrating] = useState(false);
  const [isExporting, setIsExporting] = useState(false);
  const [exportMessage, setExportMessage] = useState("");
  const [eraseConfirmation, setEraseConfirmation] = useState("");
  const [isErasing, setIsErasing] = useState(false);
  const [eraseMessage, setEraseMessage] = useState("");

  useEffect(() => {
    let canceled = false;
    void ipcApi.app.getDataPath()
      .then((value) => { if (!canceled) setDataPath(value || text.unsupportedPath); })
      .catch((error) => {
        if (!canceled) setPathError(error instanceof Error ? error.message : text.readPathFailed);
      });
    return () => { canceled = true; };
  }, [text.readPathFailed, text.unsupportedPath]);

  const onOpenDataPath = async () => {
    if (!dataPath || pathError) return;
    const result = await ipcApi.app.openPath(dataPath);
    if (result) setPathError(result);
  };
  const onCopyDataPath = async () => {
    if (!dataPath || pathError) return;
    await navigator.clipboard.writeText(dataPath);
    setCopied(true);
    window.setTimeout(() => setCopied(false), 1600);
  };
  const onChangeDataPath = async () => {
    setPathError("");
    const selectedPath = (await ipcApi.app.selectDataPath())?.filePaths?.[0];
    if (!selectedPath) return;
    setIsMigrating(true);
    try {
      const result = await ipcApi.app.migrateDataPath(selectedPath);
      if (!result?.success) return setPathError(result?.error || text.migrateFailed);
      setDataPath(result.dataPath || selectedPath);
      await loadSettings();
    } catch (error) {
      setPathError(error instanceof Error ? error.message : text.migrateFailed);
    } finally {
      setIsMigrating(false);
    }
  };
  const onExportUserData = async () => {
    setExportMessage("");
    const selectedPath = (await ipcApi.app.selectExportPath())?.filePaths?.[0];
    if (!selectedPath) return;
    setIsExporting(true);
    try {
      const result = await ipcApi.app.exportUserData(selectedPath);
      setExportMessage(result.success
        ? `${text.exportSuccess}: ${result.exportPath || selectedPath}`
        : result.error || text.exportFailed);
    } catch (error) {
      setExportMessage(error instanceof Error ? error.message : text.exportFailed);
    } finally {
      setIsExporting(false);
    }
  };
  const onEraseUserData = async () => {
    if (eraseConfirmation !== USER_DATA_ERASE_CONFIRMATION) return;
    setIsErasing(true);
    setEraseMessage("");
    try {
      const result = await ipcApi.app.eraseUserData({ confirmation: eraseConfirmation });
      setEraseMessage(result.success ? text.eraseSuccess : result.error || text.eraseFailed);
      if (result.success || result.erasedCategories.length > 0) {
        setEraseConfirmation("");
        await loadSettings();
        if (result.success) window.setTimeout(() => window.location.reload(), 800);
      }
    } catch (error) {
      setEraseMessage(error instanceof Error ? error.message : text.eraseFailed);
    } finally {
      setIsErasing(false);
    }
  };

  return {
    dataPath, pathError, copied, isMigrating, isExporting, exportMessage,
    eraseConfirmation, isErasing, eraseMessage,
    onOpenDataPath, onCopyDataPath, onChangeDataPath, onExportUserData,
    onEraseConfirmationChange: setEraseConfirmation, onEraseUserData,
  };
}
