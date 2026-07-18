/**
 * useOfficeConnection — Word/PPT 连接状态管理
 *
 * 与 useExcelConnection 配合，分别管理 Word 和 PPT 的连接检测，
 * 在前端侧边栏展示三个应用各自的连接状态。
 */

import { useState, useEffect, useCallback, useRef } from "react";
import { ipcApi } from "../services/ipcApi";

export interface OfficeAppStatus {
  connected: boolean;
  host: string;
  version?: string;
  documentName?: string;
  presentationName?: string;
  processId?: number;
  instanceId?: string;
}

const INITIAL_STATUS: OfficeAppStatus = { connected: false, host: "unknown" };
const POLL_INTERVAL_MS = 5000; // 5 秒轮询一次，比 Excel 的 30 秒更快

export function useOfficeConnection() {
  const [wordStatus, setWordStatus] = useState<OfficeAppStatus>(INITIAL_STATUS);
  const [presentationStatus, setPresentationStatus] = useState<OfficeAppStatus>(INITIAL_STATUS);
  const wordRef = useRef(wordStatus);
  const pptRef = useRef(presentationStatus);

  const detectWord = useCallback(async () => {
    try {
      const status = await ipcApi.office.detectWordStatus();
      const s = status as OfficeAppStatus;
      // 只在状态有变化时 setState，避免无效渲染
      if (!sameStatus(s, wordRef.current)) {
        wordRef.current = s;
        setWordStatus(s);
      }
    } catch {
      if (wordRef.current.connected) {
        wordRef.current = INITIAL_STATUS;
        setWordStatus(INITIAL_STATUS);
      }
    }
  }, []);

  const detectPresentation = useCallback(async () => {
    try {
      const status = await ipcApi.office.detectPresentationStatus();
      const s = status as OfficeAppStatus;
      if (!sameStatus(s, pptRef.current)) {
        pptRef.current = s;
        setPresentationStatus(s);
      }
    } catch {
      if (pptRef.current.connected) {
        pptRef.current = INITIAL_STATUS;
        setPresentationStatus(INITIAL_STATUS);
      }
    }
  }, []);

  const detectAll = useCallback(async () => {
    await Promise.allSettled([detectWord(), detectPresentation()]);
  }, [detectWord, detectPresentation]);

  useEffect(() => {
    detectAll();
    const interval = setInterval(detectAll, POLL_INTERVAL_MS);
    return () => clearInterval(interval);
  }, [detectAll]);

  return {
    wordStatus,
    presentationStatus,
    detectAll,
  };
}

function sameStatus(left: OfficeAppStatus, right: OfficeAppStatus): boolean {
  return (
    left.connected === right.connected &&
    left.host === right.host &&
    left.version === right.version &&
    left.documentName === right.documentName &&
    left.presentationName === right.presentationName &&
    left.processId === right.processId &&
    left.instanceId === right.instanceId
  );
}
