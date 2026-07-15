/**
 * useExcelConnection — Excel/WPS 连接状态管理
 *
 * 从 Sidebar.tsx 提取，管理：
 * - Excel 连接状态检测（自动每 30 秒轮询）
 * - 手动连接/重连
 * - 连接动画状态（脉冲/失败抖动）
 * - 多宿主选择弹窗（Office + WPS 同时运行时）
 */

import { useState, useEffect, useCallback, useRef } from "react";
import { ipcApi } from "../services/ipcApi";
import type { ExcelStatus } from "../utils/sidebarHelpers";

const INITIAL_STATUS: ExcelStatus = { connected: false, host: "unknown" };

export function useExcelConnection() {
  const [excelStatus, setExcelStatus] = useState<ExcelStatus>(INITIAL_STATUS);
  const [connecting, setConnecting] = useState(false);
  const [connectFailed, setConnectFailed] = useState(false);
  const [pulseDot, setPulseDot] = useState(false);
  const timeoutIdsRef = useRef<number[]>([]);
  /**
   * 待选择的可用宿主列表
   * 当 Office + WPS 同时运行时，detectStatus 返回 availableHosts，
   * UI 弹窗让用户选择后再调用 selectHost/connect。
   */
  const [pendingHosts, setPendingHosts] = useState<string[] | null>(null);

  const scheduleTimeout = useCallback((callback: () => void, delayMs: number) => {
    const timeoutId = window.setTimeout(() => {
      timeoutIdsRef.current = timeoutIdsRef.current.filter((id) => id !== timeoutId);
      callback();
    }, delayMs);
    timeoutIdsRef.current.push(timeoutId);
  }, []);

  const clearScheduledTimeouts = useCallback(() => {
    for (const timeoutId of timeoutIdsRef.current) {
      window.clearTimeout(timeoutId);
    }
    timeoutIdsRef.current = [];
  }, []);

  const detectExcel = useCallback(async () => {
    try {
      const status = await ipcApi.excel.detectStatus();
      setExcelStatus(status as ExcelStatus);
      // 多个宿主可用且未连接 → 弹出选择框
      if (
        !status.connected &&
        (status as any).availableHosts &&
        (status as any).availableHosts.length > 1
      ) {
        setPendingHosts((status as any).availableHosts);
      } else {
        // 单个宿主或已连接 → 关闭弹窗
        setPendingHosts(null);
      }
    } catch {
      setExcelStatus(INITIAL_STATUS);
    }
  }, []);

  useEffect(() => {
    detectExcel();
    const interval = setInterval(detectExcel, 30000);
    return () => clearInterval(interval);
  }, [detectExcel]);

  useEffect(() => clearScheduledTimeouts, [clearScheduledTimeouts]);

  const handleConnect = useCallback(async () => {
    setConnecting(true);
    setConnectFailed(false);
    try {
      const status = await ipcApi.excel.connect();
      setExcelStatus(status as ExcelStatus);
      if ((status as ExcelStatus).connected) {
        setPulseDot(true);
        scheduleTimeout(() => setPulseDot(false), 1500);
        setPendingHosts(null);
      } else if ((status as any).availableHosts?.length > 1) {
        // 连接返回多宿主 → 弹出选择框
        setPendingHosts((status as any).availableHosts);
        setConnectFailed(true);
        scheduleTimeout(() => setConnectFailed(false), 600);
      } else {
        setConnectFailed(true);
        scheduleTimeout(() => setConnectFailed(false), 600);
      }
    } catch {
      setExcelStatus(INITIAL_STATUS);
      setConnectFailed(true);
      scheduleTimeout(() => setConnectFailed(false), 600);
    } finally {
      setConnecting(false);
    }
  }, [scheduleTimeout]);

  /**
   * 用户从弹窗中选择目标宿主
   */
  const handleSelectHost = useCallback(
    async (host: "excel" | "wps") => {
      setConnecting(true);
      try {
        const status = await ipcApi.excel.selectHost(host);
        setExcelStatus(status as ExcelStatus);
        if (status.connected) {
          setPulseDot(true);
          scheduleTimeout(() => setPulseDot(false), 1500);
          setPendingHosts(null);
        } else {
          setConnectFailed(true);
          scheduleTimeout(() => setConnectFailed(false), 600);
        }
      } catch {
        setExcelStatus(INITIAL_STATUS);
        setConnectFailed(true);
        scheduleTimeout(() => setConnectFailed(false), 600);
      } finally {
        setConnecting(false);
      }
    },
    [scheduleTimeout],
  );

  return {
    excelStatus,
    connecting,
    connectFailed,
    pulseDot,
    pendingHosts,
    handleConnect,
    handleSelectHost,
    setPendingHosts,
  };
}
