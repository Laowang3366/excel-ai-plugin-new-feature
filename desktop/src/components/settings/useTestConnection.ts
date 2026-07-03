/**
 * useTestConnection — 测试 AI 供应商连接的共享 hook
 *
 * 从 ProviderCard / AddProviderDialog 提取，统一：
 * - testing / testResult 状态管理
 * - ipcApi.ai.testConnection 调用
 * - 错误处理
 */

import { useState, useCallback } from "react";
import { ipcApi } from "../../services/ipcApi";

export interface TestResult {
  success: boolean;
  error?: string;
  latency?: number;
}

interface UseTestConnectionOptions {
  testFailedText: string;
}

export function useTestConnection({ testFailedText }: UseTestConnectionOptions) {
  const [testing, setTesting] = useState(false);
  const [testResult, setTestResult] = useState<TestResult | null>(null);

  const testConnection = useCallback(
    async (baseUrl: string, apiKey: string, apiFormat: string, model: string) => {
      setTesting(true);
      setTestResult(null);
      try {
        const result = await ipcApi.ai.testConnection(
          baseUrl,
          apiKey,
          apiFormat,
          model
        );
        setTestResult(result);
      } catch (err: any) {
        setTestResult({ success: false, error: err?.message || testFailedText });
      } finally {
        setTesting(false);
      }
    },
    [testFailedText]
  );

  return { testing, testResult, setTestResult, testConnection };
}
