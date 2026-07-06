import { ipcMain } from "electron";
import {
  validateInput,
  AiListModelsInput,
  AiTestConnectionInput,
} from "../shared/ipcSchemas";
import { createLogger } from "../shared/logger";

const logger = createLogger("IPC");

export function registerAiIpcHandlers(): void {
  ipcMain.handle("ai:listModels", async (_event, baseUrl: unknown, apiKey: unknown, apiFormat: unknown) => {
    const validated = validateInput(AiListModelsInput, { baseUrl, apiKey, apiFormat });
    try {
      if (validated.apiFormat === "anthropic") return [];
      const url = validated.baseUrl.endsWith("/models")
        ? validated.baseUrl
        : `${validated.baseUrl.replace(/\/+$/, "")}/models`;
      const response = await fetch(url, {
        method: "GET",
        headers: {
          "Authorization": `Bearer ${validated.apiKey}`,
          "Content-Type": "application/json",
        },
        signal: AbortSignal.timeout(15000),
      });
      if (!response.ok) {
        const errorText = await response.text().catch(() => "");
        logger.error(`[ai:listModels] HTTP ${response.status}: ${errorText}`);
        return [];
      }
      const data = (await response.json()) as any;
      if (Array.isArray(data?.data)) {
        return data.data
          .map((m: any) => m.id || m.name || "")
          .filter((id: string) => id.length > 0)
          .sort();
      }
      if (Array.isArray(data)) {
        return data
          .map((m: any) => (typeof m === "string" ? m : m.id || m.name || ""))
          .filter((id: string) => id.length > 0)
          .sort();
      }
      return [];
    } catch (err: any) {
      logger.error("[ai:listModels] Error:", err?.message || err);
      return [];
    }
  });

  ipcMain.handle("ai:testConnection", async (_event, baseUrl: unknown, apiKey: unknown, apiFormat: unknown, model: unknown) => {
    const validated = validateInput(AiTestConnectionInput, { baseUrl, apiKey, apiFormat, model });
    const startTime = Date.now();
    try {
      let url: string;
      let body: any;
      let headers: Record<string, string>;

      if (validated.apiFormat === "anthropic") {
        url = `${validated.baseUrl.replace(/\/+$/, "")}/messages`;
        headers = {
          "x-api-key": validated.apiKey,
          "anthropic-version": "2023-06-01",
          "Content-Type": "application/json",
        };
        body = { model: validated.model || "claude-sonnet-4-20250514", max_tokens: 1, messages: [{ role: "user", content: "Hi" }] };
      } else if (validated.apiFormat === "responses") {
        url = `${validated.baseUrl.replace(/\/+$/, "")}/responses`;
        headers = {
          "Authorization": `Bearer ${validated.apiKey}`,
          "Content-Type": "application/json",
        };
        body = { model: validated.model || "gpt-4o", input: "Hi", max_output_tokens: 1 };
      } else {
        url = `${validated.baseUrl.replace(/\/+$/, "")}/chat/completions`;
        headers = {
          "Authorization": `Bearer ${validated.apiKey}`,
          "Content-Type": "application/json",
        };
        body = { model: validated.model || "gpt-4o", max_tokens: 1, messages: [{ role: "user", content: "Hi" }] };
      }

      const response = await fetch(url, {
        method: "POST",
        headers,
        body: JSON.stringify(body),
        signal: AbortSignal.timeout(30000),
      });

      const latency = Date.now() - startTime;
      if (response.ok) {
        return { success: true, latency };
      }

      const errorText = await response.text().catch(() => "");
      let errorMessage = `HTTP ${response.status}`;
      try {
        const errorJson = JSON.parse(errorText);
        errorMessage = errorJson.error?.message || errorJson.message || errorMessage;
      } catch {
        // use status fallback
      }
      return { success: false, error: errorMessage, latency };
    } catch (err: any) {
      const latency = Date.now() - startTime;
      return { success: false, error: err?.message || "连接失败", latency };
    }
  });
}
