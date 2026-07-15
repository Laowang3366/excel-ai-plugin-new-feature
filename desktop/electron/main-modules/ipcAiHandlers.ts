import { trustedIpcMain as ipcMain } from "../shared/trustedIpc";
import { validateInput, AiListModelsInput, AiTestConnectionInput } from "../shared/ipcSchemas";
import { createLogger } from "../shared/logger";
import { getProviderApiKey } from "./settingsManager";
import { SETTINGS_SECRET_MASK } from "./settingsSecrets";
import { secureFetch } from "../shared/outboundUrlPolicy";
import { guardDataOperation } from "./dataMaintenance";

const logger = createLogger("IPC");

function resolveApiKey(apiKey: string, providerId?: string): string {
  if (apiKey !== SETTINGS_SECRET_MASK) return apiKey;
  if (!providerId) throw new Error("provider_secret_reference_required");
  const storedApiKey = getProviderApiKey(providerId);
  if (!storedApiKey) throw new Error("provider_secret_not_found");
  return storedApiKey;
}

export function registerAiIpcHandlers(isDataMaintenanceInProgress?: () => boolean): void {
  ipcMain.handle(
    "ai:listModels",
    guardDataOperation(
      isDataMaintenanceInProgress,
      async (
        _event,
        baseUrl: unknown,
        apiKey: unknown,
        apiFormat: unknown,
        providerId?: unknown,
      ) => {
        const validated = validateInput(AiListModelsInput, {
          baseUrl,
          apiKey,
          apiFormat,
          providerId,
        });
        try {
          const resolvedApiKey = resolveApiKey(validated.apiKey, validated.providerId);
          if (validated.apiFormat === "anthropic") return [];
          const url = validated.baseUrl.endsWith("/models")
            ? validated.baseUrl
            : `${validated.baseUrl.replace(/\/+$/, "")}/models`;
          const response = await secureFetch(url, {
            method: "GET",
            headers: {
              Authorization: `Bearer ${resolvedApiKey}`,
              "Content-Type": "application/json",
            },
            signal: AbortSignal.timeout(15000),
          });
          if (!response.ok) {
            logger.error(`[ai:listModels] HTTP ${response.status}`);
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
      },
    ),
  );

  ipcMain.handle(
    "ai:testConnection",
    guardDataOperation(
      isDataMaintenanceInProgress,
      async (
        _event,
        baseUrl: unknown,
        apiKey: unknown,
        apiFormat: unknown,
        model: unknown,
        providerId?: unknown,
      ) => {
        const validated = validateInput(AiTestConnectionInput, {
          baseUrl,
          apiKey,
          apiFormat,
          model,
          providerId,
        });
        const startTime = Date.now();
        try {
          const resolvedApiKey = resolveApiKey(validated.apiKey, validated.providerId);
          let url: string;
          let body: any;
          let headers: Record<string, string>;

          if (validated.apiFormat === "anthropic") {
            url = `${validated.baseUrl.replace(/\/+$/, "")}/messages`;
            headers = {
              "x-api-key": resolvedApiKey,
              "anthropic-version": "2023-06-01",
              "Content-Type": "application/json",
            };
            body = {
              model: validated.model || "claude-sonnet-4-20250514",
              max_tokens: 1,
              messages: [{ role: "user", content: "Hi" }],
            };
          } else if (validated.apiFormat === "responses") {
            url = `${validated.baseUrl.replace(/\/+$/, "")}/responses`;
            headers = {
              Authorization: `Bearer ${resolvedApiKey}`,
              "Content-Type": "application/json",
            };
            body = { model: validated.model || "gpt-4o", input: "Hi", max_output_tokens: 1 };
          } else {
            url = `${validated.baseUrl.replace(/\/+$/, "")}/chat/completions`;
            headers = {
              Authorization: `Bearer ${resolvedApiKey}`,
              "Content-Type": "application/json",
            };
            body = {
              model: validated.model || "gpt-4o",
              max_tokens: 1,
              messages: [{ role: "user", content: "Hi" }],
            };
          }

          const response = await secureFetch(url, {
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
      },
    ),
  );
}
