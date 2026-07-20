import { useMemo, useState } from "react";
import {
  getProviderTemplate,
  PROVIDER_TEMPLATES,
  ProviderClient,
  type ConnectionMode,
  type ProviderPublicView,
  type ProviderStore,
} from "@shared/provider";
import { ProviderCard, type ProviderCardEditState } from "./ProviderCard";
import { ProviderCreateSection } from "./ProviderCreateSection";

interface Props {
  store: ProviderStore;
}

type ActionState =
  | { status: "idle" }
  | { status: "loading"; action: string }
  | { status: "success"; message: string }
  | { status: "error"; message: string };

export function ProviderSettingsPanel({ store }: Props) {
  const client = useMemo(() => new ProviderClient(), []);
  const [providers, setProviders] = useState<ProviderPublicView[]>(() => store.list());
  const [activeId, setActiveId] = useState<string | null>(() => store.getActiveId());
  const [templateId, setTemplateId] = useState(PROVIDER_TEMPLATES[0]?.id ?? "openai");
  const [apiKey, setApiKey] = useState("");
  const [connectionMode, setConnectionMode] = useState<ConnectionMode>("direct");
  const [gatewayBaseUrl, setGatewayBaseUrl] = useState("");
  const [gatewayUpstreamId, setGatewayUpstreamId] = useState("openai");
  const [editId, setEditId] = useState<string | null>(null);
  const [edit, setEdit] = useState<ProviderCardEditState | null>(null);
  const [actionState, setActionState] = useState<ActionState>({ status: "idle" });
  const [listedModels, setListedModels] = useState<string[]>([]);

  function refresh() {
    setProviders(store.list());
    setActiveId(store.getActiveId());
  }

  function addProvider() {
    store.addFromTemplate(templateId, connectionMode === "direct" ? apiKey : "", {
      connectionMode,
      gatewayBaseUrl,
      gatewayUpstreamId,
    });
    setApiKey("");
    refresh();
  }

  function startEdit(provider: ProviderPublicView) {
    const template = getProviderTemplate(provider.templateId ?? "");
    setEditId(provider.id);
    setEdit({
      name: provider.name,
      baseUrl: provider.baseUrl,
      gatewayBaseUrl: provider.gatewayBaseUrl,
      gatewayUpstreamId: provider.gatewayUpstreamId,
      connectionMode: provider.connectionMode,
      model: provider.model,
      format: provider.apiFormat,
      contextWindow: provider.contextWindowSize,
      reasoning: provider.reasoningMode,
      key: "",
      presetModels: template?.presetModels ?? [],
      reasoningOptions: template?.reasoningOptions ?? [
        { value: "off" as const, label: "关闭" },
        { value: "high" as const, label: "开启" },
      ],
    });
    setListedModels([]);
    setActionState({ status: "idle" });
  }

  function saveEdit() {
    if (!editId || !edit) return;
    store.update(editId, {
      name: edit.name,
      baseUrl: edit.baseUrl,
      gatewayBaseUrl: edit.gatewayBaseUrl,
      gatewayUpstreamId: edit.gatewayUpstreamId,
      connectionMode: edit.connectionMode,
      model: edit.model,
      apiFormat: edit.format,
      contextWindowSize: Number(edit.contextWindow) || 128_000,
      reasoningMode: edit.reasoning,
      ...(edit.connectionMode === "direct" && edit.key ? { apiKey: edit.key } : {}),
    });
    setEditId(null);
    setEdit(null);
    refresh();
  }

  async function runTest(providerId: string) {
    const secret = store.getWithSecret(providerId);
    if (!secret) return;
    if (secret.connectionMode === "direct" && !secret.apiKey) {
      setActionState({ status: "error", message: "API key 未设置，无法测试连接" });
      return;
    }
    setActionState({ status: "loading", action: "test" });
    const result = await client.testConnection({
      baseUrl: secret.baseUrl,
      gatewayBaseUrl: secret.gatewayBaseUrl,
      gatewayUpstreamId: secret.gatewayUpstreamId,
      connectionMode: secret.connectionMode,
      apiKey: secret.apiKey,
      apiFormat: secret.apiFormat,
      model: secret.model,
    });
    if (!result.ok) {
      setActionState({
        status: "error",
        message: `[${result.kind}] ${result.error}${result.status ? ` (HTTP ${result.status})` : ""}`,
      });
      return;
    }
    setActionState({
      status: "success",
      message: `连接成功 · ${result.data.latencyMs}ms · ${result.data.url}`,
    });
  }

  async function runListModels(providerId: string) {
    const secret = store.getWithSecret(providerId);
    if (!secret) return;
    if (secret.connectionMode === "direct" && !secret.apiKey) {
      setActionState({ status: "error", message: "API key 未设置，无法拉取模型" });
      return;
    }
    setActionState({ status: "loading", action: "models" });
    const result = await client.listModels({
      baseUrl: secret.baseUrl,
      gatewayBaseUrl: secret.gatewayBaseUrl,
      gatewayUpstreamId: secret.gatewayUpstreamId,
      connectionMode: secret.connectionMode,
      apiKey: secret.apiKey,
      apiFormat: secret.apiFormat,
    });
    if (!result.ok) {
      setListedModels([]);
      setActionState({
        status: "error",
        message: `[${result.kind}] ${result.error}${result.status ? ` (HTTP ${result.status})` : ""}`,
      });
      return;
    }
    setListedModels(result.data.models);
    setActionState({
      status: "success",
      message: `已拉取 ${result.data.models.length} 个模型 · ${result.data.url}`,
    });
  }

  return (
    <section className="card">
      <h2>模型供应商</h2>
      <p className="muted">
        直连模式的 API key 仅保存在内存，禁止写入 localStorage，且请求可能被 CORS
        拦截；同源 Gateway 模式由服务端注入供应商密钥，浏览器不会保存或发送真实 API key。
      </p>

      <ProviderCreateSection
        templateId={templateId}
        apiKey={apiKey}
        connectionMode={connectionMode}
        gatewayBaseUrl={gatewayBaseUrl}
        gatewayUpstreamId={gatewayUpstreamId}
        onTemplateIdChange={setTemplateId}
        onApiKeyChange={setApiKey}
        onConnectionModeChange={setConnectionMode}
        onGatewayBaseUrlChange={setGatewayBaseUrl}
        onGatewayUpstreamIdChange={setGatewayUpstreamId}
        onAdd={addProvider}
      />

      {actionState.status === "loading" && (
        <p className="muted">请求中：{actionState.action}…</p>
      )}
      {actionState.status === "success" && <p className="badge">{actionState.message}</p>}
      {actionState.status === "error" && (
        <p className="muted" role="alert">
          错误：{actionState.message}
        </p>
      )}

      <div className="list">
        {providers.map((provider) => (
          <ProviderCard
            key={provider.id}
            provider={provider}
            isActive={provider.id === activeId}
            isEditing={editId === provider.id}
            edit={editId === provider.id ? edit : null}
            listedModels={listedModels}
            actions={{
              onSetActive: () => {
                store.setActive(provider.id);
                refresh();
              },
              onStartEdit: () => startEdit(provider),
              onTest: () => void runTest(provider.id),
              onListModels: () => void runListModels(provider.id),
              onRemove: () => {
                store.remove(provider.id);
                refresh();
              },
              onSaveEdit: saveEdit,
            }}
            onEditChange={(patch) => setEdit((prev) => (prev ? { ...prev, ...patch } : prev))}
          />
        ))}
        {providers.length === 0 && <p className="muted">尚未配置供应商。</p>}
      </div>
    </section>
  );
}
