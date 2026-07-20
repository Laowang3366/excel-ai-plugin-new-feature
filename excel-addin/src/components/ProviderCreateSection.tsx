import { PROVIDER_TEMPLATES, type ConnectionMode } from "@shared/provider";

interface Props {
  templateId: string;
  apiKey: string;
  connectionMode: ConnectionMode;
  gatewayBaseUrl: string;
  gatewayUpstreamId: string;
  onTemplateIdChange: (value: string) => void;
  onApiKeyChange: (value: string) => void;
  onConnectionModeChange: (value: ConnectionMode) => void;
  onGatewayBaseUrlChange: (value: string) => void;
  onGatewayUpstreamIdChange: (value: string) => void;
  onAdd: () => void;
}

export function ProviderCreateSection({
  templateId,
  apiKey,
  connectionMode,
  gatewayBaseUrl,
  gatewayUpstreamId,
  onTemplateIdChange,
  onApiKeyChange,
  onConnectionModeChange,
  onGatewayBaseUrlChange,
  onGatewayUpstreamIdChange,
  onAdd,
}: Props) {
  return (
    <>
      <div className="row">
        <label>
          模板
          <select value={templateId} onChange={(e) => onTemplateIdChange(e.target.value)}>
            {PROVIDER_TEMPLATES.map((template) => (
              <option key={template.id} value={template.id}>
                {template.name} · {template.apiFormat}
              </option>
            ))}
          </select>
        </label>
        <label>
          连接方式
          <select
            value={connectionMode}
            onChange={(e) => onConnectionModeChange(e.target.value as ConnectionMode)}
          >
            <option value="direct">浏览器直连</option>
            <option value="gateway">同源 Gateway</option>
          </select>
        </label>
      </div>

      {connectionMode === "direct" ? (
        <div className="row">
          <label>
            API Key（可选）
            <input
              type="password"
              value={apiKey}
              onChange={(e) => onApiKeyChange(e.target.value)}
              autoComplete="off"
            />
          </label>
        </div>
      ) : (
        <>
          <p className="muted">
            Gateway 模式不在浏览器保存或发送供应商 API Key，密钥由服务端按 upstream ID 注入。
          </p>
          <div className="row">
            <label>
              Gateway 地址（留空表示同源）
              <input
                value={gatewayBaseUrl}
                onChange={(e) => onGatewayBaseUrlChange(e.target.value)}
                placeholder="https://plugin.example.com"
                autoComplete="off"
              />
            </label>
            <label>
              Upstream ID
              <input
                value={gatewayUpstreamId}
                onChange={(e) => onGatewayUpstreamIdChange(e.target.value)}
                placeholder="openai"
                autoComplete="off"
              />
            </label>
          </div>
        </>
      )}

      <div className="row">
        <button type="button" onClick={onAdd}>
          添加
        </button>
      </div>
    </>
  );
}
