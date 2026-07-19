import {
  API_FORMATS,
  type ApiFormat,
  type ProviderPublicView,
  type ReasoningMode,
} from "@shared/provider";

export interface ProviderCardEditState {
  name: string;
  baseUrl: string;
  model: string;
  format: ApiFormat;
  contextWindow: number;
  reasoning: ReasoningMode;
  key: string;
  presetModels: string[];
  reasoningOptions: Array<{ value: ReasoningMode; label: string }>;
}

export interface ProviderCardActions {
  onSetActive: () => void;
  onStartEdit: () => void;
  onTest: () => void;
  onListModels: () => void;
  onRemove: () => void;
  onSaveEdit: () => void;
}

interface Props {
  provider: ProviderPublicView;
  isActive: boolean;
  isEditing: boolean;
  edit: ProviderCardEditState | null;
  listedModels: string[];
  actions: ProviderCardActions;
  onEditChange: (patch: Partial<ProviderCardEditState>) => void;
}

export function ProviderCard({
  provider,
  isActive,
  isEditing,
  edit,
  listedModels,
  actions,
  onEditChange,
}: Props) {
  return (
    <div className={`list-item${isActive ? " active" : ""}`}>
      <strong>
        {provider.name} <span className="badge">{provider.apiFormat}</span>
      </strong>
      <span className="muted">
        {provider.provider} · {provider.model || "(无默认模型)"} · ctx {provider.contextWindowSize}{" "}
        · reason {provider.reasoningMode} · key {provider.hasApiKey ? "已设" : "未设"}
      </span>
      {!provider.hasApiKey && (
        <span className="muted">提示：未设置 API key，连接测试/拉模型将失败。</span>
      )}
      <div className="row">
        <button type="button" onClick={actions.onSetActive}>
          设为活动
        </button>
        <button type="button" onClick={actions.onStartEdit}>
          编辑
        </button>
        <button type="button" onClick={actions.onTest}>
          测试连接
        </button>
        <button type="button" onClick={actions.onListModels}>
          拉取模型
        </button>
        <button type="button" onClick={actions.onRemove}>
          删除
        </button>
      </div>
      {isEditing && edit && (
        <div className="row">
          <label>
            名称
            <input value={edit.name} onChange={(e) => onEditChange({ name: e.target.value })} />
          </label>
          <label>
            Base URL
            <input
              value={edit.baseUrl}
              onChange={(e) => onEditChange({ baseUrl: e.target.value })}
            />
          </label>
          <label>
            Model
            <input value={edit.model} onChange={(e) => onEditChange({ model: e.target.value })} />
          </label>
          {edit.presetModels.length > 0 && (
            <label>
              预设模型
              <select
                value={edit.presetModels.includes(edit.model) ? edit.model : ""}
                onChange={(e) => {
                  if (e.target.value) onEditChange({ model: e.target.value });
                }}
              >
                <option value="">自定义 / 手动输入</option>
                {edit.presetModels.map((model) => (
                  <option key={model} value={model}>
                    {model}
                  </option>
                ))}
              </select>
            </label>
          )}
          <label>
            API Format
            <select
              value={edit.format}
              onChange={(e) => onEditChange({ format: e.target.value as ApiFormat })}
            >
              {API_FORMATS.map((format) => (
                <option key={format.value} value={format.value}>
                  {format.label}
                </option>
              ))}
            </select>
          </label>
          <label>
            Context Window
            <input
              type="number"
              min={1024}
              value={edit.contextWindow}
              onChange={(e) => onEditChange({ contextWindow: Number(e.target.value) })}
            />
          </label>
          <label>
            Reasoning
            <select
              value={edit.reasoning}
              onChange={(e) => onEditChange({ reasoning: e.target.value as ReasoningMode })}
            >
              {edit.reasoningOptions.map((option) => (
                <option key={option.value} value={option.value}>
                  {option.label}
                </option>
              ))}
            </select>
          </label>
          <label>
            新 API Key
            <input
              type="password"
              value={edit.key}
              onChange={(e) => onEditChange({ key: e.target.value })}
              autoComplete="off"
            />
          </label>
          <button type="button" onClick={actions.onSaveEdit}>
            保存
          </button>
        </div>
      )}
      {listedModels.length > 0 && isEditing && <pre>{listedModels.slice(0, 30).join("\n")}</pre>}
    </div>
  );
}
