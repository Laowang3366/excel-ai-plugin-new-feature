import { PROVIDER_TEMPLATES } from "@shared/provider";

interface Props {
  templateId: string;
  apiKey: string;
  onTemplateIdChange: (value: string) => void;
  onApiKeyChange: (value: string) => void;
  onAdd: () => void;
}

export function ProviderCreateSection({
  templateId,
  apiKey,
  onTemplateIdChange,
  onApiKeyChange,
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
          API Key（可选）
          <input
            type="password"
            value={apiKey}
            onChange={(e) => onApiKeyChange(e.target.value)}
            autoComplete="off"
          />
        </label>
      </div>
      <div className="row">
        <button type="button" onClick={onAdd}>
          添加
        </button>
      </div>
    </>
  );
}
