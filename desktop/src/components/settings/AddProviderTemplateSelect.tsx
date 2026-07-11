interface ProviderTemplateOption {
  id: string;
  name: string;
}

interface AddProviderTemplateSelectProps {
  label: string;
  value: string;
  customProviderLabel: string;
  directProvidersLabel: string;
  aggregationProvidersLabel: string;
  otherLabel: string;
  directTemplates: ProviderTemplateOption[];
  aggregationTemplates: ProviderTemplateOption[];
  otherTemplates: ProviderTemplateOption[];
  onChange: (templateId: string) => void;
}

export function AddProviderTemplateSelect({
  label,
  value,
  customProviderLabel,
  directProvidersLabel,
  aggregationProvidersLabel,
  otherLabel,
  directTemplates,
  aggregationTemplates,
  otherTemplates,
  onChange,
}: AddProviderTemplateSelectProps) {
  return (
    <div className="form-group">
      <label className="form-label">{label}</label>
      <select
        className="form-input template-select"
        value={value}
        onChange={(event) => onChange(event.target.value)}
      >
        <option value="">{customProviderLabel}</option>
        <optgroup label={directProvidersLabel}>
          {directTemplates.map((template) => (
            <option key={template.id} value={template.id}>{template.name}</option>
          ))}
        </optgroup>
        <optgroup label={aggregationProvidersLabel}>
          {aggregationTemplates.map((template) => (
            <option key={template.id} value={template.id}>{template.name}</option>
          ))}
        </optgroup>
        <optgroup label={otherLabel}>
          {otherTemplates.map((template) => (
            <option key={template.id} value={template.id}>{template.name}</option>
          ))}
        </optgroup>
      </select>
    </div>
  );
}
