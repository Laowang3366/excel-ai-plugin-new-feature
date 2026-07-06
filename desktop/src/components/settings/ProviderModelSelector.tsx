import React from "react";
import type { ModelConfig } from "../../electronApi";
import { ChevronDown } from "../common/IconMap";

export type ProviderModelSelectorKind = "aggregation" | "select" | "input";

export interface ProviderModelSelectorProps {
  value: string;
  onChange: (model: string) => void;
  isAggregation: boolean;
  modelConfigs?: ModelConfig[];
  modelOptions?: string[];
  noModelLabel: string;
  placeholder?: string;
  showEmptyOption?: boolean;
  preserveCurrentValue?: boolean;
}

export function getProviderModelSelectorKind({
  isAggregation,
  modelOptions = [],
}: Pick<ProviderModelSelectorProps, "isAggregation" | "modelOptions">): ProviderModelSelectorKind {
  if (isAggregation) return "aggregation";
  return modelOptions.length > 0 ? "select" : "input";
}

function renderNoModelOption(label: string): React.ReactElement {
  return <option value="">-- {label} --</option>;
}

export const ProviderModelSelector: React.FC<ProviderModelSelectorProps> = ({
  value,
  onChange,
  isAggregation,
  modelConfigs = [],
  modelOptions = [],
  noModelLabel,
  placeholder = "model-name",
  showEmptyOption = false,
  preserveCurrentValue = false,
}) => {
  const kind = getProviderModelSelectorKind({ isAggregation, modelOptions });
  const handleChange = (event: React.ChangeEvent<HTMLInputElement | HTMLSelectElement>) => {
    onChange(event.target.value);
  };

  if (kind === "input") {
    return (
      <input
        type="text"
        className="form-input"
        value={value}
        onChange={handleChange}
        placeholder={placeholder}
      />
    );
  }

  const optionValues = kind === "aggregation"
    ? modelConfigs.map((modelConfig) => modelConfig.name)
    : modelOptions;
  const shouldPreserveCurrentValue =
    preserveCurrentValue && value.length > 0 && !optionValues.includes(value);

  return (
    <div className="model-select-wrapper">
      <select
        className="form-input model-select"
        value={value}
        onChange={handleChange}
      >
        {showEmptyOption && renderNoModelOption(noModelLabel)}
        {optionValues.map((modelName) => (
          <option key={modelName} value={modelName}>{modelName}</option>
        ))}
        {shouldPreserveCurrentValue && (
          <option value={value}>{value}</option>
        )}
      </select>
      <ChevronDown size={14} className="select-arrow" />
    </div>
  );
};
