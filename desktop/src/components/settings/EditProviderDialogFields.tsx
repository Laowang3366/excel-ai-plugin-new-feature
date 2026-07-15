import React from "react";
import type { AiProviderConfig, ModelConfig, ReasoningMode } from "../../electronApi";
import { Loader2, RotateCcw, RefreshCw } from "../common/IconMap";
import { API_FORMATS } from "../../store/settingsStore";
import { formatTokensAsK } from "../../utils/modelContextWindows";
import { buildReasoningOptions } from "../../utils/reasoningSupport";
import { ReasoningModeSelect } from "./ReasoningModeSelect";
import { ModelConfigList } from "./ModelConfigList";
import { ProviderModelSelector } from "./ProviderModelSelector";
import type { TestResult } from "./useTestConnection";
import {
  ProviderApiFormatField,
  ProviderApiKeyField,
  ProviderBaseUrlField,
  ProviderContextWindowField,
  ProviderModelField,
  ProviderNameField,
  ProviderTestButton,
  ProviderTestResult,
} from "./ProviderDialogFields";

export interface EditProviderDialogFieldsState {
  text: {
    apiKey: string;
    test: string;
    testConnection: string;
    connectionOk: (latency?: number) => string;
    connectionFailed: string;
    model: string;
    noModel: string;
    refreshModels: string;
    defaultValue: string;
    availableModels: (count: number) => string;
    contextWindowSize: string;
    contextWindowPlaceholder: string;
    contextWindowHint: (size: number) => string;
    modelList: string;
  };
  name: string;
  apiFormat: string;
  baseUrl: string;
  apiKey: string;
  showApiKey: boolean;
  model: string;
  contextWindowSize: number | undefined;
  modelConfigs: ModelConfig[];
  isAggregation: boolean;
  modelOptions: string[];
  provider: AiProviderConfig;
  templateDefaultContextWindowSize?: number;
  testing: boolean;
  testResult: TestResult | null;
  fetchingModels: boolean;
  effectiveReasoningMode: ReasoningMode;
  defaultReasoningMode: ReasoningMode;
  reasoningOptionValues: ReasoningMode[];
  reasoningAutoHint?: string;
  language: "zh-CN" | "en-US";
}

export interface EditProviderDialogFieldsActions {
  setName: (value: string) => void;
  setApiFormat: (value: string) => void;
  setBaseUrl: (value: string) => void;
  setApiKey: (value: string) => void;
  setShowApiKey: (value: boolean) => void;
  setContextWindowSize: (value: number | undefined) => void;
  setModelConfigs: (value: ModelConfig[]) => void;
  onModelChange: (model: string) => void;
  onFetchModels: () => void;
  onTestConnection: () => void;
  onReasoningModeChange: (mode: ReasoningMode) => void;
}

export interface EditProviderDialogFieldsProps {
  state: EditProviderDialogFieldsState;
  actions: EditProviderDialogFieldsActions;
}

export function EditProviderDialogFields({ state, actions }: EditProviderDialogFieldsProps) {
  const {
    text,
    name,
    apiFormat,
    baseUrl,
    apiKey,
    showApiKey,
    model,
    contextWindowSize,
    modelConfigs,
    isAggregation,
    modelOptions,
    provider,
    templateDefaultContextWindowSize,
    testing,
    testResult,
    fetchingModels,
    effectiveReasoningMode,
    defaultReasoningMode,
    reasoningOptionValues,
    reasoningAutoHint,
    language,
  } = state;
  const {
    setName,
    setApiFormat,
    setBaseUrl,
    setApiKey,
    setShowApiKey,
    setContextWindowSize,
    setModelConfigs,
    onModelChange,
    onFetchModels,
    onTestConnection,
    onReasoningModeChange,
  } = actions;

  return (
    <>
      <ProviderNameField
        label="供应商名称"
        value={name}
        placeholder="如：我的 DeepSeek"
        onChange={setName}
      />

      <ProviderApiFormatField
        label="API 格式"
        value={apiFormat}
        options={API_FORMATS}
        hint="选择 API 协议格式，影响请求方式和认证方式"
        onChange={setApiFormat}
      />

      <ProviderBaseUrlField
        label="API 地址"
        value={baseUrl}
        placeholder="https://api.example.com/v1"
        hint={
          provider.defaultBaseUrl ? `${text.defaultValue}: ${provider.defaultBaseUrl}` : undefined
        }
        action={
          provider.defaultBaseUrl ? (
            <button
              className="btn-icon-hint"
              onClick={() => setBaseUrl(provider.defaultBaseUrl!)}
              title={`恢复默认: ${provider.defaultBaseUrl}`}
            >
              <RotateCcw size={13} />
            </button>
          ) : undefined
        }
        onChange={setBaseUrl}
      />

      <ProviderApiKeyField
        label={text.apiKey}
        value={apiKey}
        showApiKey={showApiKey}
        action={
          <ProviderTestButton
            className="btn-test"
            testing={testing}
            label={text.test}
            disabled={testing || !apiKey || !baseUrl}
            title={text.testConnection}
            onClick={onTestConnection}
          />
        }
        onChange={setApiKey}
        onToggleVisibility={() => setShowApiKey(!showApiKey)}
      />

      <ProviderTestResult
        result={testResult}
        successText={text.connectionOk}
        errorFallback={text.connectionFailed}
      />

      <ProviderModelField
        label={text.model}
        selector={
          <ProviderModelSelector
            value={model}
            onChange={onModelChange}
            isAggregation={isAggregation}
            modelConfigs={modelConfigs}
            modelOptions={modelOptions}
            noModelLabel={text.noModel}
            placeholder={provider.defaultModel || "model-name"}
            showEmptyOption={isAggregation && !model}
            preserveCurrentValue
          />
        }
        action={
          !isAggregation ? (
            <button
              className="btn-refresh"
              onClick={onFetchModels}
              disabled={fetchingModels || !apiKey || !baseUrl}
              title={text.refreshModels}
            >
              {fetchingModels ? <Loader2 size={14} className="spin" /> : <RefreshCw size={14} />}
            </button>
          ) : undefined
        }
        hints={[
          ...(provider.defaultModel ? [`${text.defaultValue}: ${provider.defaultModel}`] : []),
          ...(modelOptions.length > 0 ? [text.availableModels(modelOptions.length)] : []),
        ]}
      />

      <ProviderContextWindowField
        label={text.contextWindowSize}
        value={contextWindowSize}
        placeholder={text.contextWindowPlaceholder}
        hint={
          templateDefaultContextWindowSize
            ? text.contextWindowHint(templateDefaultContextWindowSize)
            : undefined
        }
        action={
          templateDefaultContextWindowSize ? (
            <button
              className="btn-icon-hint"
              onClick={() => setContextWindowSize(templateDefaultContextWindowSize)}
              title={`恢复默认: ${formatTokensAsK(templateDefaultContextWindowSize)}`}
            >
              <RotateCcw size={13} />
            </button>
          ) : undefined
        }
        onChange={setContextWindowSize}
      />

      {reasoningOptionValues.length > 0 && (
        <ReasoningModeSelect
          reasoningOptions={buildReasoningOptions(reasoningOptionValues, language)}
          value={effectiveReasoningMode}
          defaultMode={defaultReasoningMode}
          onChange={onReasoningModeChange}
          hint={reasoningAutoHint}
        />
      )}

      {isAggregation && (
        <div className="form-group">
          <label>{text.modelList}</label>
          <ModelConfigList
            modelConfigs={modelConfigs}
            currentModel={model}
            onModelConfigsChange={(newConfigs) => setModelConfigs(newConfigs)}
          />
        </div>
      )}
    </>
  );
}
