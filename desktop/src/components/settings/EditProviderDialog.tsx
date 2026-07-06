/**
 * EditProviderDialog — 编辑供应商弹窗子组件
 *
 * 负责编辑供应商本地草稿、连接测试和保存 patch 编排。
 */

import React, { useState, useCallback } from "react";
import { useSettingsStore, PROVIDER_TEMPLATES, API_FORMATS } from "../../store/settingsStore";
import type { AiProviderConfig, ModelConfig, ReasoningMode } from "../../electronApi";
import {
  Loader2,
  RotateCcw,
  RefreshCw,
} from "../common/IconMap";
import { MODEL_TEXT } from "./modelSettingsI18n";
import { useTestConnection } from "./useTestConnection";
import { ReasoningModeSelect } from "./ReasoningModeSelect";
import { ModelConfigList } from "./ModelConfigList";
import { formatTokensAsK } from "../../utils/modelContextWindows";
import { ipcApi } from "../../services/ipcApi";
import {
  buildReasoningOptions,
  coerceReasoningMode,
  defaultReasoningModeForOptions,
  resolveReasoningOptionValues,
} from "../../utils/reasoningSupport";
import { buildEditProviderPatch } from "./editProviderPatch";
import { buildReasoningAutoHint } from "./providerReasoningHint";
import { ProviderModelSelector } from "./ProviderModelSelector";
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
import { ProviderDialogFrame } from "./ProviderDialogFrame";
import { EditProviderDialogActions } from "./ProviderDialogActions";

export interface EditProviderDialogProps {
  provider: AiProviderConfig;
  onSave: (patch: Partial<AiProviderConfig>) => void;
  onClose: () => void;
}

export const EditProviderDialog: React.FC<EditProviderDialogProps> = ({
  provider,
  onSave,
  onClose,
}) => {
  const { language } = useSettingsStore();
  const text = MODEL_TEXT[language];

  const [name, setName] = useState(provider.name);
  const [apiFormat, setApiFormat] = useState(provider.apiFormat || "openai");
  const [baseUrl, setBaseUrl] = useState(provider.baseUrl);
  const [apiKey, setApiKey] = useState(provider.apiKey);
  const [model, setModel] = useState(provider.model);
  const [contextWindowSize, setContextWindowSize] = useState<number | undefined>(provider.contextWindowSize);
  const [reasoningMode, setReasoningMode] = useState<ReasoningMode>(provider.reasoningMode || "off");
  const [modelConfigs, setModelConfigs] = useState<ModelConfig[]>(provider.modelConfigs || []);
  const [showApiKey, setShowApiKey] = useState(false);
  const [fetchingModels, setFetchingModels] = useState(false);

  const { testing, testResult, testConnection } = useTestConnection({
    testFailedText: text.testFailed,
  });

  const template = PROVIDER_TEMPLATES.find((t) => t.provider === provider.provider);
  const isAggregation = template?.category === "aggregation";

  const activeModelConfig = modelConfigs.find((m) => m.name === model);
  const reasoningOptionValues = resolveReasoningOptionValues(
    { ...provider, apiFormat, model },
    template,
    activeModelConfig,
  );
  const defaultReasoningMode = defaultReasoningModeForOptions(
    reasoningOptionValues,
    template?.defaultReasoningMode,
  );
  const effectiveReasoningMode = coerceReasoningMode(
    activeModelConfig?.reasoningMode || reasoningMode,
    reasoningOptionValues,
    defaultReasoningMode,
  );
  const reasoningAutoHint = buildReasoningAutoHint(reasoningOptionValues, language);

  const availableModels = provider.models || [];
  const presetModels = template?.presetModels || [];
  const modelOptions = availableModels.length > 0 ? availableModels : presetModels;

  const fetchModels = useCallback(async () => {
    if (!apiKey || !baseUrl) return;
    setFetchingModels(true);
    try {
      if (template?.presetModels && template.presetModels.length > 0) {
        setFetchingModels(false);
        return;
      }
      const models = await ipcApi.ai.listModels(baseUrl, apiKey, apiFormat);
      if (models.length > 0) {
        // 更新本地草稿中的模型列表
      }
    } catch {
      // 静默失败
    } finally {
      setFetchingModels(false);
    }
  }, [baseUrl, apiKey, apiFormat, template]);

  const handleTestConnection = useCallback(() => {
    if (!apiKey || !baseUrl) return;
    testConnection(baseUrl, apiKey, apiFormat, model);
  }, [baseUrl, apiKey, apiFormat, model, testConnection]);

  /**
   * 切换模型时自动应用 per-model 配置
   * 同步 contextWindowSize、reasoningMode、reasoningOptions
   */
  const applyModelConfig = useCallback((newModel: string) => {
    const mc = modelConfigs.find((m) => m.name === newModel);
    if (mc?.contextWindowSize) setContextWindowSize(mc.contextWindowSize);
    const optionValues = resolveReasoningOptionValues(
      { ...provider, apiFormat, model: newModel },
      template,
      mc,
    );
    const fallbackMode = defaultReasoningModeForOptions(optionValues, template?.defaultReasoningMode);
    setReasoningMode(coerceReasoningMode(mc?.reasoningMode || reasoningMode, optionValues, fallbackMode));
  }, [apiFormat, modelConfigs, provider, reasoningMode, template]);

  const handleModelChange = useCallback((newModel: string) => {
    setModel(newModel);
    applyModelConfig(newModel);
  }, [applyModelConfig]);

  const handleSave = () => {
    onSave(buildEditProviderPatch(provider, {
      name,
      apiFormat,
      baseUrl,
      apiKey,
      model,
      contextWindowSize,
      effectiveReasoningMode,
      modelConfigs,
    }));
  };

  return (
    <ProviderDialogFrame
      dialogClassName="edit-provider-dialog"
      title={text.editDialogTitle(name || provider.name)}
      onClose={onClose}
      actions={(
        <EditProviderDialogActions
          testing={testing}
          canTest={Boolean(apiKey && baseUrl)}
          testConnectionLabel={text.testConnection}
          cancelLabel={text.cancel}
          saveLabel={text.save}
          onTest={handleTestConnection}
          onCancel={onClose}
          onSave={handleSave}
        />
      )}
    >
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
            hint={provider.defaultBaseUrl ? `${text.defaultValue}: ${provider.defaultBaseUrl}` : undefined}
            action={provider.defaultBaseUrl ? (
              <button
                className="btn-icon-hint"
                onClick={() => setBaseUrl(provider.defaultBaseUrl!)}
                title={`恢复默认: ${provider.defaultBaseUrl}`}
              >
                <RotateCcw size={13} />
              </button>
            ) : undefined}
            onChange={setBaseUrl}
          />

          <ProviderApiKeyField
            label={text.apiKey}
            value={apiKey}
            showApiKey={showApiKey}
            action={(
              <ProviderTestButton
                className="btn-test"
                testing={testing}
                label={text.test}
                disabled={testing || !apiKey || !baseUrl}
                title={text.testConnection}
                onClick={handleTestConnection}
              />
            )}
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
            selector={(
              <ProviderModelSelector
                value={model}
                onChange={handleModelChange}
                isAggregation={isAggregation}
                modelConfigs={modelConfigs}
                modelOptions={modelOptions}
                noModelLabel={text.noModel}
                placeholder={provider.defaultModel || "model-name"}
                showEmptyOption={isAggregation && !model}
                preserveCurrentValue
              />
            )}
            action={!isAggregation ? (
              <button
                className="btn-refresh"
                onClick={fetchModels}
                disabled={fetchingModels || !apiKey || !baseUrl}
                title={text.refreshModels}
              >
                {fetchingModels ? <Loader2 size={14} className="spin" /> : <RefreshCw size={14} />}
              </button>
            ) : undefined}
            hints={[
              ...(provider.defaultModel ? [`${text.defaultValue}: ${provider.defaultModel}`] : []),
              ...(modelOptions.length > 0 ? [text.availableModels(modelOptions.length)] : []),
            ]}
          />

          <ProviderContextWindowField
            label={text.contextWindowSize}
            value={contextWindowSize}
            placeholder={text.contextWindowPlaceholder}
            hint={template?.defaultContextWindowSize ? text.contextWindowHint(template.defaultContextWindowSize) : undefined}
            action={template?.defaultContextWindowSize ? (
              <button
                className="btn-icon-hint"
                onClick={() => setContextWindowSize(template.defaultContextWindowSize)}
                title={`恢复默认: ${formatTokensAsK(template.defaultContextWindowSize)}`}
              >
                <RotateCcw size={13} />
              </button>
            ) : undefined}
            onChange={setContextWindowSize}
          />

          {/* 思考等级 — 自动根据供应商/API/模型适配 */}
          {reasoningOptionValues.length > 0 && (
            <ReasoningModeSelect
              reasoningOptions={buildReasoningOptions(reasoningOptionValues, language)}
              value={effectiveReasoningMode}
              defaultMode={defaultReasoningMode}
              onChange={(mode) => setReasoningMode(mode)}
              hint={reasoningAutoHint}
            />
          )}

          {/* 聚合平台：模型列表管理 */}
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
    </ProviderDialogFrame>
  );
};
