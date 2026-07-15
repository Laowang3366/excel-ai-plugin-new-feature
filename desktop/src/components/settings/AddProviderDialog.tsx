/**
 * AddProviderDialog — 添加供应商弹窗子组件
 *
 * 负责供应商模板选择、本地草稿、连接测试和新增提交编排。
 */

import React, { useState } from "react";
import { useSettingsStore, PROVIDER_TEMPLATES, API_FORMATS } from "../../store/settingsStore";
import type { AiProviderConfig, ModelConfig, ReasoningMode } from "../../electronApi";
import { ipcApi } from "../../services/ipcApi";
import {
  MODEL_TEXT,
  DIRECT_TEMPLATES,
  AGGREGATION_TEMPLATES,
  OTHER_TEMPLATES,
} from "./modelSettingsI18n";
import { AddProviderTemplateSelect } from "./AddProviderTemplateSelect";
import { useTestConnection } from "./useTestConnection";
import { ReasoningModeSelect } from "./ReasoningModeSelect";
import { ModelConfigList } from "./ModelConfigList";
import {
  buildReasoningOptions,
  coerceReasoningMode,
  defaultReasoningModeForOptions,
  resolveReasoningOptionValues,
} from "../../utils/reasoningSupport";
import { buildReasoningAutoHint } from "./providerReasoningHint";
import {
  buildProviderConfigFromDraft,
  createEmptyProviderDraft,
  providerDraftFromTemplate,
  type AddProviderDraft,
} from "./addProviderDraft";
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
import { AddProviderDialogActions } from "./ProviderDialogActions";

export interface AddProviderDialogProps {
  onAdd: (config: AiProviderConfig) => void;
  onClose: () => void;
  generateId: () => string;
}

export const AddProviderDialog: React.FC<AddProviderDialogProps> = ({
  onAdd,
  onClose,
  generateId,
}) => {
  const { language } = useSettingsStore();
  const text = MODEL_TEXT[language];
  const [selectedTemplateId, setSelectedTemplateId] = useState<string>("");
  const [name, setName] = useState("");
  const [apiFormat, setApiFormat] = useState("openai");
  const [baseUrl, setBaseUrl] = useState("");
  const [apiKey, setApiKey] = useState("");
  const [model, setModel] = useState("");
  const [contextWindowSize, setContextWindowSize] = useState<number | undefined>(undefined);
  const [reasoningMode, setReasoningMode] = useState<ReasoningMode>("off");
  const [modelConfigs, setModelConfigs] = useState<ModelConfig[]>([]);
  const [showApiKey, setShowApiKey] = useState(false);

  const { testing, testResult, testConnection } = useTestConnection({
    testFailedText: text.testFailed,
  });

  const selectedTemplate = PROVIDER_TEMPLATES.find((t) => t.id === selectedTemplateId) || null;
  const isAggregation = selectedTemplate?.category === "aggregation";
  const reasoningOptionValues = resolveReasoningOptionValues(
    {
      provider: selectedTemplate?.provider || "custom",
      apiFormat,
      model,
    },
    selectedTemplate,
  );
  const defaultReasoningMode = defaultReasoningModeForOptions(
    reasoningOptionValues,
    selectedTemplate?.defaultReasoningMode,
  );
  const effectiveReasoningMode = coerceReasoningMode(
    reasoningMode,
    reasoningOptionValues,
    defaultReasoningMode,
  );
  const reasoningAutoHint = buildReasoningAutoHint(reasoningOptionValues, language);

  const applyDraft = (draft: AddProviderDraft) => {
    setSelectedTemplateId(draft.selectedTemplateId);
    setName(draft.name);
    setApiFormat(draft.apiFormat);
    setBaseUrl(draft.baseUrl);
    setModel(draft.model);
    setContextWindowSize(draft.contextWindowSize);
    setReasoningMode(draft.reasoningMode);
    setModelConfigs(draft.modelConfigs);
  };

  const handleSelectTemplate = (templateId: string) => {
    const template = PROVIDER_TEMPLATES.find((t) => t.id === templateId);
    if (!template) {
      applyDraft(createEmptyProviderDraft());
      return;
    }
    applyDraft(providerDraftFromTemplate(template));
  };

  const handleTestModel = () => {
    if (!apiKey || !baseUrl || !model) return;
    testConnection(baseUrl, apiKey, apiFormat, model);
  };

  const handleAddWithTest = async () => {
    if (!apiKey || !baseUrl) return;
    const resolvedModel = model || selectedTemplate?.defaultModel || "";
    // 注意：testConnection 是异步的，需要等待结果
    // 但 useTestConnection 的 testConnection 只更新内部状态，不返回结果
    // 所以这里保持原有的内联调用方式
    try {
      const result = await ipcApi.ai.testConnection(baseUrl, apiKey, apiFormat, resolvedModel);
      if (result.success) {
        handleAdd(resolvedModel);
      }
    } catch {
      // 静默失败，用户可通过独立测试按钮查看错误
    }
  };

  const handleAdd = (resolvedModel?: string) => {
    const config: AiProviderConfig = buildProviderConfigFromDraft({
      id: generateId(),
      draft: {
        selectedTemplateId,
        name,
        apiFormat,
        baseUrl,
        apiKey,
        model,
        contextWindowSize,
        reasoningMode,
        modelConfigs,
      },
      selectedTemplate,
      effectiveReasoningMode,
      customProviderName: text.customProvider,
      resolvedModel,
    });
    onAdd(config);
  };

  const canAdd = apiKey.trim().length > 0 && baseUrl.trim().length > 0;
  const canTestModel = canAdd && model.trim().length > 0;

  return (
    <ProviderDialogFrame
      dialogClassName="add-provider-dialog"
      title={text.addDialogTitle}
      onClose={onClose}
      actions={
        <AddProviderDialogActions
          canAdd={canAdd}
          testing={testing}
          cancelLabel={text.cancel}
          addLabel={text.add}
          addAndTestLabel={text.addAndTest}
          testingLabel={text.testing}
          onCancel={onClose}
          onAdd={() => handleAdd()}
          onAddWithTest={handleAddWithTest}
        />
      }
    >
      <AddProviderTemplateSelect
        label={text.providerType}
        value={selectedTemplateId}
        customProviderLabel={text.customProvider}
        directProvidersLabel={text.directProviders}
        aggregationProvidersLabel={text.aggregationProviders}
        otherLabel={text.other}
        directTemplates={DIRECT_TEMPLATES}
        aggregationTemplates={AGGREGATION_TEMPLATES}
        otherTemplates={OTHER_TEMPLATES}
        onChange={handleSelectTemplate}
      />

      {/* 配置表单 */}
      <div className="add-provider-form">
        <ProviderNameField
          label={text.providerName}
          value={name}
          placeholder={text.providerNamePlaceholder}
          onChange={setName}
        />

        <ProviderApiFormatField
          label={text.apiFormat}
          value={apiFormat}
          options={API_FORMATS}
          onChange={setApiFormat}
        />

        <ProviderBaseUrlField
          label={text.apiUrl}
          value={baseUrl}
          placeholder="https://api.example.com/v1"
          onChange={setBaseUrl}
        />

        <ProviderApiKeyField
          label={text.apiKey}
          value={apiKey}
          showApiKey={showApiKey}
          onChange={setApiKey}
          onToggleVisibility={() => setShowApiKey(!showApiKey)}
        />

        <ProviderModelField
          label={text.model}
          selector={
            <ProviderModelSelector
              value={model}
              onChange={setModel}
              isAggregation={isAggregation}
              modelConfigs={modelConfigs}
              modelOptions={selectedTemplate?.presetModels || []}
              noModelLabel={text.noModel}
              showEmptyOption
            />
          }
        />

        <ProviderContextWindowField
          label={text.contextWindowSize}
          value={contextWindowSize}
          placeholder={text.contextWindowPlaceholder}
          hint={
            selectedTemplate?.defaultContextWindowSize
              ? text.contextWindowHint(selectedTemplate.defaultContextWindowSize)
              : undefined
          }
          action={
            <ProviderTestButton
              className="btn-test"
              testing={testing}
              label={text.testModel}
              disabled={!canTestModel || testing}
              title={text.testModel}
              onClick={handleTestModel}
            />
          }
          onChange={setContextWindowSize}
        />

        {/* 思考等级 */}
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
              onAddModelAutoSelect={(name) => {
                if (!model) setModel(name);
              }}
            />
          </div>
        )}
      </div>

      <ProviderTestResult
        result={testResult}
        successText={text.connectionOk}
        errorFallback={text.connectionFailed}
      />
    </ProviderDialogFrame>
  );
};
