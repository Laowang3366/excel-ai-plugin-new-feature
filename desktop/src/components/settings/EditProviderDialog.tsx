/**
 * EditProviderDialog — 编辑供应商弹窗子组件
 *
 * 负责编辑供应商本地草稿、连接测试和保存 patch 编排。
 */

import React, { useState, useCallback } from "react";
import { useSettingsStore, PROVIDER_TEMPLATES } from "../../store/settingsStore";
import type { AiProviderConfig, ModelConfig, ReasoningMode } from "../../electronApi";
import { MODEL_TEXT } from "./modelSettingsI18n";
import { useTestConnection } from "./useTestConnection";
import { ipcApi } from "../../services/ipcApi";
import {
  coerceReasoningMode,
  defaultReasoningModeForOptions,
  resolveReasoningOptionValues,
} from "../../utils/reasoningSupport";
import { buildEditProviderPatch } from "./editProviderPatch";
import { buildReasoningAutoHint } from "./providerReasoningHint";
import { ProviderDialogFrame } from "./ProviderDialogFrame";
import { EditProviderDialogActions } from "./ProviderDialogActions";
import { EditProviderDialogFields } from "./EditProviderDialogFields";

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
  const [contextWindowSize, setContextWindowSize] = useState<number | undefined>(
    provider.contextWindowSize,
  );
  const [reasoningMode, setReasoningMode] = useState<ReasoningMode>(
    provider.reasoningMode || "off",
  );
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
      const models = await ipcApi.ai.listModels(baseUrl, apiKey, apiFormat, provider.id);
      if (models.length > 0) {
        // 更新本地草稿中的模型列表
      }
    } catch {
      // 静默失败
    } finally {
      setFetchingModels(false);
    }
  }, [baseUrl, apiKey, apiFormat, provider.id, template]);

  const handleTestConnection = useCallback(() => {
    if (!apiKey || !baseUrl) return;
    testConnection(baseUrl, apiKey, apiFormat, model, provider.id);
  }, [baseUrl, apiKey, apiFormat, model, provider.id, testConnection]);

  /**
   * 切换模型时自动应用 per-model 配置
   * 同步 contextWindowSize、reasoningMode、reasoningOptions
   */
  const applyModelConfig = useCallback(
    (newModel: string) => {
      const mc = modelConfigs.find((m) => m.name === newModel);
      if (mc?.contextWindowSize) setContextWindowSize(mc.contextWindowSize);
      const optionValues = resolveReasoningOptionValues(
        { ...provider, apiFormat, model: newModel },
        template,
        mc,
      );
      const fallbackMode = defaultReasoningModeForOptions(
        optionValues,
        template?.defaultReasoningMode,
      );
      setReasoningMode(
        coerceReasoningMode(mc?.reasoningMode || reasoningMode, optionValues, fallbackMode),
      );
    },
    [apiFormat, modelConfigs, provider, reasoningMode, template],
  );

  const handleModelChange = useCallback(
    (newModel: string) => {
      setModel(newModel);
      applyModelConfig(newModel);
    },
    [applyModelConfig],
  );

  const handleSave = () => {
    onSave(
      buildEditProviderPatch(provider, {
        name,
        apiFormat,
        baseUrl,
        apiKey,
        model,
        contextWindowSize,
        effectiveReasoningMode,
        modelConfigs,
      }),
    );
  };

  return (
    <ProviderDialogFrame
      dialogClassName="edit-provider-dialog"
      title={text.editDialogTitle(name || provider.name)}
      onClose={onClose}
      actions={
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
      }
    >
      <EditProviderDialogFields
        state={{
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
          templateDefaultContextWindowSize: template?.defaultContextWindowSize,
          testing,
          testResult,
          fetchingModels,
          effectiveReasoningMode,
          defaultReasoningMode,
          reasoningOptionValues,
          reasoningAutoHint,
          language,
        }}
        actions={{
          setName,
          setApiFormat,
          setBaseUrl,
          setApiKey,
          setShowApiKey,
          setContextWindowSize,
          setModelConfigs,
          onModelChange: handleModelChange,
          onFetchModels: fetchModels,
          onTestConnection: handleTestConnection,
          onReasoningModeChange: setReasoningMode,
        }}
      />
    </ProviderDialogFrame>
  );
};
