/**
 * AddProviderDialog — 添加供应商弹窗子组件
 *
 * 从 ModelSettings.tsx 提取，负责：
 * - 供应商模板选择（直连/聚合/自定义）
 * - 配置表单（名称、API 格式、地址、密钥、模型、上下文窗口、思考等级）
 * - 测试连接 / 测试并添加
 * - 聚合平台的模型列表管理
 *
 * 关联模块：
 * - modelSettingsI18n.ts — 提供 MODEL_TEXT、DIRECT_TEMPLATES、AGGREGATION_TEMPLATES、OTHER_TEMPLATES
 * - settingsStore — PROVIDER_TEMPLATES、API_FORMATS、useSettingsStore
 * - electronApi — AiProviderConfig、ModelConfig、ReasoningMode 类型
 * - useTestConnection — 测试连接共享 hook
 * - ReasoningModeSelect — 思考等级共享组件
 * - ModelConfigList — 聚合模型列表共享组件
 */

import React, { useState } from "react";
import {
  useSettingsStore,
  PROVIDER_TEMPLATES,
  API_FORMATS,
} from "../../store/settingsStore";
import type { AiProviderConfig, ModelConfig, ReasoningMode } from "../../electronApi";
import { ipcApi } from "../../services/ipcApi";
import {
  Eye,
  EyeOff,
  CheckCircle,
  XCircle,
  Loader2,
  ChevronDown,
  Zap,
  X,
} from "../common/IconMap";
import {
  MODEL_TEXT,
  DIRECT_TEMPLATES,
  AGGREGATION_TEMPLATES,
  OTHER_TEMPLATES,
} from "./modelSettingsI18n";
import { useTestConnection } from "./useTestConnection";
import { ReasoningModeSelect } from "./ReasoningModeSelect";
import { ModelConfigList } from "./ModelConfigList";
import {
  buildReasoningOptions,
  coerceReasoningMode,
  defaultReasoningModeForOptions,
  formatReasoningOptionLabels,
  resolveReasoningOptionValues,
} from "../../utils/reasoningSupport";

// ============================================================
// 类型定义
// ============================================================

export interface AddProviderDialogProps {
  onAdd: (config: AiProviderConfig) => void;
  onClose: () => void;
  generateId: () => string;
}

// ============================================================
// 组件实现
// ============================================================

export const AddProviderDialog: React.FC<AddProviderDialogProps> = ({ onAdd, onClose, generateId }) => {
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

  // 当前选中的模板
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
  const reasoningAutoHint = language === "zh-CN"
    ? `已根据当前供应商/API/模型自动适配：${formatReasoningOptionLabels(reasoningOptionValues, language)}`
    : `Automatically adapted for this provider/API/model: ${formatReasoningOptionLabels(reasoningOptionValues, language)}`;

  // 选择模板后自动填充
  const handleSelectTemplate = (templateId: string) => {
    const template = PROVIDER_TEMPLATES.find((t) => t.id === templateId);
    if (!template) {
      setSelectedTemplateId("");
      setName("");
      setApiFormat("openai");
      setBaseUrl("");
      setModel("");
      setContextWindowSize(undefined);
      setReasoningMode("off");
      setModelConfigs([]);
      return;
    }
    setSelectedTemplateId(templateId);
    setName(template.name);
    setApiFormat(template.apiFormat);
    setBaseUrl(template.baseUrl);
    setModel(template.defaultModel);
    setContextWindowSize(template.defaultContextWindowSize);
    setReasoningMode(template.defaultReasoningMode || "off");
    setModelConfigs([]);
  };

  // 测试模型可用性
  const handleTestModel = () => {
    if (!apiKey || !baseUrl || !model) return;
    testConnection(baseUrl, apiKey, apiFormat, model);
  };

  // 测试并添加
  const handleAddWithTest = async () => {
    if (!apiKey || !baseUrl) return;
    const resolvedModel = model || selectedTemplate?.defaultModel || "";
    // 注意：testConnection 是异步的，需要等待结果
    // 但 useTestConnection 的 testConnection 只更新内部状态，不返回结果
    // 所以这里保持原有的内联调用方式
    try {
      const result = await ipcApi.ai.testConnection(
        baseUrl,
        apiKey,
        apiFormat,
        resolvedModel
      );
      if (result.success) {
        handleAdd(resolvedModel);
      }
    } catch {
      // 静默失败，用户可通过独立测试按钮查看错误
    }
  };

  // 直接添加（不测试）
  const handleAdd = (resolvedModel?: string) => {
    const finalModel = resolvedModel || model || selectedTemplate?.defaultModel || "";
    const config: AiProviderConfig = {
      id: generateId(),
      name: name || (selectedTemplate?.name || text.customProvider),
      provider: selectedTemplate?.provider || "custom",
      apiKey,
      baseUrl,
      model: finalModel,
      defaultBaseUrl: selectedTemplate?.baseUrl || baseUrl,
      defaultModel: selectedTemplate?.defaultModel || "",
      enableReasoning: effectiveReasoningMode !== "off" ? true : undefined,
      reasoningMode: effectiveReasoningMode,
      apiFormat,
      models: selectedTemplate?.presetModels || undefined,
      modelConfigs: modelConfigs.length > 0 ? modelConfigs : undefined,
      contextWindowSize: contextWindowSize && contextWindowSize > 0 ? contextWindowSize : undefined,
    };
    onAdd(config);
  };

  const canAdd = apiKey.trim().length > 0 && baseUrl.trim().length > 0;
  const canTestModel = canAdd && model.trim().length > 0;

  return (
    <div className="dialog-overlay" onClick={onClose}>
      <div className="dialog add-provider-dialog" onClick={(e) => e.stopPropagation()}>
        {/* 标题栏 */}
        <div className="dialog-header">
          <h3>{text.addDialogTitle}</h3>
          <button className="dialog-close" onClick={onClose}><X size={16} /></button>
        </div>

        {/* 表单内容 */}
        <div className="dialog-body">
          {/* 供应商类型下拉选择 */}
          <div className="form-group">
            <label className="form-label">{text.providerType}</label>
            <select
              className="form-input template-select"
              value={selectedTemplateId}
              onChange={(e) => handleSelectTemplate(e.target.value)}
            >
              <option value="">{text.customProvider}</option>
              <optgroup label={text.directProviders}>
                {DIRECT_TEMPLATES.map((t) => (
                  <option key={t.id} value={t.id}>{t.name}</option>
                ))}
              </optgroup>
              <optgroup label={text.aggregationProviders}>
                {AGGREGATION_TEMPLATES.map((t) => (
                  <option key={t.id} value={t.id}>{t.name}</option>
                ))}
              </optgroup>
              <optgroup label={text.other}>
                {OTHER_TEMPLATES.map((t) => (
                  <option key={t.id} value={t.id}>{t.name}</option>
                ))}
              </optgroup>
            </select>
          </div>

          {/* 配置表单 */}
          <div className="add-provider-form">
            <div className="form-group">
              <label>{text.providerName}</label>
              <input
                type="text"
                className="form-input"
                value={name}
                onChange={(e) => setName(e.target.value)}
                placeholder={text.providerNamePlaceholder}
              />
            </div>

            <div className="form-group">
              <label>{text.apiFormat}</label>
              <select
                className="form-input"
                value={apiFormat}
                onChange={(e) => setApiFormat(e.target.value)}
              >
                {API_FORMATS.map((f) => (
                  <option key={f.value} value={f.value}>{f.label}</option>
                ))}
              </select>
            </div>

            <div className="form-group">
              <label>{text.apiUrl}</label>
              <input
                type="text"
                className="form-input"
                value={baseUrl}
                onChange={(e) => setBaseUrl(e.target.value)}
                placeholder="https://api.example.com/v1"
              />
            </div>

            <div className="form-group">
              <label>{text.apiKey}</label>
              <div className="input-with-toggle">
                <input
                  type={showApiKey ? "text" : "password"}
                  className="form-input"
                  value={apiKey}
                  onChange={(e) => setApiKey(e.target.value)}
                  placeholder="sk-..."
                />
                <button
                  className="toggle-visibility"
                  onClick={() => setShowApiKey(!showApiKey)}
                >
                  {showApiKey ? <EyeOff size={14} /> : <Eye size={14} />}
                </button>
              </div>
            </div>

            {/* 模型选择/输入 */}
            <div className="form-group">
              <label>{text.model}</label>
              {isAggregation ? (
                // 聚合类供应商：当前模型选择（从 modelConfigs 中选择）
                <div className="model-select-wrapper">
                  <select
                    className="form-input model-select"
                    value={model}
                    onChange={(e) => setModel(e.target.value)}
                  >
                    <option value="">-- {text.noModel} --</option>
                    {modelConfigs.map((m) => (
                      <option key={m.name} value={m.name}>{m.name}</option>
                    ))}
                  </select>
                  <ChevronDown size={14} className="select-arrow" />
                </div>
              ) : selectedTemplate?.presetModels && selectedTemplate.presetModels.length > 0 ? (
                // 直连供应商：下拉选择
                <div className="model-select-wrapper">
                  <select
                    className="form-input model-select"
                    value={model}
                    onChange={(e) => setModel(e.target.value)}
                  >
                    <option value="">-- {text.noModel} --</option>
                    {selectedTemplate.presetModels.map((m) => (
                      <option key={m} value={m}>{m}</option>
                    ))}
                  </select>
                  <ChevronDown size={14} className="select-arrow" />
                </div>
              ) : (
                // 自定义/无预设：文本输入
                <input
                  type="text"
                  className="form-input"
                  value={model}
                  onChange={(e) => setModel(e.target.value)}
                  placeholder="model-name"
                />
              )}
            </div>

            {/* 上下文窗口大小 */}
            <div className="form-group">
              <label>{text.contextWindowSize}</label>
              <div className="input-with-action">
                <input
                  type="number"
                  className="form-input context-window-input"
                  value={contextWindowSize ?? ""}
                  onChange={(e) => {
                    const val = e.target.value ? parseInt(e.target.value, 10) : undefined;
                    setContextWindowSize(val && val > 0 ? val : undefined);
                  }}
                  placeholder={text.contextWindowPlaceholder}
                  min={1000}
                  step={1000}
                />
                <button
                  className="btn-test"
                  onClick={handleTestModel}
                  disabled={!canTestModel || testing}
                  title={text.testModel}
                >
                  {testing ? <Loader2 size={14} className="spin" /> : <Zap size={14} />}
                  {text.testModel}
                </button>
              </div>
              {selectedTemplate?.defaultContextWindowSize && (
                <span className="form-hint">
                  {text.contextWindowHint(selectedTemplate.defaultContextWindowSize)}
                </span>
              )}
            </div>

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

          {/* 测试结果 */}
          {testResult && (
            <div className={`test-result ${testResult.success ? "success" : "error"}`}>
              {testResult.success ? (
                <><CheckCircle size={14} /> {text.connectionOk(testResult.latency)}</>
              ) : (
                <><XCircle size={14} /> {testResult.error || text.connectionFailed}</>
              )}
            </div>
          )}
        </div>

        {/* 操作按钮 */}
        <div className="dialog-actions">
          <button className="btn-secondary" onClick={onClose}>{text.cancel}</button>
          <button
            className="btn-primary"
            onClick={() => handleAdd()}
            disabled={!canAdd}
          >
            {text.add}
          </button>
          <button
            className="btn-primary"
            onClick={handleAddWithTest}
            disabled={!canAdd || testing}
          >
            {testing ? <><Loader2 size={14} className="spin" /> {text.testing}</> : <><Zap size={14} /> {text.addAndTest}</>}
          </button>
        </div>
      </div>
    </div>
  );
};
