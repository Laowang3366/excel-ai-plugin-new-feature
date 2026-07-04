/**
 * EditProviderDialog — 编辑供应商弹窗子组件
 *
 * 以弹窗形式编辑供应商配置，所有修改保存在本地草稿状态，
 * 点击"保存"按钮后一次性提交到 settingsStore。
 *
 * 关联模块：
 * - modelSettingsI18n.ts — 提供 MODEL_TEXT 双语文本
 * - settingsStore — PROVIDER_TEMPLATES、API_FORMATS、useSettingsStore
 * - useTestConnection — 测试连接共享 hook
 * - ReasoningModeSelect — 思考等级共享组件
 * - ModelConfigList — 聚合模型列表共享组件
 */

import React, { useState, useCallback } from "react";
import { useSettingsStore, PROVIDER_TEMPLATES, API_FORMATS } from "../../store/settingsStore";
import type { AiProviderConfig, ModelConfig, ReasoningMode } from "../../electronApi";
import {
  Eye,
  EyeOff,
  CheckCircle,
  XCircle,
  Loader2,
  ChevronDown,
  RotateCcw,
  Zap,
  X,
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
  formatReasoningOptionLabels,
  resolveReasoningOptionValues,
} from "../../utils/reasoningSupport";

// ============================================================
// 类型定义
// ============================================================

export interface EditProviderDialogProps {
  provider: AiProviderConfig;
  onSave: (patch: Partial<AiProviderConfig>) => void;
  onClose: () => void;
}

// ============================================================
// 组件实现
// ============================================================

export const EditProviderDialog: React.FC<EditProviderDialogProps> = ({
  provider,
  onSave,
  onClose,
}) => {
  const { language } = useSettingsStore();
  const text = MODEL_TEXT[language];

  // 本地草稿状态 — 所有修改保存在此，不直接写入 store
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

  // 判断是否为聚合类供应商（无预设模型）
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
  const reasoningAutoHint = language === "zh-CN"
    ? `已根据当前供应商/API/模型自动适配：${formatReasoningOptionLabels(reasoningOptionValues, language)}`
    : `Automatically adapted for this provider/API/model: ${formatReasoningOptionLabels(reasoningOptionValues, language)}`;

  // 可选模型列表
  const availableModels = provider.models || [];
  const presetModels = template?.presetModels || [];
  const modelOptions = availableModels.length > 0 ? availableModels : presetModels;

  // 获取模型列表
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

  // 测试连接
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

  // 保存：收集所有变更并提交
  const handleSave = () => {
    const patch: Partial<AiProviderConfig> = {};
    const normalizedModelConfigs = modelConfigs.map((modelConfig) => {
      const { reasoningOptions: _legacyReasoningOptions, ...rest } = modelConfig;
      return rest;
    });
    if (name !== provider.name) patch.name = name;
    if (apiFormat !== (provider.apiFormat || "openai")) patch.apiFormat = apiFormat;
    if (baseUrl !== provider.baseUrl) patch.baseUrl = baseUrl;
    if (apiKey !== provider.apiKey) patch.apiKey = apiKey;
    if (model !== provider.model) patch.model = model;
    if (contextWindowSize !== provider.contextWindowSize) patch.contextWindowSize = contextWindowSize;
    if (effectiveReasoningMode !== (provider.reasoningMode || "off")) patch.reasoningMode = effectiveReasoningMode;
    if (JSON.stringify(normalizedModelConfigs) !== JSON.stringify(provider.modelConfigs || [])) {
      patch.modelConfigs = normalizedModelConfigs;
    }
    onSave(patch);
  };

  return (
    <div className="dialog-overlay" onClick={onClose}>
      <div className="dialog edit-provider-dialog" onClick={(e) => e.stopPropagation()}>
        {/* 标题栏 */}
        <div className="dialog-header">
          <h3>{text.editDialogTitle(name || provider.name)}</h3>
          <button className="dialog-close" onClick={onClose}><X size={16} /></button>
        </div>

        {/* 表单内容 */}
        <div className="dialog-body">
          {/* 供应商名称 */}
          <div className="form-group">
            <label>供应商名称</label>
            <input
              type="text"
              className="form-input"
              value={name}
              onChange={(e) => setName(e.target.value)}
              placeholder="如：我的 DeepSeek"
            />
          </div>

          {/* API 格式 */}
          <div className="form-group">
            <label>API 格式</label>
            <select
              className="form-input"
              value={apiFormat}
              onChange={(e) => setApiFormat(e.target.value)}
            >
              {API_FORMATS.map((f) => (
                <option key={f.value} value={f.value}>{f.label}</option>
              ))}
            </select>
            <span className="form-hint">选择 API 协议格式，影响请求方式和认证方式</span>
          </div>

          {/* API 地址 */}
          <div className="form-group">
            <label>API 地址</label>
            <div className="input-with-action">
              <input
                type="text"
                className="form-input"
                value={baseUrl}
                onChange={(e) => setBaseUrl(e.target.value)}
                placeholder="https://api.example.com/v1"
              />
              {provider.defaultBaseUrl && (
                <button
                  className="btn-icon-hint"
                  onClick={() => setBaseUrl(provider.defaultBaseUrl!)}
                  title={`恢复默认: ${provider.defaultBaseUrl}`}
                >
                  <RotateCcw size={13} />
                </button>
              )}
            </div>
            {provider.defaultBaseUrl && (
              <span className="form-hint">{text.defaultValue}: {provider.defaultBaseUrl}</span>
            )}
          </div>

          {/* API 密钥 */}
          <div className="form-group">
            <label>{text.apiKey}</label>
            <div className="input-with-action">
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
              <button
                className="btn-test"
                onClick={handleTestConnection}
                disabled={testing || !apiKey || !baseUrl}
                title={text.testConnection}
              >
                {testing ? <Loader2 size={14} className="spin" /> : <Zap size={14} />}
                {text.test}
              </button>
            </div>
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

          {/* 模型选择/输入 */}
          <div className="form-group">
            <label>{text.model}</label>
            <div className="input-with-action">
              {isAggregation ? (
                <div className="model-select-wrapper">
                  <select
                    className="form-input model-select"
                    value={model}
                    onChange={(e) => {
                      const newModel = e.target.value;
                      setModel(newModel);
                      applyModelConfig(newModel);
                    }}
                  >
                    {!model && <option value="">-- {text.noModel} --</option>}
                    {modelConfigs.map((m) => (
                      <option key={m.name} value={m.name}>{m.name}</option>
                    ))}
                    {model && !modelConfigs.some(m => m.name === model) && (
                      <option value={model}>{model}</option>
                    )}
                  </select>
                  <ChevronDown size={14} className="select-arrow" />
                </div>
              ) : modelOptions.length > 0 ? (
                <div className="model-select-wrapper">
                  <select
                    className="form-input model-select"
                    value={model}
                    onChange={(e) => {
                      const newModel = e.target.value;
                      setModel(newModel);
                      applyModelConfig(newModel);
                    }}
                  >
                    {model && !modelOptions.includes(model) && (
                      <option value={model}>{model}</option>
                    )}
                    {modelOptions.map((m) => (
                      <option key={m} value={m}>{m}</option>
                    ))}
                  </select>
                  <ChevronDown size={14} className="select-arrow" />
                </div>
              ) : (
                <>
                  <input
                    type="text"
                    className="form-input"
                    value={model}
                    onChange={(e) => {
                      const newModel = e.target.value;
                      setModel(newModel);
                      applyModelConfig(newModel);
                    }}
                    placeholder={provider.defaultModel || "model-name"}
                  />
                  <button
                    className="btn-refresh"
                    onClick={fetchModels}
                    disabled={fetchingModels || !apiKey || !baseUrl}
                    title={text.refreshModels}
                  >
                    {fetchingModels ? <Loader2 size={14} className="spin" /> : <RefreshCw size={14} />}
                  </button>
                </>
              )}
              {!isAggregation && modelOptions.length > 0 && (
                <button
                  className="btn-refresh"
                  onClick={fetchModels}
                  disabled={fetchingModels || !apiKey || !baseUrl}
                  title={text.refreshModels}
                >
                  {fetchingModels ? <Loader2 size={14} className="spin" /> : <RefreshCw size={14} />}
                </button>
              )}
            </div>
            {provider.defaultModel && (
              <span className="form-hint">{text.defaultValue}: {provider.defaultModel}</span>
            )}
            {modelOptions.length > 0 && (
              <span className="form-hint">{text.availableModels(modelOptions.length)}</span>
            )}
          </div>

          {/* 上下文窗口大小 */}
          <div className="form-group">
            <label>{text.contextWindowSize}</label>
            <div className="input-with-action">
              <input
                type="number"
                className="form-input context-window-input"
                value={contextWindowSize || ""}
                onChange={(e) => {
                  const val = e.target.value ? parseInt(e.target.value, 10) : undefined;
                  setContextWindowSize(val && val > 0 ? val : undefined);
                }}
                placeholder={text.contextWindowPlaceholder}
                min={1000}
                step={1000}
              />
              {template?.defaultContextWindowSize && (
                <button
                  className="btn-icon-hint"
                  onClick={() => setContextWindowSize(template.defaultContextWindowSize)}
                  title={`恢复默认: ${formatTokensAsK(template.defaultContextWindowSize)}`}
                >
                  <RotateCcw size={13} />
                </button>
              )}
            </div>
            {template?.defaultContextWindowSize && (
              <span className="form-hint">{text.contextWindowHint(template.defaultContextWindowSize)}</span>
            )}
          </div>

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
        </div>

        {/* 操作按钮：测试 + 保存 + 取消 */}
        <div className="dialog-actions">
          <button
            className="btn-test-full"
            onClick={handleTestConnection}
            disabled={testing || !apiKey || !baseUrl}
          >
            {testing ? <Loader2 size={14} className="spin" /> : <Zap size={14} />}
            {text.testConnection}
          </button>
          <div className="dialog-actions-right">
            <button className="btn-secondary" onClick={onClose}>{text.cancel}</button>
            <button className="btn-primary" onClick={handleSave}>
              {text.save}
            </button>
          </div>
        </div>
      </div>
    </div>
  );
};
