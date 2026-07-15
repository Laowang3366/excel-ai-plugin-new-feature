import React from "react";
import type { TestResult } from "./useTestConnection";
import { CheckCircle, Eye, EyeOff, Loader2, Zap, XCircle } from "../common/IconMap";

export interface ApiFormatOption {
  value: string;
  label: string;
}

interface ProviderNameFieldProps {
  label: string;
  value: string;
  placeholder: string;
  onChange: (value: string) => void;
}

export function ProviderNameField({ label, value, placeholder, onChange }: ProviderNameFieldProps) {
  return (
    <div className="form-group">
      <label>{label}</label>
      <input
        type="text"
        className="form-input"
        value={value}
        onChange={(event) => onChange(event.target.value)}
        placeholder={placeholder}
      />
    </div>
  );
}

interface ProviderApiFormatFieldProps {
  label: string;
  value: string;
  options: readonly ApiFormatOption[];
  hint?: string;
  onChange: (value: string) => void;
}

export function ProviderApiFormatField({
  label,
  value,
  options,
  hint,
  onChange,
}: ProviderApiFormatFieldProps) {
  return (
    <div className="form-group">
      <label>{label}</label>
      <select
        className="form-input"
        value={value}
        onChange={(event) => onChange(event.target.value)}
      >
        {options.map((option) => (
          <option key={option.value} value={option.value}>
            {option.label}
          </option>
        ))}
      </select>
      {hint && <span className="form-hint">{hint}</span>}
    </div>
  );
}

interface ProviderBaseUrlFieldProps {
  label: string;
  value: string;
  placeholder: string;
  hint?: string;
  action?: React.ReactNode;
  onChange: (value: string) => void;
}

export function ProviderBaseUrlField({
  label,
  value,
  placeholder,
  hint,
  action,
  onChange,
}: ProviderBaseUrlFieldProps) {
  return (
    <div className="form-group">
      <label>{label}</label>
      {action ? (
        <div className="input-with-action">
          <input
            type="text"
            className="form-input"
            value={value}
            onChange={(event) => onChange(event.target.value)}
            placeholder={placeholder}
          />
          {action}
        </div>
      ) : (
        <input
          type="text"
          className="form-input"
          value={value}
          onChange={(event) => onChange(event.target.value)}
          placeholder={placeholder}
        />
      )}
      {hint && <span className="form-hint">{hint}</span>}
    </div>
  );
}

interface ProviderApiKeyFieldProps {
  label: string;
  value: string;
  showApiKey: boolean;
  action?: React.ReactNode;
  onChange: (value: string) => void;
  onToggleVisibility: () => void;
}

export function ProviderApiKeyField({
  label,
  value,
  showApiKey,
  action,
  onChange,
  onToggleVisibility,
}: ProviderApiKeyFieldProps) {
  const input = (
    <div className="input-with-toggle">
      <input
        type={showApiKey ? "text" : "password"}
        className="form-input"
        value={value}
        onChange={(event) => onChange(event.target.value)}
        placeholder="sk-..."
      />
      <button className="toggle-visibility" onClick={onToggleVisibility}>
        {showApiKey ? <EyeOff size={14} /> : <Eye size={14} />}
      </button>
    </div>
  );

  return (
    <div className="form-group">
      <label>{label}</label>
      {action ? (
        <div className="input-with-action">
          {input}
          {action}
        </div>
      ) : (
        input
      )}
    </div>
  );
}

interface ProviderContextWindowFieldProps {
  label: string;
  value: number | undefined;
  placeholder: string;
  hint?: string;
  action?: React.ReactNode;
  onChange: (value: number | undefined) => void;
}

export function ProviderContextWindowField({
  label,
  value,
  placeholder,
  hint,
  action,
  onChange,
}: ProviderContextWindowFieldProps) {
  return (
    <div className="form-group">
      <label>{label}</label>
      <div className="input-with-action">
        <input
          type="number"
          className="form-input context-window-input"
          value={value ?? ""}
          onChange={(event) => {
            const parsed = event.target.value ? parseInt(event.target.value, 10) : undefined;
            onChange(parsed && parsed > 0 ? parsed : undefined);
          }}
          placeholder={placeholder}
          min={1000}
          step={1000}
        />
        {action}
      </div>
      {hint && <span className="form-hint">{hint}</span>}
    </div>
  );
}

interface ProviderModelFieldProps {
  label: string;
  selector: React.ReactNode;
  action?: React.ReactNode;
  hints?: string[];
}

export function ProviderModelField({
  label,
  selector,
  action,
  hints = [],
}: ProviderModelFieldProps) {
  return (
    <div className="form-group">
      <label>{label}</label>
      {action ? (
        <div className="input-with-action">
          {selector}
          {action}
        </div>
      ) : (
        selector
      )}
      {hints.map((hint) => (
        <span key={hint} className="form-hint">
          {hint}
        </span>
      ))}
    </div>
  );
}

interface ProviderTestResultProps {
  result: TestResult | null;
  successText: (latency?: number) => string;
  errorFallback: string;
}

interface ProviderTestButtonProps {
  className: string;
  testing: boolean;
  label: string;
  testingLabel?: string;
  disabled: boolean;
  title?: string;
  onClick: () => void;
}

export function ProviderTestButton({
  className,
  testing,
  label,
  testingLabel,
  disabled,
  title,
  onClick,
}: ProviderTestButtonProps) {
  return (
    <button className={className} onClick={onClick} disabled={disabled} title={title}>
      {testing ? (
        <>
          <Loader2 size={14} className="spin" /> {testingLabel ?? label}
        </>
      ) : (
        <>
          <Zap size={14} /> {label}
        </>
      )}
    </button>
  );
}

export function ProviderTestResult({
  result,
  successText,
  errorFallback,
}: ProviderTestResultProps) {
  if (!result) return null;
  return (
    <div className={`test-result ${result.success ? "success" : "error"}`}>
      {result.success ? (
        <>
          <CheckCircle size={14} /> {successText(result.latency)}
        </>
      ) : (
        <>
          <XCircle size={14} /> {result.error || errorFallback}
        </>
      )}
    </div>
  );
}
