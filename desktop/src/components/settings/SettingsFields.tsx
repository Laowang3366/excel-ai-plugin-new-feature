import type { CSSProperties, ReactNode } from "react";

interface SettingsSwitchFieldProps {
  groupLabel?: string;
  label: string;
  checked: boolean;
  hint: string;
  onChange: (checked: boolean) => void;
}

interface SettingsSliderFieldProps {
  label: string;
  value: number;
  min: number;
  max: number;
  step: number;
  hint: string;
  valueText?: string;
  disabled?: boolean;
  className?: string;
  info?: ReactNode;
  onChange: (value: number) => void;
}

function getSliderFillPercent(value: number, min: number, max: number): number {
  if (max <= min) return 0;
  return Math.min(100, Math.max(0, ((value - min) / (max - min)) * 100));
}

function getSliderStyle(value: number, min: number, max: number): CSSProperties {
  return {
    "--slider-fill": `${getSliderFillPercent(value, min, max)}%`,
  } as CSSProperties;
}

export function SettingsSwitchField({
  groupLabel,
  label,
  checked,
  hint,
  onChange,
}: SettingsSwitchFieldProps) {
  return (
    <div className="form-group">
      {groupLabel && <label>{groupLabel}</label>}
      <label className="settings-switch-row">
        <input
          type="checkbox"
          checked={checked}
          onChange={(event) => onChange(event.target.checked)}
        />
        <span>{label}</span>
      </label>
      <span className="form-hint">{hint}</span>
    </div>
  );
}

export function SettingsSliderField({
  label,
  value,
  min,
  max,
  step,
  hint,
  valueText,
  disabled,
  className,
  info,
  onChange,
}: SettingsSliderFieldProps) {
  const fillPercent = getSliderFillPercent(value, min, max);

  return (
    <div className={`form-group${className ? ` ${className}` : ""}`}>
      <label>{label}</label>
      {info}
      <div className="compaction-threshold-row">
        <div className="compaction-slider-track">
          <input
            type="range"
            className="compaction-slider"
            min={min}
            max={max}
            step={step}
            value={value}
            aria-label={label}
            style={getSliderStyle(value, min, max)}
            onChange={(event) => onChange(Number(event.target.value))}
            disabled={disabled}
          />
          <span className={`compaction-threshold-value${fillPercent > 65 ? " over-fill" : ""}`}>
            {valueText ?? value}
          </span>
        </div>
      </div>
      <span className="form-hint">{hint}</span>
    </div>
  );
}
