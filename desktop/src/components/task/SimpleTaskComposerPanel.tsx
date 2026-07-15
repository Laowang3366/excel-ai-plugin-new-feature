import React from "react";
import { Ruler } from "../common/IconMap";

export type SimpleTaskIntent = "clean" | "chart";

export interface SimpleTaskComposerText {
  dataSourceRange: string;
  rangePlaceholder: string;
  pickRange: string;
  requirement: string;
  simplePlaceholders: Record<SimpleTaskIntent, string>;
  simplePrefixes: Record<SimpleTaskIntent, string>;
  sendToAi: string;
}

export interface SimpleTaskPayloadInput {
  prefix: string;
  rangeLabel: string;
  requirementLabel: string;
  range: string;
  task: string;
}

export function buildSimpleTaskPayload({
  prefix,
  rangeLabel,
  requirementLabel,
  range,
  task,
}: SimpleTaskPayloadInput): string {
  const lines = [prefix];
  if (range.trim()) lines.push(`${rangeLabel}: ${range.trim()}`);
  if (task.trim()) lines.push(`${requirementLabel}: ${task.trim()}`);
  return lines.join("\n");
}

interface SimpleTaskComposerPanelProps {
  intent: SimpleTaskIntent;
  range: string;
  task: string;
  text: SimpleTaskComposerText;
  onRangeChange: (range: string) => void;
  onTaskChange: (task: string) => void;
  onPickRange: (intent: SimpleTaskIntent) => void;
  onSubmit: (payload: string) => void;
}

export const SimpleTaskComposerPanel: React.FC<SimpleTaskComposerPanelProps> = ({
  intent,
  range,
  task,
  text,
  onRangeChange,
  onTaskChange,
  onPickRange,
  onSubmit,
}) => {
  return (
    <div className="task-composer-panel">
      <div className="task-field">
        <label className="task-field-label">{text.dataSourceRange}</label>
        <div className="range-input-row">
          <input
            className="task-field-input"
            placeholder={text.rangePlaceholder}
            value={range}
            onChange={(event) => onRangeChange(event.target.value)}
          />
          <button className="btn-pick-range" onClick={() => onPickRange(intent)}>
            <Ruler size={13} /> {text.pickRange}
          </button>
        </div>
      </div>
      <div className="task-field">
        <label className="task-field-label">{text.requirement}</label>
        <textarea
          className="task-field-textarea"
          value={task}
          onChange={(event) => onTaskChange(event.target.value)}
          placeholder={text.simplePlaceholders[intent]}
        />
      </div>
      <button
        className="task-submit-btn"
        onClick={() =>
          onSubmit(
            buildSimpleTaskPayload({
              prefix: text.simplePrefixes[intent],
              rangeLabel: text.dataSourceRange,
              requirementLabel: text.requirement,
              range,
              task,
            }),
          )
        }
      >
        {text.sendToAi}
      </button>
    </div>
  );
};
