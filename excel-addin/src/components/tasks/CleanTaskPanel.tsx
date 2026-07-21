import { useState } from "react";
import {
  buildCleanTaskPayload,
  CLEAN_MODE_LABELS,
  type CleanOpMode,
} from "@shared/tasks";
import { useTaskSubmit } from "./useTaskSubmit";
import { useSelectionAddress } from "./useSelectionAddress";
import { TaskSubmitStatus } from "./TaskSubmitStatus";

const ALL_MODES: CleanOpMode[] = ["drop_empty", "dedupe", "normalize"];

export function CleanTaskPanel() {
  const { submit, busy, error, lastResult, adapter } = useTaskSubmit();
  const { readSelection, busy: picking, error: pickError } =
    useSelectionAddress(adapter);
  const [range, setRange] = useState("");
  const [task, setTask] = useState("");
  const [modes, setModes] = useState<CleanOpMode[]>(["drop_empty", "dedupe"]);

  const toggleMode = (mode: CleanOpMode) => {
    setModes((prev) =>
      prev.includes(mode) ? prev.filter((m) => m !== mode) : [...prev, mode],
    );
  };

  return (
    <section className="card task-panel">
      <h2>数据清洗</h2>
      <p className="muted">
        当前工作簿内清洗（range 读写）；非 Power Query / 外部 ETL。
      </p>
      <div className="task-field">
        <label>
          数据源选区
          <div className="row">
            <input
              value={range}
              onChange={(e) => setRange(e.target.value)}
              aria-label="清洗数据源"
            />
            <button
              type="button"
              disabled={picking || !adapter}
              onClick={async () => {
                const addr = await readSelection();
                if (addr) setRange(addr);
              }}
            >
              读取选区
            </button>
          </div>
        </label>
      </div>
      <fieldset className="task-field">
        <legend>操作模式</legend>
        {ALL_MODES.map((mode) => (
          <label key={mode} className="task-check">
            <input
              type="checkbox"
              checked={modes.includes(mode)}
              onChange={() => toggleMode(mode)}
            />
            {CLEAN_MODE_LABELS[mode]}
          </label>
        ))}
      </fieldset>
      <div className="task-field">
        <label>
          清洗要求
          <textarea
            rows={3}
            value={task}
            onChange={(e) => setTask(e.target.value)}
            placeholder="例如：删除空行、按姓名去重、日期统一格式…"
          />
        </label>
      </div>
      <button
        type="button"
        disabled={busy || !adapter}
        onClick={() =>
          void submit(buildCleanTaskPayload({ range, task, modes }))
        }
      >
        提交到 AI
      </button>
      <TaskSubmitStatus busy={busy} error={error || pickError} lastResult={lastResult} />
    </section>
  );
}
