import { useState } from "react";
import { buildChartTaskPayload } from "@shared/tasks";
import { useTaskSubmit } from "./useTaskSubmit";
import { useSelectionAddress } from "./useSelectionAddress";
import { TaskSubmitStatus } from "./TaskSubmitStatus";

const CHART_TYPES = [
  "column",
  "bar",
  "line",
  "pie",
  "area",
  "scatter",
  "doughnut",
] as const;

export function ChartTaskPanel() {
  const { submit, busy, error, lastResult, adapter } = useTaskSubmit();
  const { readSelection, busy: picking, error: pickError } =
    useSelectionAddress(adapter);
  const [range, setRange] = useState("");
  const [task, setTask] = useState("");
  const [chartType, setChartType] = useState<string>("column");
  const [title, setTitle] = useState("");
  const [showLegend, setShowLegend] = useState(true);
  const [positionNote, setPositionNote] = useState("");
  const isWps = adapter?.kind === "wps-jsa";

  return (
    <section className="card task-panel">
      <h2>图表制作</h2>
      <p className="muted">基于选区创建图表；Office.js 走 chart.* 工具。</p>
      {isWps && (
        <div className="chat-banner error" role="note">
          当前宿主为 WPS JSA：仓库内无图表合同，chart 工具可能返回 typed
          unsupported。入口可打开并提交说明，不会猜测 WPS 图表 API。
        </div>
      )}
      <div className="task-field">
        <label>
          数据源选区
          <div className="row">
            <input
              value={range}
              onChange={(e) => setRange(e.target.value)}
              aria-label="图表数据源"
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
      <div className="task-field">
        <label>
          图表类型
          <select
            value={chartType}
            onChange={(e) => setChartType(e.target.value)}
          >
            {CHART_TYPES.map((t) => (
              <option key={t} value={t}>
                {t}
              </option>
            ))}
          </select>
        </label>
        <label>
          标题
          <input value={title} onChange={(e) => setTitle(e.target.value)} />
        </label>
        <label className="task-check">
          <input
            type="checkbox"
            checked={showLegend}
            onChange={(e) => setShowLegend(e.target.checked)}
          />
          显示图例
        </label>
        <label>
          位置/布局说明
          <input
            value={positionNote}
            onChange={(e) => setPositionNote(e.target.value)}
            placeholder="可选"
          />
        </label>
      </div>
      <div className="task-field">
        <label>
          需求说明
          <textarea
            rows={3}
            value={task}
            onChange={(e) => setTask(e.target.value)}
          />
        </label>
      </div>
      <button
        type="button"
        disabled={busy || !adapter}
        onClick={() =>
          void submit(
            buildChartTaskPayload({
              range,
              task,
              chartType,
              title,
              showLegend,
              positionNote,
            }),
          )
        }
      >
        提交到 AI
      </button>
      <TaskSubmitStatus busy={busy} error={error || pickError} lastResult={lastResult} />
    </section>
  );
}
