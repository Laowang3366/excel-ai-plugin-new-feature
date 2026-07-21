import { useState } from "react";
import {
  buildReportTaskPayload,
  REPORT_OUTPUT_FORMAT_LABELS,
  type ReportOutputFormat,
} from "@shared/tasks";
import { useTaskSubmit } from "./useTaskSubmit";
import { useSelectionAddress } from "./useSelectionAddress";
import { TaskSubmitStatus } from "./TaskSubmitStatus";

export function ReportTaskPanel() {
  const { submit, busy, error, lastResult, adapter } = useTaskSubmit();
  const { readSelection, busy: picking, error: pickError } =
    useSelectionAddress(adapter);
  const [range, setRange] = useState("");
  const [task, setTask] = useState("");
  const [outputFormat, setOutputFormat] =
    useState<ReportOutputFormat>("excel");

  return (
    <section className="card task-panel">
      <h2>报告生成</h2>
      <p className="muted">
        默认在当前工作簿新增/更新报告工作表；不支持磁盘路径与 Word/PPT 输出。
      </p>
      <div className="task-field">
        <label>
          数据源选区
          <div className="row">
            <input
              value={range}
              onChange={(e) => setRange(e.target.value)}
              aria-label="报告数据源"
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
          输出格式
          <select
            value={outputFormat}
            onChange={(e) =>
              setOutputFormat(e.target.value as ReportOutputFormat)
            }
          >
            <option value="excel">{REPORT_OUTPUT_FORMAT_LABELS.excel}</option>
            <option value="word" disabled>
              {REPORT_OUTPUT_FORMAT_LABELS.word}（加载项不支持）
            </option>
            <option value="ppt" disabled>
              {REPORT_OUTPUT_FORMAT_LABELS.ppt}（加载项不支持）
            </option>
          </select>
        </label>
      </div>
      <div className="task-field">
        <label>
          需求说明
          <textarea
            rows={3}
            value={task}
            onChange={(e) => setTask(e.target.value)}
            placeholder="报告结构、指标、工作表命名…"
          />
        </label>
      </div>
      <button
        type="button"
        disabled={busy || !adapter || outputFormat !== "excel"}
        onClick={() =>
          void submit(buildReportTaskPayload({ range, task, outputFormat }))
        }
      >
        提交到 AI
      </button>
      <TaskSubmitStatus busy={busy} error={error || pickError} lastResult={lastResult} />
    </section>
  );
}
