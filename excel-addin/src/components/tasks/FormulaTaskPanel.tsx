import { useEffect, useState } from "react";
import {
  buildFormulaTaskPayload,
  normalizeHostEnvironment,
  type HostEnvironment,
  type ReferenceSampleMode,
} from "@shared/tasks";
import { useTaskSubmit } from "./useTaskSubmit";
import { useSelectionAddress } from "./useSelectionAddress";
import { TaskSubmitStatus } from "./TaskSubmitStatus";

export function FormulaTaskPanel() {
  const { submit, busy, error, lastResult, adapter } = useTaskSubmit();
  const { readSelection, busy: picking, error: pickError } =
    useSelectionAddress(adapter);
  const [dataSourceRanges, setDataSourceRanges] = useState<string[]>([]);
  const [dataSourceInput, setDataSourceInput] = useState("");
  const [referenceSampleRange, setReferenceSampleRange] = useState("");
  const [referenceSampleMode, setReferenceSampleMode] =
    useState<ReferenceSampleMode>("partial");
  const [outputRange, setOutputRange] = useState("");
  const [hostEnvironment, setHostEnvironment] =
    useState<HostEnvironment>("unknown");
  const [task, setTask] = useState("");

  useEffect(() => {
    let cancelled = false;
    void (async () => {
      if (!adapter) return;
      const status = await adapter.getStatus();
      if (cancelled || !status.ok) return;
      setHostEnvironment(
        normalizeHostEnvironment({
          connected: status.data.connected,
          host: status.data.hostName ?? status.data.kind,
          workbookName: status.data.workbookName ?? undefined,
        }),
      );
    })();
    return () => {
      cancelled = true;
    };
  }, [adapter]);

  const addDataSource = async () => {
    const addr = await readSelection();
    if (addr && !dataSourceRanges.includes(addr)) {
      setDataSourceRanges((prev) => [...prev, addr]);
    }
  };

  const onSubmit = async () => {
    const payload = buildFormulaTaskPayload({
      dataSourceRanges,
      dataSourceInput,
      referenceSampleRange,
      referenceSampleMode,
      outputRange,
      hostEnvironment,
      task,
    });
    await submit(payload);
  };

  return (
    <section className="card task-panel">
      <h2>公式助手</h2>
      <p className="muted">生成 Excel/WPS 函数公式；提交后进入共享聊天会话。</p>
      <div className="task-field">
        <label>
          数据源选区
          <div className="row">
            <input
              value={dataSourceInput}
              onChange={(e) => setDataSourceInput(e.target.value)}
              placeholder="可手输 A1 或点选"
              aria-label="数据源输入"
            />
            <button type="button" onClick={() => void addDataSource()} disabled={picking || !adapter}>
              读取选区
            </button>
          </div>
        </label>
        {dataSourceRanges.length > 0 && (
          <ul className="task-chip-list">
            {dataSourceRanges.map((r) => (
              <li key={r}>
                <button
                  type="button"
                  className="task-chip"
                  onClick={() =>
                    setDataSourceRanges((prev) => prev.filter((x) => x !== r))
                  }
                >
                  {r} ×
                </button>
              </li>
            ))}
          </ul>
        )}
      </div>
      <div className="task-field">
        <label>
          参考样例选区
          <div className="row">
            <input
              value={referenceSampleRange}
              onChange={(e) => setReferenceSampleRange(e.target.value)}
              aria-label="参考样例"
            />
            <button
              type="button"
              onClick={async () => {
                const addr = await readSelection();
                if (addr) setReferenceSampleRange(addr);
              }}
              disabled={picking || !adapter}
            >
              读取选区
            </button>
          </div>
        </label>
        <label>
          样例类型
          <select
            value={referenceSampleMode}
            onChange={(e) =>
              setReferenceSampleMode(e.target.value as ReferenceSampleMode)
            }
          >
            <option value="partial">部分样例</option>
            <option value="complete">完整样例</option>
          </select>
        </label>
      </div>
      <div className="task-field">
        <label>
          输出锚点
          <div className="row">
            <input
              value={outputRange}
              onChange={(e) => setOutputRange(e.target.value)}
              aria-label="输出锚点"
            />
            <button
              type="button"
              onClick={async () => {
                const addr = await readSelection();
                if (addr) setOutputRange(addr);
              }}
              disabled={picking || !adapter}
            >
              读取选区
            </button>
          </div>
        </label>
      </div>
      <div className="task-field">
        <label>
          需求说明
          <textarea
            rows={3}
            value={task}
            onChange={(e) => setTask(e.target.value)}
            placeholder="描述要生成的公式…"
          />
        </label>
      </div>
      <p className="muted">环境：{hostEnvironment}</p>
      <button type="button" onClick={() => void onSubmit()} disabled={busy || !adapter}>
        提交到 AI
      </button>
      <TaskSubmitStatus busy={busy} error={error || pickError} lastResult={lastResult} />
    </section>
  );
}
