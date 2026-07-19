import { useMemo, useState } from "react";
import type { HostAdapter } from "@shared/host";
import { ToolExecutor, TOOL_DEFINITIONS, type ToolName } from "@shared/tools";

interface Props {
  adapter: HostAdapter;
}

export function ToolDemoPanel({ adapter }: Props) {
  const executor = useMemo(() => new ToolExecutor(adapter), [adapter]);
  const [tool, setTool] = useState<ToolName>("selection.get");
  const [sheetName, setSheetName] = useState("Sheet1");
  const [range, setRange] = useState("A1");
  const [valuesJson, setValuesJson] = useState('[["hello"]]');
  const [formula, setFormula] = useState("=SUM(1,2)");
  const [newName, setNewName] = useState("Sheet2");
  const [output, setOutput] = useState<string>("");

  async function run() {
    const args: Record<string, unknown> = {};
    if (
      tool === "range.read" ||
      tool === "range.write" ||
      tool === "range.clear" ||
      tool === "formula.read" ||
      tool === "formula.write" ||
      tool === "sheet.add" ||
      tool === "sheet.rename" ||
      tool === "sheet.delete"
    ) {
      args.sheetName = sheetName;
    }
    if (
      tool === "range.read" ||
      tool === "range.write" ||
      tool === "range.clear" ||
      tool === "formula.read" ||
      tool === "formula.write"
    ) {
      args.range = range;
    }
    if (tool === "range.write") {
      args.values = JSON.parse(valuesJson) as unknown;
      args.verify = true;
    }
    if (tool === "formula.write") {
      args.formula = formula;
      args.verify = true;
    }
    if (tool === "sheet.rename") {
      args.newName = newName;
    }

    const result = await executor.execute({ name: tool, arguments: args });
    setOutput(JSON.stringify(result, null, 2));
  }

  return (
    <section className="card">
      <h2>Excel 工具演示（Phase 1–36）</h2>
      <div className="row">
        <label>
          工具
          <select value={tool} onChange={(e) => setTool(e.target.value as ToolName)}>
            {TOOL_DEFINITIONS.map((def) => (
              <option key={def.name} value={def.name}>
                {def.name}
              </option>
            ))}
          </select>
        </label>
        <label>
          工作表
          <input value={sheetName} onChange={(e) => setSheetName(e.target.value)} />
        </label>
        <label>
          区域
          <input value={range} onChange={(e) => setRange(e.target.value)} />
        </label>
      </div>
      {(tool === "range.write" || tool === "formula.write" || tool === "sheet.rename") && (
        <div className="row">
          {tool === "range.write" && (
            <label>
              values JSON
              <textarea
                rows={3}
                value={valuesJson}
                onChange={(e) => setValuesJson(e.target.value)}
              />
            </label>
          )}
          {tool === "formula.write" && (
            <label>
              formula
              <input value={formula} onChange={(e) => setFormula(e.target.value)} />
            </label>
          )}
          {tool === "sheet.rename" && (
            <label>
              newName
              <input value={newName} onChange={(e) => setNewName(e.target.value)} />
            </label>
          )}
        </div>
      )}
      <div className="row">
        <button type="button" onClick={() => void run()}>
          执行
        </button>
      </div>
      {output && <pre>{output}</pre>}
    </section>
  );
}
