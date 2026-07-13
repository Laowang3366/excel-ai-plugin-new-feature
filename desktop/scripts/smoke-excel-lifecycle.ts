import { mkdtemp, rm } from "node:fs/promises";
import os from "node:os";
import path from "node:path";

import { OfficeComActionBridge } from "../electron/agent/tools/implementations/office/officeComActionBridge";
import { applyExcelAdvancedAction } from "../electron/agent/tools/implementations/officeOpenXml/advancedExcel";

async function main(): Promise<void> {
  const tempDir = await mkdtemp(path.join(os.tmpdir(), "wengge-excel-lifecycle-"));
  const filePath = path.join(tempDir, "reopen.xlsx");

  try {
    const fixture = await applyExcelAdvancedAction({
      operation: "createWorkbook",
      filePath,
      params: { values: [["项目", "金额"], ["测试", 100]] },
    });
    if (fixture.status !== "done") throw new Error(fixture.error || fixture.summary);

    const bridge = new OfficeComActionBridge();
    for (let attempt = 1; attempt <= 2; attempt += 1) {
      const result = await bridge.executeAction({
        app: "excel",
        action: "inspect",
        operation: "inspectWorkbookFormatting",
        filePath,
      });
      if (result.status !== "done") {
        throw new Error(`第 ${attempt} 次打开失败: ${result.error || result.summary}`);
      }
    }

    process.stdout.write(`${JSON.stringify({ ok: true, attempts: 2, filePath }, null, 2)}\n`);
  } finally {
    await rm(tempDir, { recursive: true, force: true, maxRetries: 5, retryDelay: 200 });
  }
}

void main().catch((error) => {
  process.stderr.write(`${error instanceof Error ? error.stack || error.message : String(error)}\n`);
  process.exitCode = 1;
});
