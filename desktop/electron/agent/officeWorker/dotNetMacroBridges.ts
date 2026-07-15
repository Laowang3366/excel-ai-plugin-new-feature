import type {
  ExcelUiBridge,
  ExcelVbaBridge,
  JsaWriteOptions,
  JsaWriteResult,
  MacroLanguageCapability,
  VbaModuleWriteOptions,
  VbaModuleWriteResult,
  WpsJsaBridge,
} from "../tools/contracts/excel";
import { existsSync } from "node:fs";
import path from "node:path";
import { resolveHotPatchPath } from "../../main-modules/hotPatchManager";
import { getOfficeWorkerClient, type OfficeWorkerClient } from "./officeWorkerClient";

export class DotNetVbaBridge implements ExcelVbaBridge {
  constructor(private readonly client: OfficeWorkerClient = getOfficeWorkerClient()) {}

  detectCapabilities(): Promise<{
    supported: boolean;
    version?: string;
    host?: "excel" | "wps";
    reason?: string;
  }> {
    return this.client.invoke("excel.vba.detect");
  }

  runMacro(macroName: string, args?: unknown[]): Promise<unknown> {
    return this.client.invoke("excel.vba.run", { macroName, args: args || [] });
  }

  writeModule(
    moduleName: string,
    code: string,
    options: VbaModuleWriteOptions = {},
  ): Promise<VbaModuleWriteResult> {
    return this.client.invoke("excel.vba.writeModule", { moduleName, code, ...options });
  }
}

export class DotNetJsaBridge implements WpsJsaBridge {
  constructor(private readonly client: OfficeWorkerClient = getOfficeWorkerClient()) {}

  detectCapabilities(): Promise<MacroLanguageCapability> {
    return this.client.invoke("wps.jsa.detect");
  }

  writeCode(code: string, options: JsaWriteOptions = {}): Promise<JsaWriteResult> {
    return this.client.invoke(
      "wps.jsa.write",
      {
        code,
        ...options,
        sourceDir: resolveWpsJsaSourceDir(),
      },
      30_000,
    );
  }
}

function resolveWpsJsaSourceDir(): string {
  const candidates = [
    resolveHotPatchPath("public/wps-jsa-bridge"),
    path.join(process.cwd(), "public", "wps-jsa-bridge"),
    process.resourcesPath ? path.join(process.resourcesPath, "public", "wps-jsa-bridge") : "",
  ].filter((candidate): candidate is string => Boolean(candidate));
  const source = candidates.find((candidate) => existsSync(path.join(candidate, "index.html")));
  if (!source) throw new Error("安装包缺少 WPS JSA 内部桥接资源");
  return source;
}

export class DotNetUiBridge implements ExcelUiBridge {
  constructor(private readonly client: OfficeWorkerClient = getOfficeWorkerClient()) {}

  addControl(params: Parameters<ExcelUiBridge["addControl"]>[0]): Promise<unknown> {
    return this.client.invoke("excel.ui.addControl", params);
  }

  async removeControl(sheetName: string, name: string): Promise<void> {
    await this.client.invoke("excel.ui.removeControl", { sheetName, name });
  }

  listControls(sheetName: string): Promise<unknown[]> {
    return this.client.invoke("excel.ui.listControls", { sheetName });
  }

  createForm(params: Parameters<ExcelUiBridge["createForm"]>[0]): Promise<unknown> {
    return this.client.invoke("excel.ui.createForm", params);
  }

  addMenu(params: Parameters<ExcelUiBridge["addMenu"]>[0]): Promise<unknown> {
    return this.client.invoke("excel.ui.addMenu", params);
  }
}
