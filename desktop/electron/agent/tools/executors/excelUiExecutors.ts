/**
 * Excel UI 控件工具执行器
 *
 * 只注册工作表控件、表单和菜单相关工具。
 */

import type { ToolExecutor } from "../../shared/types";
import type { ExcelUiBridge } from "../contracts/excel";
import { validateArgs } from "./validation";

export interface ExcelUiExecutorDeps {
  uiBridge: ExcelUiBridge;
}

export function addExcelUiExecutors(
  target: Map<string, ToolExecutor>,
  deps: ExcelUiExecutorDeps,
): void {
  const { uiBridge } = deps;

  target.set("ui.addControl", {
    name: "ui.addControl",
    execute: async (args: Record<string, unknown>) => {
      const err = validateArgs(args, {
        sheetName: "string",
        controlType: "string",
        name: "string",
        left: "number",
        top: "number",
        width: "number",
        height: "number",
      });
      if (err) return { success: false, error: err };
      const result = await uiBridge.addControl({
        sheetName: args.sheetName as string,
        controlType: args.controlType as string,
        name: args.name as string,
        left: args.left as number,
        top: args.top as number,
        width: args.width as number,
        height: args.height as number,
        caption: args.caption as string | undefined,
        macroName: args.macroName as string | undefined,
        linkedCell: args.linkedCell as string | undefined,
      });
      return { success: true, data: result };
    },
  });

  target.set("ui.removeControl", {
    name: "ui.removeControl",
    execute: async (args: Record<string, unknown>) => {
      const err = validateArgs(args, { sheetName: "string", name: "string" });
      if (err) return { success: false, error: err };
      await uiBridge.removeControl(args.sheetName as string, args.name as string);
      return { success: true, data: "控件已删除" };
    },
  });

  target.set("ui.listControls", {
    name: "ui.listControls",
    execute: async (args: Record<string, unknown>) => {
      const err = validateArgs(args, { sheetName: "string" });
      if (err) return { success: false, error: err };
      const controls = await uiBridge.listControls(args.sheetName as string);
      return { success: true, data: controls };
    },
  });

  target.set("ui.createForm", {
    name: "ui.createForm",
    execute: async (args: Record<string, unknown>) => {
      const err = validateArgs(args, { formName: "string", caption: "string" });
      if (err) return { success: false, error: err };
      const result = await uiBridge.createForm({
        formName: args.formName as string,
        caption: args.caption as string,
        width: args.width as number | undefined,
        height: args.height as number | undefined,
        controls: args.controls as Array<Record<string, unknown>> | undefined,
        eventCode: args.eventCode as string | undefined,
      });
      return { success: true, data: result };
    },
  });

  target.set("ui.addMenu", {
    name: "ui.addMenu",
    execute: async (args: Record<string, unknown>) => {
      const err = validateArgs(args, {
        menuBar: "string",
        caption: "string",
        macroName: "string",
      });
      if (err) return { success: false, error: err };
      const result = await uiBridge.addMenu({
        menuBar: args.menuBar as string,
        caption: args.caption as string,
        macroName: args.macroName as string,
        beforeId: args.beforeId as number | undefined,
        faceId: args.faceId as number | undefined,
      });
      return { success: true, data: result };
    },
  });
}
