import type { ToolExecutor } from "../../shared/types";
import type { WordDocumentBridge, OfficeActionBridge } from "../contracts/office";
import { validateArgs } from "./validation";

export interface OfficeWordExecutorDeps {
  wordBridge: WordDocumentBridge;
  officeActionBridge?: OfficeActionBridge;
  inspectFileWithOpenXml: (
    officeActionBridge: OfficeActionBridge,
    app: "word" | "presentation",
    filePath: string,
    openError?: string,
  ) => Promise<{ success: boolean; error: string; data: Record<string, unknown> } | undefined>;
}

export function addOfficeWordExecutors(
  target: Map<string, ToolExecutor>,
  deps: OfficeWordExecutorDeps,
): void {
  const { wordBridge, officeActionBridge, inspectFileWithOpenXml } = deps;

  target.set("word.open", {
    name: "word.open",
    execute: async (args: Record<string, unknown>) => {
      const err = validateArgs(args, { filePath: "string" });
      if (err) return { success: false, error: err };
      const filePath = args.filePath as string;
      const result = await wordBridge.openDocument(filePath);
      if (!result.success && officeActionBridge) {
        const fallback = await inspectFileWithOpenXml(
          officeActionBridge,
          "word",
          filePath,
          result.error,
        );
        if (fallback) return fallback;
      }
      return { success: result.success, data: result, error: result.error };
    },
  });

  target.set("word.inspect", {
    name: "word.inspect",
    execute: async (_args: Record<string, unknown>) => {
      const result = await wordBridge.inspectDocument();
      return { success: true, data: result };
    },
  });

  target.set("word.readText", {
    name: "word.readText",
    execute: async (args: Record<string, unknown>) => {
      const maxChars = typeof args.maxChars === "number" ? args.maxChars : undefined;
      const result = await wordBridge.readText(maxChars);
      return { success: true, data: result };
    },
  });

  target.set("word.insertText", {
    name: "word.insertText",
    execute: async (args: Record<string, unknown>) => {
      const err = validateArgs(args, { text: "string" });
      if (err) return { success: false, error: err };
      const result = await wordBridge.insertText(
        args.text as string,
        args.position as string | undefined,
      );
      return { success: true, data: result };
    },
  });

  target.set("word.insertHeading", {
    name: "word.insertHeading",
    execute: async (args: Record<string, unknown>) => {
      const err = validateArgs(args, { text: "string" });
      if (err) return { success: false, error: err };
      const result = await wordBridge.insertHeading(
        args.text as string,
        typeof args.level === "number" ? args.level : undefined,
        args.position as string | undefined,
      );
      return { success: true, data: result };
    },
  });

  target.set("word.replaceText", {
    name: "word.replaceText",
    execute: async (args: Record<string, unknown>) => {
      const err = validateArgs(args, { findText: "string", replaceText: "string" });
      if (err) return { success: false, error: err };
      const result = await wordBridge.replaceText(
        args.findText as string,
        args.replaceText as string,
        typeof args.matchCase === "boolean" ? args.matchCase : undefined,
      );
      return { success: true, data: result };
    },
  });

  target.set("word.save", {
    name: "word.save",
    execute: async (args: Record<string, unknown>) => {
      const result = await wordBridge.saveDocument(args.saveAsPath as string | undefined);
      return { success: result.success, data: result, error: result.error };
    },
  });
}
