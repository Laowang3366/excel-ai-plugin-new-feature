/**
 * useTaskDrafts — 任务编排面板草稿状态管理
 *
 * 从 ChatPage.tsx 提取，管理：
 * - 各意图类型的草稿数据（formula/code/ocr/clean/report/chart）
 * - 草稿更新回调
 * - 简易功能面板选区获取
 * - 面板关闭时清理草稿
 */

import { useState, useCallback, useMemo, type Dispatch, type SetStateAction } from "react";
import type { IntentKind } from "../components/Sidebar";
import type { FormulaTaskDraft } from "../components/task/FormulaTaskComposerPanel";
import type { CodeTaskDraft } from "../components/task/CodeTaskComposerPanel";
import type { OCRTaskDraft } from "../components/task/OCRTaskComposerPanel";
import type { ReportTaskDraft } from "../components/task/ReportTaskComposerPanel";
import { pickExcelRange } from "../utils/chatHelpers";

type ActiveIntentKind = NonNullable<IntentKind>;
type SimpleTaskIntent = Extract<ActiveIntentKind, "clean" | "chart">;

interface SimpleTaskDraft {
  range: string;
  task: string;
}

export interface TaskDrafts {
  formula?: FormulaTaskDraft;
  code?: CodeTaskDraft;
  ocr?: OCRTaskDraft;
  report?: ReportTaskDraft;
  clean?: SimpleTaskDraft;
  chart?: SimpleTaskDraft;
}

export type TaskDraftStore = Record<string, TaskDrafts>;

export function getTaskDraftsForKey(store: TaskDraftStore, draftKey: string): TaskDrafts {
  return store[draftKey] ?? {};
}

export function updateTaskDraftStore(
  store: TaskDraftStore,
  draftKey: string,
  update: SetStateAction<TaskDrafts>
): TaskDraftStore {
  const currentDrafts = getTaskDraftsForKey(store, draftKey);
  const nextDrafts = typeof update === "function"
    ? (update as (prev: TaskDrafts) => TaskDrafts)(currentDrafts)
    : update;
  return { ...store, [draftKey]: nextDrafts };
}

export function useTaskDrafts(
  activeIntent: IntentKind,
  onIntentClick: (intent: IntentKind) => void,
  draftKey = "default"
) {
  const [taskDraftStore, setTaskDraftStore] = useState<TaskDraftStore>({});
  const activeDraftKey = draftKey || "default";
  const taskDrafts = useMemo(
    () => getTaskDraftsForKey(taskDraftStore, activeDraftKey),
    [taskDraftStore, activeDraftKey]
  );

  const setTaskDrafts: Dispatch<SetStateAction<TaskDrafts>> = useCallback((update) => {
    setTaskDraftStore((prev) => updateTaskDraftStore(prev, activeDraftKey, update));
  }, [activeDraftKey]);

  const closeActiveTaskPanel = useCallback(() => {
    if (activeIntent) {
      setTaskDrafts((prev) => ({ ...prev, [activeIntent]: undefined }));
    }
    onIntentClick(null);
  }, [activeIntent, onIntentClick, setTaskDrafts]);

  const updateFormulaDraft = useCallback((draft: FormulaTaskDraft) => {
    setTaskDrafts((prev) => ({ ...prev, formula: draft }));
  }, [setTaskDrafts]);

  const updateCodeDraft = useCallback((draft: CodeTaskDraft) => {
    setTaskDrafts((prev) => ({ ...prev, code: draft }));
  }, [setTaskDrafts]);

  const updateOCRDraft = useCallback((draft: OCRTaskDraft) => {
    setTaskDrafts((prev) => ({ ...prev, ocr: draft }));
  }, [setTaskDrafts]);

  const updateReportDraft = useCallback((draft: ReportTaskDraft) => {
    setTaskDrafts((prev) => ({ ...prev, report: draft }));
  }, [setTaskDrafts]);

  // 简易功能面板选区按钮
  const handleSimplePickRange = useCallback(async (intent: SimpleTaskIntent) => {
    const addr = await pickExcelRange();
    if (addr) {
      setTaskDrafts((prev) => ({
        ...prev,
        [intent]: {
          range: addr,
          task: prev[intent]?.task ?? "",
        },
      }));
    }
  }, [setTaskDrafts]);

  return {
    taskDrafts,
    setTaskDrafts,
    closeActiveTaskPanel,
    updateFormulaDraft,
    updateCodeDraft,
    updateOCRDraft,
    updateReportDraft,
    handleSimplePickRange,
  };
}
