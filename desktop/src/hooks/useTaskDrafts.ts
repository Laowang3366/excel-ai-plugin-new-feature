/**
 * useTaskDrafts — 任务编排面板草稿状态管理
 *
 * 从 ChatPage.tsx 提取，管理：
 * - 各意图类型的草稿数据（formula/code/ocr/clean/report/chart）
 * - 草稿更新回调
 * - 简易功能面板选区获取
 *
 * 草稿按会话键保存，功能侧栏关闭不会清理草稿。
 */

import { useState, useCallback, useMemo, type Dispatch, type SetStateAction } from "react";
import type { FormulaTaskDraft } from "../components/task/FormulaTaskComposerPanel";
import type { CodeTaskDraft } from "../components/task/codeTaskComposerModel";
import type { OCRTaskDraft } from "../components/task/OCRTaskComposerPanel";
import type { ReportTaskDraft } from "../components/task/ReportTaskComposerPanel";
import { pickExcelRange } from "../utils/chatHelpers";
import type { IntentKind } from "../utils/sidebarHelpers";

type SimpleTaskIntent = Extract<NonNullable<IntentKind>, "clean" | "chart">;

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
  update: SetStateAction<TaskDrafts>,
): TaskDraftStore {
  const currentDrafts = getTaskDraftsForKey(store, draftKey);
  const nextDrafts =
    typeof update === "function"
      ? (update as (prev: TaskDrafts) => TaskDrafts)(currentDrafts)
      : update;
  return { ...store, [draftKey]: nextDrafts };
}

export function moveTaskDraftStore(
  store: TaskDraftStore,
  fromKey: string,
  toKey: string,
): TaskDraftStore {
  if (fromKey === toKey || !store[fromKey] || store[toKey]) return store;
  const { [fromKey]: drafts, ...rest } = store;
  return { ...rest, [toKey]: drafts };
}

export function createEmptyFormulaDraft(current?: FormulaTaskDraft): FormulaTaskDraft {
  return {
    dataSourceRanges: [],
    dataSourceInput: "",
    referenceSampleRange: "",
    referenceSampleMode: "partial",
    outputRange: "",
    hostEnvironment: current?.hostEnvironment ?? "unknown",
    task: "",
  };
}

export function useTaskDrafts(draftKey = "default") {
  const [taskDraftStore, setTaskDraftStore] = useState<TaskDraftStore>({});
  const activeDraftKey = draftKey || "default";
  const taskDrafts = useMemo(
    () => getTaskDraftsForKey(taskDraftStore, activeDraftKey),
    [taskDraftStore, activeDraftKey],
  );

  const setTaskDrafts: Dispatch<SetStateAction<TaskDrafts>> = useCallback(
    (update) => {
      setTaskDraftStore((prev) => updateTaskDraftStore(prev, activeDraftKey, update));
    },
    [activeDraftKey],
  );

  const updateFormulaDraft = useCallback(
    (draft: FormulaTaskDraft) => {
      setTaskDrafts((prev) => ({ ...prev, formula: draft }));
    },
    [setTaskDrafts],
  );

  const resetFormulaDraft = useCallback(() => {
    setTaskDrafts((prev) => ({
      ...prev,
      formula: createEmptyFormulaDraft(prev.formula),
    }));
  }, [setTaskDrafts]);

  const updateCodeDraft = useCallback(
    (draft: CodeTaskDraft) => {
      setTaskDrafts((prev) => ({ ...prev, code: draft }));
    },
    [setTaskDrafts],
  );

  const updateOCRDraft = useCallback(
    (draft: OCRTaskDraft) => {
      setTaskDrafts((prev) => ({ ...prev, ocr: draft }));
    },
    [setTaskDrafts],
  );

  const updateReportDraft = useCallback(
    (draft: ReportTaskDraft) => {
      setTaskDrafts((prev) => ({ ...prev, report: draft }));
    },
    [setTaskDrafts],
  );

  // 简易功能面板选区按钮
  const handleSimplePickRange = useCallback(
    async (intent: SimpleTaskIntent) => {
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
    },
    [setTaskDrafts],
  );

  const moveTaskDrafts = useCallback((fromKey: string, toKey: string) => {
    setTaskDraftStore((prev) => moveTaskDraftStore(prev, fromKey, toKey));
  }, []);

  return {
    taskDrafts,
    setTaskDrafts,
    updateFormulaDraft,
    resetFormulaDraft,
    updateCodeDraft,
    updateOCRDraft,
    updateReportDraft,
    handleSimplePickRange,
    moveTaskDrafts,
  };
}
