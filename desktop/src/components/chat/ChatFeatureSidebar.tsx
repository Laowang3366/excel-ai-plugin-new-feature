import React from "react";

import type { useTaskDrafts } from "../../hooks/useTaskDrafts";
import { getAppText } from "../../i18n";
import type { AppLanguage } from "../../store/settingsStore";
import type { IntentKind } from "../../utils/sidebarHelpers";
import { FeatureSidebarPanel } from "../common/FeatureSidebarPanel";

const FormulaTaskComposerPanel = React.lazy(() =>
  import("../task/FormulaTaskComposerPanel").then((module) => ({
    default: module.FormulaTaskComposerPanel,
  })),
);
const CodeTaskComposerPanel = React.lazy(() =>
  import("../task/CodeTaskComposerPanel").then((module) => ({
    default: module.CodeTaskComposerPanel,
  })),
);
const OCRTaskComposerPanel = React.lazy(() =>
  import("../task/OCRTaskComposerPanel").then((module) => ({
    default: module.OCRTaskComposerPanel,
  })),
);
const ReportTaskComposerPanel = React.lazy(() =>
  import("../task/ReportTaskComposerPanel").then((module) => ({
    default: module.ReportTaskComposerPanel,
  })),
);
const SimpleTaskComposerPanel = React.lazy(() =>
  import("../task/SimpleTaskComposerPanel").then((module) => ({
    default: module.SimpleTaskComposerPanel,
  })),
);
const OfficeAutomationPanel = React.lazy(() =>
  import("../office/OfficeAutomationPanel").then((module) => ({
    default: module.OfficeAutomationPanel,
  })),
);

interface ChatFeatureSidebarProps {
  activeIntent: IntentKind;
  composerDraftKey: string;
  controller: ReturnType<typeof useTaskDrafts>;
  isOpen: boolean;
  language: AppLanguage;
  onClose: () => void;
  onIntentClick: (intent: NonNullable<IntentKind>) => void;
  onTaskSubmit: (payload: string) => void;
}

export function ChatFeatureSidebar({
  activeIntent,
  composerDraftKey,
  controller,
  isOpen,
  language,
  onClose,
  onIntentClick,
  onTaskSubmit,
}: ChatFeatureSidebarProps) {
  const text = getAppText(language);
  const {
    handleSimplePickRange,
    resetFormulaDraft,
    taskDrafts,
    updateCodeDraft,
    updateFormulaDraft,
    updateOCRDraft,
    updateReportDraft,
    updateSimpleRange,
    updateSimpleTask,
  } = controller;

  return (
    <FeatureSidebarPanel
      isOpen={isOpen}
      activeIntent={activeIntent}
      language={language}
      onIntentClick={onIntentClick}
      onClose={onClose}
    >
      <React.Suspense
        fallback={
          <div className="feature-sidebar-loading" role="status">
            {language === "zh-CN" ? "正在加载功能面板..." : "Loading feature panel..."}
          </div>
        }
      >
        {activeIntent === "formula" && (
          <FormulaTaskComposerPanel
            key={`${composerDraftKey}:formula`}
            embedded
            draft={taskDrafts.formula}
            onDraftChange={updateFormulaDraft}
            onSubmit={(payload) => {
              resetFormulaDraft();
              onTaskSubmit(payload);
            }}
            onClose={onClose}
          />
        )}
        {activeIntent === "code" && (
          <CodeTaskComposerPanel
            key={`${composerDraftKey}:code`}
            embedded
            draft={taskDrafts.code}
            onDraftChange={updateCodeDraft}
            onSubmit={onTaskSubmit}
            onClose={onClose}
          />
        )}
        {activeIntent === "ocr" && (
          <OCRTaskComposerPanel
            key={`${composerDraftKey}:ocr`}
            embedded
            draft={taskDrafts.ocr}
            onDraftChange={updateOCRDraft}
            onClose={onClose}
          />
        )}
        {activeIntent === "report" && (
          <ReportTaskComposerPanel
            key={`${composerDraftKey}:report`}
            embedded
            draft={taskDrafts.report}
            onDraftChange={updateReportDraft}
            onSubmit={onTaskSubmit}
            onClose={onClose}
          />
        )}
        {(activeIntent === "clean" || activeIntent === "chart") && (
          <SimpleTaskComposerPanel
            key={`${composerDraftKey}:${activeIntent}`}
            intent={activeIntent}
            range={taskDrafts[activeIntent]?.range ?? ""}
            task={taskDrafts[activeIntent]?.task ?? ""}
            text={text.chat}
            onRangeChange={(range) => updateSimpleRange(activeIntent, range)}
            onTaskChange={(task) => updateSimpleTask(activeIntent, task)}
            onPickRange={handleSimplePickRange}
            onSubmit={onTaskSubmit}
          />
        )}
        {activeIntent === "office" && <OfficeAutomationPanel />}
      </React.Suspense>
    </FeatureSidebarPanel>
  );
}
