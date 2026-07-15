import type { OfficeAutomationWorkflow } from "../../electronApi";
import { Check, OctagonX, Play } from "../common/IconMap";
import { formatOfficeTime, officeStatusLabel } from "./officeAutomationViewModel";
import { EmptyState, StatusBadge, Toolbar } from "./OfficeAutomationPanelShared";

export type OfficeAutomationWorkflowsTabProps = {
  state: {
    workflows: OfficeAutomationWorkflow[];
    selectedWorkflowId: string;
    busy: string;
  };
  actions: {
    setSelectedWorkflowId: (id: string) => void;
    refresh: () => void;
    resumeWorkflow: (id: string) => void;
    cancelWorkflow: (id: string) => void;
  };
};

export function OfficeAutomationWorkflowsTab({
  state,
  actions,
}: OfficeAutomationWorkflowsTabProps) {
  const selectedWorkflow = state.workflows.find((item) => item.id === state.selectedWorkflowId);

  return (
    <div className="office-automation-view">
      <Toolbar
        title="工作流"
        count={state.workflows.length}
        onRefresh={actions.refresh}
        busy={state.busy === "refresh:workflows"}
      />
      <div className="office-master-detail">
        <div className="office-master-list">
          {state.workflows.map((workflow) => (
            <button
              key={workflow.id}
              type="button"
              className={workflow.id === state.selectedWorkflowId ? "selected" : ""}
              onClick={() => actions.setSelectedWorkflowId(workflow.id)}
            >
              <span className={`office-status-dot ${workflow.status}`} />
              <span>
                <strong>{workflow.steps[0]?.operation || "Office 工作流"}</strong>
                <small>
                  {workflow.completedSteps}/{workflow.steps.length} ·{" "}
                  {officeStatusLabel(workflow.status)}
                </small>
              </span>
            </button>
          ))}
          {state.workflows.length === 0 && <EmptyState text="暂无工作流记录" />}
        </div>
        <div className="office-detail-pane">
          {selectedWorkflow ? (
            <>
              <div className="office-detail-title">
                <div>
                  <strong>{selectedWorkflow.steps[0]?.operation || "Office 工作流"}</strong>
                  <small>
                    {formatOfficeTime(selectedWorkflow.updatedAt)} ·{" "}
                    {selectedWorkflow.id.slice(0, 8)}
                  </small>
                </div>
                <StatusBadge status={selectedWorkflow.status} />
              </div>
              <div className="office-progress">
                <span
                  style={{
                    width: `${selectedWorkflow.steps.length ? (selectedWorkflow.completedSteps / selectedWorkflow.steps.length) * 100 : 0}%`,
                  }}
                />
              </div>
              <div className="office-detail-actions">
                {["paused", "failed", "cancelled"].includes(selectedWorkflow.status) && (
                  <button
                    type="button"
                    className="office-command primary"
                    disabled={state.busy === `resume:${selectedWorkflow.id}`}
                    onClick={() => void actions.resumeWorkflow(selectedWorkflow.id)}
                  >
                    <Play size={14} />
                    继续
                  </button>
                )}
                {selectedWorkflow.status === "running" && (
                  <button
                    type="button"
                    className="office-command danger"
                    disabled={state.busy === `cancel:${selectedWorkflow.id}`}
                    onClick={() => void actions.cancelWorkflow(selectedWorkflow.id)}
                  >
                    <OctagonX size={14} />
                    取消
                  </button>
                )}
              </div>
              {selectedWorkflow.error && (
                <div className="office-inline-error">{selectedWorkflow.error}</div>
              )}
              <div className="office-step-list">
                {selectedWorkflow.stepRecords.map((step) => (
                  <div key={step.step} className="office-step-row">
                    <span className={`office-step-index ${step.status}`}>
                      {step.status === "done" ? <Check size={12} /> : step.step}
                    </span>
                    <span>
                      <strong>
                        {selectedWorkflow.steps[step.step - 1]?.operation || `步骤 ${step.step}`}
                      </strong>
                      <small>
                        {officeStatusLabel(step.status)}
                        {step.attempts && step.attempts > 1 ? ` · 尝试 ${step.attempts} 次` : ""}
                      </small>
                      {step.result?.error && <em>{step.result.error}</em>}
                    </span>
                  </div>
                ))}
              </div>
            </>
          ) : (
            <EmptyState text="选择工作流查看步骤" />
          )}
        </div>
      </div>
    </div>
  );
}
