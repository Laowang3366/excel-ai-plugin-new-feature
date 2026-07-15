import type { OfficeAutomationTemplate, OfficeAutomationWorkflow } from "../../electronApi";
import { Play, Save, Trash2 } from "../common/IconMap";
import { formatOfficeTime, officeStatusLabel } from "./officeAutomationViewModel";
import { EmptyState, Toolbar } from "./OfficeAutomationPanelShared";

export type OfficeAutomationTemplatesTabProps = {
  state: {
    templates: OfficeAutomationTemplate[];
    workflows: OfficeAutomationWorkflow[];
    templateWorkflowId: string;
    templateName: string;
    templateDescription: string;
    templateVariables: string;
    deleteTemplateId: string;
    busy: string;
  };
  actions: {
    setTemplateWorkflowId: (id: string) => void;
    setTemplateName: (name: string) => void;
    setTemplateDescription: (description: string) => void;
    setTemplateVariables: (variables: string) => void;
    setDeleteTemplateId: (id: string) => void;
    refresh: () => void;
    saveTemplate: () => void;
    runTemplate: (template: OfficeAutomationTemplate) => void;
    removeTemplate: (id: string) => void;
  };
};

export function OfficeAutomationTemplatesTab({
  state,
  actions,
}: OfficeAutomationTemplatesTabProps) {
  return (
    <div className="office-automation-view">
      <Toolbar
        title="工作流模板"
        count={state.templates.length}
        onRefresh={actions.refresh}
        busy={state.busy === "refresh:templates"}
      />
      <div className="office-template-form">
        <select
          aria-label="来源工作流"
          value={state.templateWorkflowId}
          onChange={(event) => actions.setTemplateWorkflowId(event.target.value)}
        >
          <option value="">选择已有工作流</option>
          {state.workflows.map((workflow) => (
            <option key={workflow.id} value={workflow.id}>
              {workflow.steps[0]?.operation || workflow.id.slice(0, 8)} ·{" "}
              {officeStatusLabel(workflow.status)}
            </option>
          ))}
        </select>
        <input
          aria-label="模板名称"
          value={state.templateName}
          maxLength={120}
          placeholder="模板名称"
          onChange={(event) => actions.setTemplateName(event.target.value)}
        />
        <input
          aria-label="模板说明"
          value={state.templateDescription}
          maxLength={500}
          placeholder="说明（可选）"
          onChange={(event) => actions.setTemplateDescription(event.target.value)}
        />
        <button
          type="button"
          className="office-command primary"
          disabled={
            !state.templateWorkflowId ||
            !state.templateName.trim() ||
            state.busy === "template:save"
          }
          onClick={() => void actions.saveTemplate()}
        >
          <Save size={14} />
          保存模板
        </button>
      </div>
      <label className="office-variables-field">
        <span>运行参数（JSON）</span>
        <textarea
          value={state.templateVariables}
          rows={3}
          spellCheck={false}
          onChange={(event) => actions.setTemplateVariables(event.target.value)}
        />
      </label>
      <div className="office-template-list">
        {state.templates.map((template) => (
          <div key={template.id} className="office-template-row">
            <div>
              <strong>{template.name}</strong>
              <small>
                {template.description || `${template.steps.length} 个步骤`} ·{" "}
                {formatOfficeTime(template.updatedAt)}
              </small>
            </div>
            <div>
              {state.deleteTemplateId === template.id ? (
                <>
                  <button
                    type="button"
                    className="office-command"
                    onClick={() => actions.setDeleteTemplateId("")}
                  >
                    取消
                  </button>
                  <button
                    type="button"
                    className="office-command danger"
                    onClick={() => void actions.removeTemplate(template.id)}
                  >
                    确认删除
                  </button>
                </>
              ) : (
                <>
                  <button
                    type="button"
                    className="office-icon-button"
                    title="运行模板"
                    disabled={state.busy === `template:run:${template.id}`}
                    onClick={() => void actions.runTemplate(template)}
                  >
                    <Play size={15} />
                  </button>
                  <button
                    type="button"
                    className="office-icon-button danger"
                    title="删除模板"
                    onClick={() => actions.setDeleteTemplateId(template.id)}
                  >
                    <Trash2 size={15} />
                  </button>
                </>
              )}
            </div>
          </div>
        ))}
        {state.templates.length === 0 && <EmptyState text="暂无工作流模板" />}
      </div>
    </div>
  );
}
