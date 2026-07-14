import { useCallback, useEffect, useMemo, useState } from "react";

import type {
  OfficeAutomationApp,
  OfficeAutomationDocument,
  OfficeAutomationObject,
  OfficeAutomationResult,
  OfficeAutomationTemplate,
  OfficeAutomationTransaction,
  OfficeAutomationWorkflow,
} from "../../electronApi";
import { ipcApi } from "../../services/ipcApi";
import {
  AlertTriangle,
  Check,
  ChevronRight,
  Files,
  History,
  LayoutTemplate,
  LocateFixed,
  OctagonX,
  Play,
  Redo2,
  RefreshCw,
  Save,
  Trash2,
  Undo2,
  Workflow,
  X,
} from "../common/IconMap";
import {
  formatOfficeTime,
  officeAppLabel,
  officeStatusLabel,
  parseTemplateVariables,
  shortOfficePath,
} from "./officeAutomationViewModel";

type AutomationTab = "documents" | "workflows" | "transactions" | "templates";
type ForceRequest = { action: "undo" | "redo"; transaction: OfficeAutomationTransaction };

const TABS = [
  { id: "documents" as const, label: "文档与对象", icon: Files },
  { id: "workflows" as const, label: "工作流", icon: Workflow },
  { id: "transactions" as const, label: "事务", icon: History },
  { id: "templates" as const, label: "模板", icon: LayoutTemplate },
];

export function OfficeAutomationPanel() {
  const [tab, setTab] = useState<AutomationTab>("documents");
  const [documents, setDocuments] = useState<OfficeAutomationDocument[]>([]);
  const [documentFilter, setDocumentFilter] = useState<OfficeAutomationApp | "all">("all");
  const [selectedDocumentKey, setSelectedDocumentKey] = useState("");
  const [objects, setObjects] = useState<OfficeAutomationObject[]>([]);
  const [objectKind, setObjectKind] = useState("all");
  const [workflows, setWorkflows] = useState<OfficeAutomationWorkflow[]>([]);
  const [selectedWorkflowId, setSelectedWorkflowId] = useState("");
  const [transactions, setTransactions] = useState<OfficeAutomationTransaction[]>([]);
  const [selectedTransactionId, setSelectedTransactionId] = useState("");
  const [templates, setTemplates] = useState<OfficeAutomationTemplate[]>([]);
  const [templateWorkflowId, setTemplateWorkflowId] = useState("");
  const [templateName, setTemplateName] = useState("");
  const [templateDescription, setTemplateDescription] = useState("");
  const [templateVariables, setTemplateVariables] = useState("{}");
  const [deleteTemplateId, setDeleteTemplateId] = useState("");
  const [forceRequest, setForceRequest] = useState<ForceRequest | null>(null);
  const [busy, setBusy] = useState("");
  const [error, setError] = useState("");

  const loadDocuments = useCallback(async () => {
    const value = unwrap(await ipcApi.office.automation.documents.list());
    setDocuments(value);
    setSelectedDocumentKey((current) => value.some((item) => documentKey(item) === current) ? current : documentKey(value[0]));
  }, []);
  const loadWorkflows = useCallback(async () => {
    const value = unwrap(await ipcApi.office.automation.workflows.list());
    setWorkflows(value);
    setSelectedWorkflowId((current) => value.some((item) => item.id === current) ? current : value[0]?.id || "");
    setTemplateWorkflowId((current) => value.some((item) => item.id === current) ? current : value[0]?.id || "");
  }, []);
  const loadTransactions = useCallback(async () => {
    const value = unwrap(await ipcApi.office.automation.transactions.list());
    setTransactions(value);
    setSelectedTransactionId((current) => value.some((item) => item.id === current) ? current : value[0]?.id || "");
  }, []);
  const loadTemplates = useCallback(async () => setTemplates(unwrap(await ipcApi.office.automation.templates.list())), []);

  useEffect(() => {
    void Promise.all([loadDocuments(), loadWorkflows(), loadTransactions(), loadTemplates()]).catch((reason) => setError(errorText(reason)));
  }, [loadDocuments, loadTemplates, loadTransactions, loadWorkflows]);

  const selectedDocument = documents.find((item) => documentKey(item) === selectedDocumentKey);
  useEffect(() => {
    if (!selectedDocument?.fullName) { setObjects([]); return; }
    let cancelled = false;
    ipcApi.office.automation.objects.list({ app: selectedDocument.app, filePath: selectedDocument.fullName, instanceId: selectedDocument.instanceId })
      .then((response) => { if (!cancelled) setObjects(unwrap(response)); })
      .catch((reason) => { if (!cancelled) setError(errorText(reason)); });
    return () => { cancelled = true; };
  }, [selectedDocument?.app, selectedDocument?.fullName, selectedDocument?.instanceId]);

  useEffect(() => {
    if (!workflows.some((item) => item.status === "running")) return;
    const timer = window.setInterval(() => { void loadWorkflows().catch(() => undefined); }, 1500);
    return () => window.clearInterval(timer);
  }, [loadWorkflows, workflows]);

  const visibleDocuments = useMemo(() => documentFilter === "all" ? documents : documents.filter((item) => item.app === documentFilter), [documentFilter, documents]);
  const objectKinds = useMemo(() => [...new Set(objects.map((item) => item.kind))].sort(), [objects]);
  const visibleObjects = objectKind === "all" ? objects : objects.filter((item) => item.kind === objectKind);
  const selectedWorkflow = workflows.find((item) => item.id === selectedWorkflowId);
  const selectedTransaction = transactions.find((item) => item.id === selectedTransactionId);

  const perform = async (key: string, operation: () => Promise<void>) => {
    setBusy(key); setError("");
    try { await operation(); }
    catch (reason) { setError(errorText(reason)); }
    finally { setBusy(""); }
  };

  const activateDocument = (document: OfficeAutomationDocument) => perform(`document:${documentKey(document)}`, async () => {
    if (!document.fullName) throw new Error("文档尚未保存，无法按完整路径激活");
    unwrap(await ipcApi.office.automation.documents.activate({ app: document.app, filePath: document.fullName, instanceId: document.instanceId }));
    await loadDocuments();
  });
  const activateObject = (item: OfficeAutomationObject) => perform(`object:${item.locator}`, async () => {
    if (!selectedDocument?.fullName) throw new Error("请先选择文档");
    unwrap(await ipcApi.office.automation.objects.activate({ app: selectedDocument.app, filePath: selectedDocument.fullName, instanceId: selectedDocument.instanceId, locator: item.locator }));
  });
  const resumeWorkflow = (id: string) => perform(`resume:${id}`, async () => { unwrap(await ipcApi.office.automation.workflows.resume(id)); await Promise.all([loadWorkflows(), loadTransactions()]); });
  const cancelWorkflow = (id: string) => perform(`cancel:${id}`, async () => { unwrap(await ipcApi.office.automation.workflows.cancel(id)); await loadWorkflows(); });
  const applyTransaction = (action: "undo" | "redo", transaction: OfficeAutomationTransaction, force = false) => perform(`${action}:${transaction.id}`, async () => {
    const response = action === "undo"
      ? await ipcApi.office.automation.transactions.undo(transaction.id, force)
      : await ipcApi.office.automation.transactions.redo(transaction.id, force);
    const updated = unwrap(response);
    if (updated.status === "conflicted" && !force) setForceRequest({ action, transaction: updated });
    else setForceRequest(null);
    await loadTransactions();
    setSelectedTransactionId(updated.id);
  });
  const saveTemplate = () => perform("template:save", async () => {
    if (!templateWorkflowId || !templateName.trim()) throw new Error("请选择工作流并填写模板名称");
    unwrap(await ipcApi.office.automation.templates.saveFromWorkflow({ workflowId: templateWorkflowId, name: templateName, description: templateDescription || undefined }));
    setTemplateName(""); setTemplateDescription(""); await loadTemplates();
  });
  const runTemplate = (template: OfficeAutomationTemplate) => perform(`template:run:${template.id}`, async () => {
    unwrap(await ipcApi.office.automation.templates.run({ templateId: template.id, variables: parseTemplateVariables(templateVariables) }));
    await Promise.all([loadWorkflows(), loadTransactions()]);
  });
  const removeTemplate = (id: string) => perform(`template:delete:${id}`, async () => {
    unwrap(await ipcApi.office.automation.templates.delete(id)); setDeleteTemplateId(""); await loadTemplates();
  });

  return (
    <div className="office-automation-panel">
      <div className="office-automation-tabs" role="tablist" aria-label="Office 自动化管理">
        {TABS.map((item) => (
          <button key={item.id} type="button" role="tab" aria-label={item.label} title={item.label} aria-selected={tab === item.id} className={tab === item.id ? "active" : ""} onClick={() => setTab(item.id)}>
            <item.icon size={15} /><span>{item.label}</span>
          </button>
        ))}
      </div>

      {error && <div className="office-automation-error" role="alert"><AlertTriangle size={14} /><span>{error}</span><button type="button" title="关闭" onClick={() => setError("")}><X size={14} /></button></div>}

      <div className="office-automation-body">
        {tab === "documents" && (
          <div className="office-automation-view">
            <Toolbar title="已打开文档" count={documents.length} onRefresh={() => void perform("refresh:documents", loadDocuments)} busy={busy === "refresh:documents"} />
            <div className="office-app-filter" role="group" aria-label="应用筛选">
              {(["all", "excel", "word", "presentation"] as const).map((app) => <button key={app} type="button" className={documentFilter === app ? "active" : ""} onClick={() => setDocumentFilter(app)}>{app === "all" ? "全部" : app === "presentation" ? "PPT" : app[0].toUpperCase() + app.slice(1)}</button>)}
            </div>
            <div className="office-document-list">
              {visibleDocuments.map((document) => (
                <div className={`office-document-row${documentKey(document) === selectedDocumentKey ? " selected" : ""}`} key={documentKey(document)}>
                  <button className="office-row-main" type="button" onClick={() => setSelectedDocumentKey(documentKey(document))}>
                    <span className={`office-app-mark ${document.app}`} title={officeAppLabel(document.app)}>{document.app === "presentation" ? "P" : document.app[0].toUpperCase()}</span>
                    <span className="office-row-copy"><strong>{document.name || shortOfficePath(document.fullName)}</strong><small>{document.host === "wps" ? "WPS" : "Microsoft Office"} · PID {document.processId || "-"}{document.saved === false ? " · 未保存" : ""}</small></span>
                  </button>
                  <button className="office-icon-button" type="button" title="激活文档" disabled={!document.fullName || busy === `document:${documentKey(document)}`} onClick={() => void activateDocument(document)}><LocateFixed size={15} /></button>
                </div>
              ))}
              {visibleDocuments.length === 0 && <EmptyState text="没有检测到已打开的 Office 文档" />}
            </div>
            <div className="office-object-header">
              <span>对象</span><span className="office-count">{objects.length}</span>
              <select aria-label="对象类型" value={objectKind} onChange={(event) => setObjectKind(event.target.value)}><option value="all">全部类型</option>{objectKinds.map((kind) => <option key={kind} value={kind}>{kind}</option>)}</select>
            </div>
            <div className="office-object-list">
              {visibleObjects.map((item) => <button key={item.locator} className="office-object-row" type="button" disabled={busy === `object:${item.locator}`} onClick={() => void activateObject(item)}><span><strong>{item.name}</strong><small>{item.kind}{item.parent ? ` · ${item.parent}` : ""}</small></span><ChevronRight size={14} /></button>)}
              {selectedDocument && visibleObjects.length === 0 && <EmptyState text="当前文档没有可选对象" />}
            </div>
          </div>
        )}

        {tab === "workflows" && (
          <div className="office-automation-view">
            <Toolbar title="工作流" count={workflows.length} onRefresh={() => void perform("refresh:workflows", loadWorkflows)} busy={busy === "refresh:workflows"} />
            <div className="office-master-detail">
              <div className="office-master-list">{workflows.map((workflow) => <button key={workflow.id} type="button" className={workflow.id === selectedWorkflowId ? "selected" : ""} onClick={() => setSelectedWorkflowId(workflow.id)}><span className={`office-status-dot ${workflow.status}`} /><span><strong>{workflow.steps[0]?.operation || "Office 工作流"}</strong><small>{workflow.completedSteps}/{workflow.steps.length} · {officeStatusLabel(workflow.status)}</small></span></button>)}{workflows.length === 0 && <EmptyState text="暂无工作流记录" />}</div>
              <div className="office-detail-pane">{selectedWorkflow ? <>
                <div className="office-detail-title"><div><strong>{selectedWorkflow.steps[0]?.operation || "Office 工作流"}</strong><small>{formatOfficeTime(selectedWorkflow.updatedAt)} · {selectedWorkflow.id.slice(0, 8)}</small></div><StatusBadge status={selectedWorkflow.status} /></div>
                <div className="office-progress"><span style={{ width: `${selectedWorkflow.steps.length ? (selectedWorkflow.completedSteps / selectedWorkflow.steps.length) * 100 : 0}%` }} /></div>
                <div className="office-detail-actions">{["paused", "failed", "cancelled"].includes(selectedWorkflow.status) && <button type="button" className="office-command primary" disabled={busy === `resume:${selectedWorkflow.id}`} onClick={() => void resumeWorkflow(selectedWorkflow.id)}><Play size={14} />继续</button>}{selectedWorkflow.status === "running" && <button type="button" className="office-command danger" disabled={busy === `cancel:${selectedWorkflow.id}`} onClick={() => void cancelWorkflow(selectedWorkflow.id)}><OctagonX size={14} />取消</button>}</div>
                {selectedWorkflow.error && <div className="office-inline-error">{selectedWorkflow.error}</div>}
                <div className="office-step-list">{selectedWorkflow.stepRecords.map((step) => <div key={step.step} className="office-step-row"><span className={`office-step-index ${step.status}`}>{step.status === "done" ? <Check size={12} /> : step.step}</span><span><strong>{selectedWorkflow.steps[step.step - 1]?.operation || `步骤 ${step.step}`}</strong><small>{officeStatusLabel(step.status)}{step.attempts && step.attempts > 1 ? ` · 尝试 ${step.attempts} 次` : ""}</small>{step.result?.error && <em>{step.result.error}</em>}</span></div>)}</div>
              </> : <EmptyState text="选择工作流查看步骤" />}</div>
            </div>
          </div>
        )}

        {tab === "transactions" && (
          <div className="office-automation-view">
            <Toolbar title="事务与恢复" count={transactions.length} onRefresh={() => void perform("refresh:transactions", loadTransactions)} busy={busy === "refresh:transactions"} />
            <div className="office-master-detail">
              <div className="office-master-list">{transactions.map((transaction) => <button key={transaction.id} type="button" className={transaction.id === selectedTransactionId ? "selected" : ""} onClick={() => setSelectedTransactionId(transaction.id)}><span className={`office-status-dot ${transaction.status}`} /><span><strong>{transaction.changes[0]?.detail || "Office 事务"}</strong><small>{officeStatusLabel(transaction.status)} · {transaction.changes.length} 项修改</small></span></button>)}{transactions.length === 0 && <EmptyState text="暂无事务记录" />}</div>
              <div className="office-detail-pane">{selectedTransaction ? <>
                <div className="office-detail-title"><div><strong>修改清单</strong><small>{formatOfficeTime(selectedTransaction.updatedAt)} · {selectedTransaction.id.slice(0, 8)}</small></div><StatusBadge status={selectedTransaction.status} /></div>
                <div className="office-detail-actions">{canUndo(selectedTransaction) && <button type="button" className="office-command" disabled={busy === `undo:${selectedTransaction.id}`} onClick={() => void applyTransaction("undo", selectedTransaction)}><Undo2 size={14} />撤销</button>}{canRedo(selectedTransaction) && <button type="button" className="office-command" disabled={busy === `redo:${selectedTransaction.id}`} onClick={() => void applyTransaction("redo", selectedTransaction)}><Redo2 size={14} />重做</button>}</div>
                {selectedTransaction.error && <div className="office-inline-error">{selectedTransaction.error}</div>}
                {forceRequest?.transaction.id === selectedTransaction.id && <div className="office-force-confirm"><AlertTriangle size={15} /><span>{forceRequest.transaction.conflicts?.map((item) => `${shortOfficePath(item.filePath)}：${item.reason}`).join("；") || "文件已在事务外修改"}</span><div><button type="button" onClick={() => setForceRequest(null)}>取消</button><button type="button" className="danger" onClick={() => void applyTransaction(forceRequest.action, forceRequest.transaction, true)}>确认覆盖</button></div></div>}
                <div className="office-change-list">{selectedTransaction.changes.map((change, index) => <div key={`${change.kind}:${index}`}><strong>{change.detail}</strong><small>{change.target || change.kind}</small></div>)}{selectedTransaction.changes.length === 0 && <EmptyState text="没有记录到文件修改" />}</div>
              </> : <EmptyState text="选择事务查看修改清单" />}</div>
            </div>
          </div>
        )}

        {tab === "templates" && (
          <div className="office-automation-view">
            <Toolbar title="工作流模板" count={templates.length} onRefresh={() => void perform("refresh:templates", loadTemplates)} busy={busy === "refresh:templates"} />
            <div className="office-template-form">
              <select aria-label="来源工作流" value={templateWorkflowId} onChange={(event) => setTemplateWorkflowId(event.target.value)}><option value="">选择已有工作流</option>{workflows.map((workflow) => <option key={workflow.id} value={workflow.id}>{workflow.steps[0]?.operation || workflow.id.slice(0, 8)} · {officeStatusLabel(workflow.status)}</option>)}</select>
              <input aria-label="模板名称" value={templateName} maxLength={120} placeholder="模板名称" onChange={(event) => setTemplateName(event.target.value)} />
              <input aria-label="模板说明" value={templateDescription} maxLength={500} placeholder="说明（可选）" onChange={(event) => setTemplateDescription(event.target.value)} />
              <button type="button" className="office-command primary" disabled={!templateWorkflowId || !templateName.trim() || busy === "template:save"} onClick={() => void saveTemplate()}><Save size={14} />保存模板</button>
            </div>
            <label className="office-variables-field"><span>运行参数（JSON）</span><textarea value={templateVariables} rows={3} spellCheck={false} onChange={(event) => setTemplateVariables(event.target.value)} /></label>
            <div className="office-template-list">{templates.map((template) => <div key={template.id} className="office-template-row"><div><strong>{template.name}</strong><small>{template.description || `${template.steps.length} 个步骤`} · {formatOfficeTime(template.updatedAt)}</small></div><div>{deleteTemplateId === template.id ? <><button type="button" className="office-command" onClick={() => setDeleteTemplateId("")}>取消</button><button type="button" className="office-command danger" onClick={() => void removeTemplate(template.id)}>确认删除</button></> : <><button type="button" className="office-icon-button" title="运行模板" disabled={busy === `template:run:${template.id}`} onClick={() => void runTemplate(template)}><Play size={15} /></button><button type="button" className="office-icon-button danger" title="删除模板" onClick={() => setDeleteTemplateId(template.id)}><Trash2 size={15} /></button></>}</div></div>)}{templates.length === 0 && <EmptyState text="暂无工作流模板" />}</div>
          </div>
        )}
      </div>
    </div>
  );
}

function Toolbar({ title, count, onRefresh, busy }: { title: string; count: number; onRefresh: () => void; busy: boolean }) {
  return <div className="office-view-toolbar"><div><strong>{title}</strong><span className="office-count">{count}</span></div><button type="button" className="office-icon-button" title="刷新" onClick={onRefresh} disabled={busy}><RefreshCw size={15} className={busy ? "spin" : ""} /></button></div>;
}

function EmptyState({ text }: { text: string }) { return <div className="office-empty-state">{text}</div>; }
function StatusBadge({ status }: { status: string }) { return <span className={`office-status-badge ${status}`}>{officeStatusLabel(status)}</span>; }
function documentKey(document?: OfficeAutomationDocument): string { return document ? `${document.instanceId}|${document.fullName || document.name}|${document.index}` : ""; }
function unwrap<T>(response: OfficeAutomationResult<T>): T { if (!response.success || response.data === undefined) throw new Error(response.error || "Office 自动化操作失败"); return response.data; }
function errorText(reason: unknown): string { return reason instanceof Error ? reason.message : String(reason); }
function canUndo(transaction: OfficeAutomationTransaction): boolean { return transaction.status === "applied" || (transaction.status === "conflicted" && transaction.conflictBaseStatus !== "undone"); }
function canRedo(transaction: OfficeAutomationTransaction): boolean { return transaction.status === "undone" || (transaction.status === "conflicted" && transaction.conflictBaseStatus === "undone"); }
