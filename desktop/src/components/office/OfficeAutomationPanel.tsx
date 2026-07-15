import { useCallback, useEffect, useState } from "react";

import type {
  OfficeAutomationApp,
  OfficeAutomationDocument,
  OfficeAutomationObject,
  OfficeAutomationTemplate,
  OfficeAutomationTransaction,
  OfficeAutomationWorkflow,
} from "../../electronApi";
import { ipcApi } from "../../services/ipcApi";
import { parseTemplateVariables } from "./officeAutomationViewModel";
import { OfficeAutomationDocumentsTab } from "./OfficeAutomationDocumentsTab";
import {
  AutomationErrorBanner,
  AutomationTabList,
  type AutomationTab,
  documentKey,
  errorText,
  unwrap,
} from "./OfficeAutomationPanelShared";
import { OfficeAutomationTemplatesTab } from "./OfficeAutomationTemplatesTab";
import { OfficeAutomationTransactionsTab } from "./OfficeAutomationTransactionsTab";
import { OfficeAutomationWorkflowsTab } from "./OfficeAutomationWorkflowsTab";

type ForceRequest = { action: "undo" | "redo"; transaction: OfficeAutomationTransaction };

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
    setSelectedDocumentKey((current) =>
      value.some((item) => documentKey(item) === current) ? current : documentKey(value[0]),
    );
  }, []);
  const loadWorkflows = useCallback(async () => {
    const value = unwrap(await ipcApi.office.automation.workflows.list());
    setWorkflows(value);
    setSelectedWorkflowId((current) =>
      value.some((item) => item.id === current) ? current : value[0]?.id || "",
    );
    setTemplateWorkflowId((current) =>
      value.some((item) => item.id === current) ? current : value[0]?.id || "",
    );
  }, []);
  const loadTransactions = useCallback(async () => {
    const value = unwrap(await ipcApi.office.automation.transactions.list());
    setTransactions(value);
    setSelectedTransactionId((current) =>
      value.some((item) => item.id === current) ? current : value[0]?.id || "",
    );
  }, []);
  const loadTemplates = useCallback(
    async () => setTemplates(unwrap(await ipcApi.office.automation.templates.list())),
    [],
  );

  useEffect(() => {
    void Promise.all([loadDocuments(), loadWorkflows(), loadTransactions(), loadTemplates()]).catch(
      (reason) => setError(errorText(reason)),
    );
  }, [loadDocuments, loadTemplates, loadTransactions, loadWorkflows]);

  const selectedDocument = documents.find((item) => documentKey(item) === selectedDocumentKey);
  useEffect(() => {
    if (!selectedDocument?.fullName) {
      setObjects([]);
      return;
    }
    let cancelled = false;
    ipcApi.office.automation.objects
      .list({
        app: selectedDocument.app,
        filePath: selectedDocument.fullName,
        instanceId: selectedDocument.instanceId,
      })
      .then((response) => {
        if (!cancelled) setObjects(unwrap(response));
      })
      .catch((reason) => {
        if (!cancelled) setError(errorText(reason));
      });
    return () => {
      cancelled = true;
    };
  }, [selectedDocument?.app, selectedDocument?.fullName, selectedDocument?.instanceId]);

  useEffect(() => {
    if (!workflows.some((item) => item.status === "running")) return;
    const timer = window.setInterval(() => void loadWorkflows().catch(() => undefined), 1500);
    return () => window.clearInterval(timer);
  }, [loadWorkflows, workflows]);

  const perform = async (key: string, operation: () => Promise<void>) => {
    setBusy(key);
    setError("");
    try {
      await operation();
    } catch (reason) {
      setError(errorText(reason));
    } finally {
      setBusy("");
    }
  };

  const activateDocument = (document: OfficeAutomationDocument) =>
    perform(`document:${documentKey(document)}`, async () => {
      if (!document.fullName) throw new Error("文档尚未保存，无法按完整路径激活");
      unwrap(
        await ipcApi.office.automation.documents.activate({
          app: document.app,
          filePath: document.fullName,
          instanceId: document.instanceId,
        }),
      );
      await loadDocuments();
    });
  const activateObject = (item: OfficeAutomationObject) =>
    perform(`object:${item.locator}`, async () => {
      if (!selectedDocument?.fullName) throw new Error("请先选择文档");
      unwrap(
        await ipcApi.office.automation.objects.activate({
          app: selectedDocument.app,
          filePath: selectedDocument.fullName,
          instanceId: selectedDocument.instanceId,
          locator: item.locator,
        }),
      );
    });
  const resumeWorkflow = (id: string) =>
    perform(`resume:${id}`, async () => {
      unwrap(await ipcApi.office.automation.workflows.resume(id));
      await Promise.all([loadWorkflows(), loadTransactions()]);
    });
  const cancelWorkflow = (id: string) =>
    perform(`cancel:${id}`, async () => {
      unwrap(await ipcApi.office.automation.workflows.cancel(id));
      await loadWorkflows();
    });
  const applyTransaction = (
    action: "undo" | "redo",
    transaction: OfficeAutomationTransaction,
    force = false,
  ) =>
    perform(`${action}:${transaction.id}`, async () => {
      const tx = ipcApi.office.automation.transactions;
      const updated = unwrap(
        action === "undo"
          ? await tx.undo(transaction.id, force)
          : await tx.redo(transaction.id, force),
      );
      if (updated.status === "conflicted" && !force)
        setForceRequest({ action, transaction: updated });
      else setForceRequest(null);
      await loadTransactions();
      setSelectedTransactionId(updated.id);
    });
  const saveTemplate = () =>
    perform("template:save", async () => {
      if (!templateWorkflowId || !templateName.trim())
        throw new Error("请选择工作流并填写模板名称");
      unwrap(
        await ipcApi.office.automation.templates.saveFromWorkflow({
          workflowId: templateWorkflowId,
          name: templateName,
          description: templateDescription || undefined,
        }),
      );
      setTemplateName("");
      setTemplateDescription("");
      await loadTemplates();
    });
  const runTemplate = (template: OfficeAutomationTemplate) =>
    perform(`template:run:${template.id}`, async () => {
      unwrap(
        await ipcApi.office.automation.templates.run({
          templateId: template.id,
          variables: parseTemplateVariables(templateVariables),
        }),
      );
      await Promise.all([loadWorkflows(), loadTransactions()]);
    });
  const removeTemplate = (id: string) =>
    perform(`template:delete:${id}`, async () => {
      unwrap(await ipcApi.office.automation.templates.delete(id));
      setDeleteTemplateId("");
      await loadTemplates();
    });

  return (
    <div className="office-automation-panel">
      <AutomationTabList tab={tab} onChange={setTab} />
      <AutomationErrorBanner error={error} onClose={() => setError("")} />
      <div className="office-automation-body">
        {tab === "documents" && (
          <OfficeAutomationDocumentsTab
            state={{ documents, documentFilter, selectedDocumentKey, objects, objectKind, busy }}
            actions={{
              setDocumentFilter,
              setSelectedDocumentKey,
              setObjectKind,
              refresh: () => void perform("refresh:documents", loadDocuments),
              activateDocument,
              activateObject,
            }}
          />
        )}
        {tab === "workflows" && (
          <OfficeAutomationWorkflowsTab
            state={{ workflows, selectedWorkflowId, busy }}
            actions={{
              setSelectedWorkflowId,
              refresh: () => void perform("refresh:workflows", loadWorkflows),
              resumeWorkflow,
              cancelWorkflow,
            }}
          />
        )}
        {tab === "transactions" && (
          <OfficeAutomationTransactionsTab
            state={{ transactions, selectedTransactionId, forceRequest, busy }}
            actions={{
              setSelectedTransactionId,
              setForceRequest,
              refresh: () => void perform("refresh:transactions", loadTransactions),
              applyTransaction,
            }}
          />
        )}
        {tab === "templates" && (
          <OfficeAutomationTemplatesTab
            state={{
              templates,
              workflows,
              templateWorkflowId,
              templateName,
              templateDescription,
              templateVariables,
              deleteTemplateId,
              busy,
            }}
            actions={{
              setTemplateWorkflowId,
              setTemplateName,
              setTemplateDescription,
              setTemplateVariables,
              setDeleteTemplateId,
              refresh: () => void perform("refresh:templates", loadTemplates),
              saveTemplate,
              runTemplate,
              removeTemplate,
            }}
          />
        )}
      </div>
    </div>
  );
}
