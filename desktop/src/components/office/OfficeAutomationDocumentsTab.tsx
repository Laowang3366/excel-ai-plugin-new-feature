import { useMemo } from "react";

import type {
  OfficeAutomationApp,
  OfficeAutomationDocument,
  OfficeAutomationObject,
} from "../../electronApi";
import { ChevronRight, LocateFixed } from "../common/IconMap";
import { officeAppLabel, shortOfficePath } from "./officeAutomationViewModel";
import { documentKey, EmptyState, Toolbar } from "./OfficeAutomationPanelShared";

export type OfficeAutomationDocumentsTabProps = {
  state: {
    documents: OfficeAutomationDocument[];
    documentFilter: OfficeAutomationApp | "all";
    selectedDocumentKey: string;
    objects: OfficeAutomationObject[];
    objectKind: string;
    busy: string;
  };
  actions: {
    setDocumentFilter: (app: OfficeAutomationApp | "all") => void;
    setSelectedDocumentKey: (key: string) => void;
    setObjectKind: (kind: string) => void;
    refresh: () => void;
    activateDocument: (document: OfficeAutomationDocument) => void;
    activateObject: (item: OfficeAutomationObject) => void;
  };
};

export function OfficeAutomationDocumentsTab({
  state,
  actions,
}: OfficeAutomationDocumentsTabProps) {
  const selectedDocument = state.documents.find(
    (item) => documentKey(item) === state.selectedDocumentKey,
  );
  const visibleDocuments = useMemo(
    () =>
      state.documentFilter === "all"
        ? state.documents
        : state.documents.filter((item) => item.app === state.documentFilter),
    [state.documentFilter, state.documents],
  );
  const objectKinds = useMemo(
    () => [...new Set(state.objects.map((item) => item.kind))].sort(),
    [state.objects],
  );
  const visibleObjects =
    state.objectKind === "all"
      ? state.objects
      : state.objects.filter((item) => item.kind === state.objectKind);

  return (
    <div className="office-automation-view">
      <Toolbar
        title="已打开文档"
        count={state.documents.length}
        onRefresh={actions.refresh}
        busy={state.busy === "refresh:documents"}
      />
      <div className="office-app-filter" role="group" aria-label="应用筛选">
        {(["all", "excel", "word", "presentation"] as const).map((app) => (
          <button
            key={app}
            type="button"
            className={state.documentFilter === app ? "active" : ""}
            onClick={() => actions.setDocumentFilter(app)}
          >
            {app === "all"
              ? "全部"
              : app === "presentation"
                ? "PPT"
                : app[0].toUpperCase() + app.slice(1)}
          </button>
        ))}
      </div>
      <div className="office-document-list">
        {visibleDocuments.map((document) => (
          <div
            className={`office-document-row${documentKey(document) === state.selectedDocumentKey ? " selected" : ""}`}
            key={documentKey(document)}
          >
            <button
              className="office-row-main"
              type="button"
              onClick={() => actions.setSelectedDocumentKey(documentKey(document))}
            >
              <span
                className={`office-app-mark ${document.app}`}
                title={officeAppLabel(document.app)}
              >
                {document.app === "presentation" ? "P" : document.app[0].toUpperCase()}
              </span>
              <span className="office-row-copy">
                <strong>{document.name || shortOfficePath(document.fullName)}</strong>
                <small>
                  {document.host === "wps" ? "WPS" : "Microsoft Office"} · PID{" "}
                  {document.processId || "-"}
                  {document.saved === false ? " · 未保存" : ""}
                </small>
              </span>
            </button>
            <button
              className="office-icon-button"
              type="button"
              title="激活文档"
              disabled={!document.fullName || state.busy === `document:${documentKey(document)}`}
              onClick={() => void actions.activateDocument(document)}
            >
              <LocateFixed size={15} />
            </button>
          </div>
        ))}
        {visibleDocuments.length === 0 && <EmptyState text="没有检测到已打开的 Office 文档" />}
      </div>
      <div className="office-object-header">
        <span>对象</span>
        <span className="office-count">{state.objects.length}</span>
        <select
          aria-label="对象类型"
          value={state.objectKind}
          onChange={(event) => actions.setObjectKind(event.target.value)}
        >
          <option value="all">全部类型</option>
          {objectKinds.map((kind) => (
            <option key={kind} value={kind}>
              {kind}
            </option>
          ))}
        </select>
      </div>
      <div className="office-object-list">
        {visibleObjects.map((item) => (
          <button
            key={item.locator}
            className="office-object-row"
            type="button"
            disabled={state.busy === `object:${item.locator}`}
            onClick={() => void actions.activateObject(item)}
          >
            <span>
              <strong>{item.name}</strong>
              <small>
                {item.kind}
                {item.parent ? ` · ${item.parent}` : ""}
              </small>
            </span>
            <ChevronRight size={14} />
          </button>
        ))}
        {selectedDocument && visibleObjects.length === 0 && (
          <EmptyState text="当前文档没有可选对象" />
        )}
      </div>
    </div>
  );
}
