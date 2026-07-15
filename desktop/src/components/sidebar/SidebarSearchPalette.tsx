import React, { useEffect, useMemo, useRef, useState } from "react";

import type { FolderFileInfo, ThreadMetadata } from "../../electronApi";
import { getAppText } from "../../i18n";
import type { AppLanguage, PinnedFolder } from "../../store/settingsStore";
import { useDocumentDismiss } from "../../hooks/useDocumentDismiss";
import { buildSidebarSearchResults } from "../../utils/sidebarSearch";
import { formatTime } from "../../utils/sidebarHelpers";
import {
  FileText,
  FolderOpen,
  MessageSquare,
  PenLine,
  Search,
  Settings,
  X,
} from "../common/IconMap";

type SearchTab = "all" | "threads" | "files" | "actions";

interface SidebarSearchPaletteProps {
  open: boolean;
  threads: ThreadMetadata[];
  folders: PinnedFolder[];
  folderFiles: Record<string, FolderFileInfo[]>;
  language: AppLanguage;
  activeThreadId: string | null;
  onClose: () => void;
  onSwitchThread: (threadId: string) => void;
  onAddFile: (file: FolderFileInfo) => void;
  onCreateNewThread: () => void;
  onAddFolder: () => void;
  onOpenSettings: () => void;
}

export const SidebarSearchPalette: React.FC<SidebarSearchPaletteProps> = ({
  open,
  threads,
  folders,
  folderFiles,
  language,
  activeThreadId,
  onClose,
  onSwitchThread,
  onAddFile,
  onCreateNewThread,
  onAddFolder,
  onOpenSettings,
}) => {
  const text = getAppText(language);
  const [query, setQuery] = useState("");
  const [tab, setTab] = useState<SearchTab>("all");
  const cardRef = useRef<HTMLDivElement>(null);
  const inputRef = useRef<HTMLInputElement>(null);
  const dismissBoundaryRefs = useMemo(() => [cardRef], []);

  useEffect(() => {
    if (!open) return;
    setQuery("");
    setTab("all");
    requestAnimationFrame(() => inputRef.current?.focus());
  }, [open]);

  useDocumentDismiss({
    active: open,
    boundaryRefs: dismissBoundaryRefs,
    onDismiss: onClose,
    pointerEvent: "mousedown",
  });

  const actions = useMemo(
    () => [
      { id: "newThread", label: text.sidebar.newThread },
      { id: "addFolder", label: text.sidebar.addFolder },
      { id: "settings", label: text.sidebar.settings },
    ],
    [text],
  );

  const results = useMemo(
    () =>
      buildSidebarSearchResults({
        query,
        threads,
        folders,
        folderFiles,
        actions,
      }),
    [actions, folderFiles, folders, query, threads],
  );

  if (!open) return null;

  const showThreads = tab === "all" || tab === "threads";
  const showFiles = tab === "all" || tab === "files";
  const showActions = tab === "all" || tab === "actions";
  const hasResults =
    results.threads.length > 0 || results.files.length > 0 || results.actions.length > 0;

  const runAction = (id: string) => {
    if (id === "newThread") onCreateNewThread();
    if (id === "addFolder") onAddFolder();
    if (id === "settings") onOpenSettings();
    onClose();
  };

  return (
    <div className="sidebar-search-overlay">
      <div
        className="sidebar-search-palette"
        ref={cardRef}
        onClick={(event) => event.stopPropagation()}
      >
        <div className="sidebar-search-header">
          <div className="sidebar-search-input-row">
            <Search size={18} />
            <input
              ref={inputRef}
              value={query}
              onChange={(event) => setQuery(event.target.value)}
              placeholder={text.sidebar.searchPalettePlaceholder}
            />
          </div>
          <button
            type="button"
            className="sidebar-search-close"
            title={text.chat.close}
            aria-label={text.chat.close}
            onClick={onClose}
          >
            <X size={18} />
          </button>
        </div>

        <div className="sidebar-search-tabs">
          {(
            [
              ["all", text.sidebar.searchAll],
              ["threads", text.sidebar.searchThreads],
              ["files", text.sidebar.searchFiles],
              ["actions", text.sidebar.searchActions],
            ] as const
          ).map(([key, label]) => (
            <button
              key={key}
              className={`sidebar-search-tab${tab === key ? " active" : ""}`}
              onClick={() => setTab(key)}
            >
              {label}
            </button>
          ))}
        </div>

        <div className="sidebar-search-results">
          {showThreads && results.threads.length > 0 && (
            <section className="sidebar-search-section">
              <div className="sidebar-search-section-title">{text.sidebar.searchThreads}</div>
              {results.threads.map(({ thread, folder }) => (
                <button
                  key={thread.threadId}
                  className={`sidebar-search-result${thread.threadId === activeThreadId ? " active" : ""}`}
                  onClick={() => {
                    onSwitchThread(thread.threadId);
                    onClose();
                  }}
                >
                  <MessageSquare size={16} />
                  <span className="sidebar-search-result-main">
                    <span className="sidebar-search-result-title">
                      {thread.name || thread.preview || text.chat.newChat}
                    </span>
                    {folder && (
                      <span className="sidebar-search-result-subtitle">{folder.name}</span>
                    )}
                  </span>
                  <span className="sidebar-search-result-meta">
                    {formatTime(thread.updatedAt, language)}
                  </span>
                </button>
              ))}
            </section>
          )}

          {showFiles && results.files.length > 0 && (
            <section className="sidebar-search-section">
              <div className="sidebar-search-section-title">{text.sidebar.searchFiles}</div>
              {results.files.map(({ file, folder }) => (
                <button
                  key={file.filePath}
                  className="sidebar-search-result"
                  onClick={() => {
                    onAddFile(file);
                    onClose();
                  }}
                >
                  <FileText size={16} />
                  <span className="sidebar-search-result-main">
                    <span className="sidebar-search-result-title">{file.fileName}</span>
                    <span className="sidebar-search-result-subtitle">{folder.name}</span>
                  </span>
                </button>
              ))}
            </section>
          )}

          {showActions && results.actions.length > 0 && (
            <section className="sidebar-search-section">
              <div className="sidebar-search-section-title">{text.sidebar.searchActions}</div>
              {results.actions.map((action) => {
                const Icon =
                  action.id === "newThread"
                    ? PenLine
                    : action.id === "addFolder"
                      ? FolderOpen
                      : Settings;
                return (
                  <button
                    key={action.id}
                    className="sidebar-search-result"
                    onClick={() => runAction(action.id)}
                  >
                    <Icon size={16} />
                    <span className="sidebar-search-result-main">
                      <span className="sidebar-search-result-title">{action.label}</span>
                    </span>
                  </button>
                );
              })}
            </section>
          )}

          {!hasResults && (
            <div className="sidebar-search-empty">{text.sidebar.noSearchResults}</div>
          )}
        </div>
      </div>
    </div>
  );
};
