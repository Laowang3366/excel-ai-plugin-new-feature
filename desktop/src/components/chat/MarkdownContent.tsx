import React, { useRef, useState } from "react";
import ReactMarkdown from "react-markdown";
import type { Components } from "react-markdown";
import remarkGfm from "remark-gfm";
import { Check, Copy } from "../common/IconMap";
import { useSettingsStore } from "../../store/settingsStore";
import { getAppText } from "../../i18n";
import { ipcApi } from "../../services/ipcApi";

interface MarkdownContentProps {
  content: string;
}

export function MarkdownContent({ content }: MarkdownContentProps) {
  const { language } = useSettingsStore();
  const text = getAppText(language);
  const displayContent = normalizeVisibleMarkdown(content);
  const markdownComponents: Components = {
    a({ href, children, ...props }) {
      const externalUrl = normalizeExternalHttpUrl(href);
      return (
        <a
          {...props}
          href={externalUrl ?? undefined}
          rel="noreferrer noopener"
          target="_blank"
          onClick={(event) => {
            event.preventDefault();
            if (externalUrl) void ipcApi.app.openExternal(externalUrl);
          }}
        >
          {children}
        </a>
      );
    },
    pre({ children }) {
      const codeText = extractText(children).replace(/\n$/, "");

      return (
        <div className="markdown-code-block">
          <CopyFeedbackButton
            className="markdown-code-copy-btn"
            copiedTitle={text.assistant.copied}
            textToCopy={codeText}
            title={text.assistant.copyCode}
          />
          <pre>{children}</pre>
        </div>
      );
    },
    table({ children }) {
      return (
        <div className="markdown-table-wrapper">
          <table>{children}</table>
        </div>
      );
    },
  };

  return (
    <ReactMarkdown remarkPlugins={[remarkGfm]} components={markdownComponents}>
      {displayContent}
    </ReactMarkdown>
  );
}

export function normalizeExternalHttpUrl(href: string | undefined): string | null {
  if (!href) return null;
  try {
    const url = new URL(href);
    return url.protocol === "http:" || url.protocol === "https:" ? url.toString() : null;
  } catch {
    return null;
  }
}

export function normalizeVisibleMarkdown(content: string): string {
  const lines = content.split("\n");
  let inFence = false;

  return lines.map((line) => {
    if (/^\s*```/.test(line)) {
      inFence = !inFence;
      return line;
    }
    if (inFence) return line;

    return line.replace(/^(\s*)#{1,6}(?:\s+|(?=[^\s#]))/, "$1");
  }).join("\n");
}

interface CopyFeedbackButtonProps {
  className: string;
  copiedTitle: string;
  textToCopy: string;
  title: string;
}

export function CopyFeedbackButton({ className, copiedTitle, textToCopy, title }: CopyFeedbackButtonProps) {
  const [copied, setCopied] = useState(false);
  const resetTimerRef = useRef<number | null>(null);

  async function handleCopy() {
    await navigator.clipboard?.writeText(textToCopy);
    setCopied(true);
    if (resetTimerRef.current !== null) {
      window.clearTimeout(resetTimerRef.current);
    }
    resetTimerRef.current = window.setTimeout(() => {
      setCopied(false);
      resetTimerRef.current = null;
    }, 1400);
  }

  return (
    <button
      className={`${className}${copied ? " copied" : ""}`}
      onClick={handleCopy}
      title={copied ? copiedTitle : title}
      type="button"
    >
      {copied ? (
        <>
          <Check size={13} />
          <span>{copiedTitle}</span>
        </>
      ) : (
        <Copy size={13} />
      )}
    </button>
  );
}

function extractText(node: React.ReactNode): string {
  if (node === null || node === undefined || typeof node === "boolean") return "";
  if (typeof node === "string" || typeof node === "number") return String(node);
  if (Array.isArray(node)) return node.map(extractText).join("");
  if (React.isValidElement<{ children?: React.ReactNode }>(node)) {
    return extractText(node.props.children);
  }
  return "";
}
