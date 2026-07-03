/**
 * 压缩通知 — 上下文压缩时的用户提示
 */

import React, { useState } from "react";
import type { CompactedItem, CompactProgressItem } from "../../electronApi";
import { Minimize2, X } from "../common/IconMap";
import { useSettingsStore } from "../../store/settingsStore";
import { getAppText } from "../../i18n";

interface CompactionNoticeProps {
  item?: CompactedItem | CompactProgressItem;
  message?: string;
}

export const CompactionNotice: React.FC<CompactionNoticeProps> = ({ item, message }) => {
  const { language } = useSettingsStore();
  const text = getAppText(language);
  const [dismissed, setDismissed] = useState(false);

  if (dismissed) return null;

  const displayMessage = message || (item?.type === "compact_progress"
    ? item.message
    : item
      ? text.assistant.compactionWithTokens(item.tokensBefore, item.tokensAfter)
      : text.assistant.compactionAuto);

  return (
    <div className="compaction-notice">
      <span className="compaction-icon"><Minimize2 size={14} /></span>
      <span className="compaction-text">{displayMessage}</span>
      <button
        className="compaction-dismiss"
        onClick={() => setDismissed(true)}
      >
        <X size={12} />
      </button>
    </div>
  );
};
