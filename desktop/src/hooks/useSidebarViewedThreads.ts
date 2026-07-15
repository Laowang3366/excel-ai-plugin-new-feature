import { useCallback, useEffect, useRef, useState } from "react";
import type { ThreadMetadata } from "../electronApi";

export function useSidebarViewedThreads(threads: ThreadMetadata[], activeThreadId: string | null) {
  const [viewedThreadStatusAt, setViewedThreadStatusAt] = useState<Record<string, number>>({});
  const initializedViewedStatuses = useRef(false);

  useEffect(() => {
    if (initializedViewedStatuses.current || threads.length === 0) return;
    initializedViewedStatuses.current = true;
    const viewed: Record<string, number> = {};
    threads.forEach((thread) => {
      viewed[thread.threadId] = thread.updatedAt;
    });
    setViewedThreadStatusAt(viewed);
  }, [threads]);

  const markThreadViewed = useCallback(
    (threadId: string) => {
      const thread = threads.find((item) => item.threadId === threadId);
      if (!thread) return;
      setViewedThreadStatusAt((prev) => {
        if (prev[threadId] === thread.updatedAt) return prev;
        return { ...prev, [threadId]: thread.updatedAt };
      });
    },
    [threads],
  );

  useEffect(() => {
    if (activeThreadId) markThreadViewed(activeThreadId);
  }, [activeThreadId, markThreadViewed]);

  return {
    viewedThreadStatusAt,
    markThreadViewed,
  };
}
