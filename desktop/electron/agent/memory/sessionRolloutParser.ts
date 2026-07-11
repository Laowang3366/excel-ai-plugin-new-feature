import {
  type RolloutLine,
  type Thread,
  type ThreadId,
  type ThreadMetadata,
  type Turn,
  type TurnId,
  mergeTokenUsage,
} from "../shared/types";

export function parseRolloutContent(content: string, threadId: ThreadId): Thread {
  const lines = content.split("\n").filter((line) => line.trim());
  let metadata: ThreadMetadata = {
    threadId,
    preview: "",
    modelProvider: "unknown",
    createdAt: Date.now(),
    updatedAt: Date.now(),
  };
  const turnsMap = new Map<TurnId, Turn>();
  const turnOrder: TurnId[] = [];
  let hasSessionMeta = false;

  for (const line of lines) {
    try {
      const parsed: RolloutLine = JSON.parse(line);
      const { item } = parsed;

      switch (item.type) {
        case "session_meta": {
          const metaTimestamp = new Date(item.meta.timestamp).getTime();
          metadata = {
            ...metadata,
            threadId: item.meta.id,
            modelProvider: item.meta.modelProvider,
            model: item.meta.model,
            createdAt: hasSessionMeta ? metadata.createdAt : metaTimestamp,
            updatedAt: hasSessionMeta ? metaTimestamp : metadata.updatedAt,
            ...(Object.prototype.hasOwnProperty.call(item.meta, "name")
              ? { name: item.meta.name ?? undefined }
              : {}),
            folderId: item.meta.folderId,
          };
          hasSessionMeta = true;
          break;
        }
        case "turn_context": {
          if (!turnsMap.has(item.turnId)) {
            turnsMap.set(item.turnId, {
              turnId: item.turnId,
              threadId: metadata.threadId,
              status: "in_progress",
              items: [],
              startedAt: Date.now(),
            });
            turnOrder.push(item.turnId);
          }
          break;
        }
        case "turn_item": {
          const turnId = item.turnId;
          if (!turnsMap.has(turnId)) {
            turnsMap.set(turnId, {
              turnId,
              threadId: metadata.threadId,
              status: "in_progress",
              items: [],
              startedAt: Date.now(),
            });
            turnOrder.push(turnId);
          }
          const turn = turnsMap.get(turnId)!;
          turn.items.push(item.item);

          if (item.item.type === "user_message" && !metadata.preview) {
            metadata.preview = item.item.content.slice(0, 100);
          }
          if (item.item.type === "assistant_message") {
            metadata.updatedAt = item.item.timestamp;
          }
          break;
        }
        case "turn_usage": {
          if (!turnsMap.has(item.turnId)) {
            turnsMap.set(item.turnId, {
              turnId: item.turnId,
              threadId: metadata.threadId,
              status: "in_progress",
              items: [],
              startedAt: Date.now(),
            });
            turnOrder.push(item.turnId);
          }
          const turn = turnsMap.get(item.turnId)!;
          turn.tokenUsage = item.usage;
          metadata.totalTokenUsage = metadata.totalTokenUsage
            ? mergeTokenUsage(metadata.totalTokenUsage, item.usage)
            : item.usage;
          break;
        }
        case "compacted": {
          metadata.compactedHistory = item.replacementHistory;
          for (const tId of turnOrder) {
            turnsMap.delete(tId);
          }
          turnOrder.length = 0;
          break;
        }
      }
    } catch {
      continue;
    }
  }

  const turns = Array.from(turnsMap.values());
  for (const turn of turns) {
    if (turn.status === "in_progress") {
      const hasFinalMessage = turn.items.some(
        (item) => item.type === "assistant_message" && item.phase === "final"
      );
      turn.status = hasFinalMessage ? "completed" : "interrupted";
    }
  }
  const lastTurn = turns.length > 0 ? turns[turns.length - 1] : undefined;
  if (lastTurn) {
    metadata.lastTurnStatus = lastTurn.status;
  }

  return {
    metadata,
    turns,
  };
}
