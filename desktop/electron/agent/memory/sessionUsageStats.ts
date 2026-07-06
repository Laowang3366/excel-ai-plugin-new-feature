import * as fs from "fs";
import * as path from "path";
import { promisify } from "util";
import { type RolloutLine, type TokenUsage, mergeTokenUsage } from "../shared/types";

const readFile = promisify(fs.readFile);

export interface TurnStats {
  turnId: string;
  threadId: string;
  model: string;
  timestamp: number;
  messages: number;
  tokens: number;
  estimated: boolean;
}

export async function readUsageSummaryFromRolloutFiles(files: string[]): Promise<TurnStats[]> {
  const results: TurnStats[] = [];
  for (const filePath of files) {
    try {
      const content = await readFile(filePath, "utf-8");
      const threadId = extractThreadIdFromRolloutPath(filePath);
      results.push(...parseRolloutForStats(content, threadId));
    } catch {
      // 跳过损坏的文件
    }
  }
  return results;
}

function extractThreadIdFromRolloutPath(filePath: string): string {
  const filename = path.basename(filePath, ".jsonl");
  const threadIdMatch = filename.match(/thread-(.+)$/);
  return threadIdMatch ? `thread-${threadIdMatch[1]}` : filename;
}

function parseRolloutForStats(content: string, threadId: string): TurnStats[] {
  const lines = content.split("\n").filter((line) => line.trim());
  let model = "unknown";
  const turnUsageMap = new Map<string, TokenUsage>();
  const turnMessagesMap = new Map<string, number>();
  const turnTimestampMap = new Map<string, number>();

  for (const line of lines) {
    try {
      const parsed: RolloutLine = JSON.parse(line);
      const { item } = parsed;

      switch (item.type) {
        case "session_meta":
          if (item.meta.model) model = item.meta.model;
          break;

        case "turn_item": {
          const turnId = item.turnId;
          const itemType = item.item.type;
          if (itemType === "user_message" || itemType === "assistant_message") {
            turnMessagesMap.set(turnId, (turnMessagesMap.get(turnId) || 0) + 1);
          }
          const ts = (item.item as any).timestamp as number | undefined;
          if (ts && ts > 0) {
            turnTimestampMap.set(turnId, Math.max(turnTimestampMap.get(turnId) || 0, ts));
          }
          break;
        }

        case "turn_usage": {
          const turnId = item.turnId;
          const existing = turnUsageMap.get(turnId);
          turnUsageMap.set(turnId, existing ? mergeTokenUsage(existing, item.usage) : item.usage);
          break;
        }

        case "turn_context":
          if (!turnTimestampMap.has(item.turnId)) {
            turnTimestampMap.set(item.turnId, Date.now());
          }
          break;
      }
    } catch {
      continue;
    }
  }

  const results: TurnStats[] = [];
  for (const [turnId, usage] of turnUsageMap) {
    const tokens = usage.inputTokens + usage.outputTokens + (usage.reasoningOutputTokens ?? 0);
    if (tokens === 0) continue;

    results.push({
      turnId,
      threadId,
      model,
      timestamp: turnTimestampMap.get(turnId) || Date.now(),
      messages: turnMessagesMap.get(turnId) || 0,
      tokens,
      estimated: false,
    });
  }

  return results;
}
