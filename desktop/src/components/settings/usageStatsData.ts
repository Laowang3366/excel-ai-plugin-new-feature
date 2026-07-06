export type RangeDays = 7 | 30;

export interface UsageRow {
  dateKey: string;
  timestamp: number;
  model: string;
  threadId: string;
  messages: number;
  tokens: number;
  estimated: boolean;
}

export interface UsageSummaryStat {
  turnId: string;
  threadId: string;
  model: string;
  timestamp: number;
  messages: number;
  tokens: number;
  estimated: boolean;
}

export interface UsageModelRow {
  model: string;
  tokens: number;
  messages: number;
  turns: number;
}

export const MODEL_COLORS = ["#1683f7", "#1f8f45", "#8b6ff0", "#ef2d2d", "#e68600", "#12a6a6", "#64748b"];

export const USAGE_TEXT = {
  "zh-CN": {
    title: "使用统计",
    tab: "应用用量",
    desc: "基于本地会话历史统计 Token、消息和模型使用情况。",
    rangeLabel: "时间范围",
    last7: "最近 7 天",
    last30: "最近 30 天",
    refresh: "刷新统计",
    tokenUsage: "tokens 用量",
    sessions: "会话数量",
    messages: "消息数量",
    activeDays: "活跃天数",
    streak: "当前连续天数",
    topModel: "最常用模型",
    none: "暂无",
    estimated: "包含估算",
    realStats: "真实统计",
    share: (value: string) => `占比 ${value}`,
    heatmap: "活跃热力图",
    less: "较少",
    more: "较多",
    dailyTrend: "按天 Token 趋势",
    modelUsage: "模型用量",
    empty: "暂无可统计的会话数据",
  },
  "en-US": {
    title: "Usage",
    tab: "App usage",
    desc: "Token, message, and model usage based on local chat history.",
    rangeLabel: "Time range",
    last7: "Last 7 days",
    last30: "Last 30 days",
    refresh: "Refresh stats",
    tokenUsage: "Token usage",
    sessions: "Chats",
    messages: "Messages",
    activeDays: "Active days",
    streak: "Current streak",
    topModel: "Most used model",
    none: "None",
    estimated: "Includes estimates",
    realStats: "Actual stats",
    share: (value: string) => `${value} share`,
    heatmap: "Activity heatmap",
    less: "Less",
    more: "More",
    dailyTrend: "Daily token trend",
    modelUsage: "Model usage",
    empty: "No chat data to analyze yet",
  },
} as const;

export function dateKey(timestamp: number): string {
  return new Date(timestamp).toISOString().slice(0, 10);
}

function startOfLocalDay(timestamp: number): number {
  const date = new Date(timestamp);
  date.setHours(0, 0, 0, 0);
  return date.getTime();
}

export function formatNumber(value: number): string {
  if (value >= 100_000_000) return `${(value / 100_000_000).toFixed(value >= 1_000_000_000 ? 1 : 2)}亿`;
  if (value >= 10_000) return `${(value / 10_000).toFixed(value >= 100_000 ? 1 : 2)}万`;
  return value.toLocaleString("zh-CN");
}

export function formatExactNumber(value: number): string {
  return Math.round(value).toLocaleString("zh-CN");
}

export function formatPercent(value: number): string {
  if (!Number.isFinite(value)) return "0%";
  if (value < 1 && value > 0) return `${value.toFixed(1)}%`;
  return `${Math.round(value)}%`;
}

export function formatDayLabel(dateKeyValue: string): string {
  const date = new Date(`${dateKeyValue}T00:00:00`);
  return `${date.getMonth() + 1}月${date.getDate()}日`;
}

export function buildRowsFromStats(stats: UsageSummaryStat[]): UsageRow[] {
  return stats
    .filter((s) => s.tokens > 0)
    .map((s) => ({
      dateKey: dateKey(s.timestamp),
      timestamp: s.timestamp,
      model: s.model,
      threadId: s.threadId,
      messages: s.messages,
      tokens: s.tokens,
      estimated: s.estimated,
    }))
    .sort((a, b) => a.timestamp - b.timestamp);
}

export function getRangeDateKeys(days: RangeDays, now = Date.now()): string[] {
  const today = startOfLocalDay(now);
  return Array.from({ length: days }, (_, index) => dateKey(today - (days - 1 - index) * 86400000));
}

export function buildUsageStatsData(rows: UsageRow[], rangeDays: RangeDays, now = Date.now()) {
  const rangeStart = startOfLocalDay(now) - (rangeDays - 1) * 86400000;
  const filtered = rows.filter((row) => row.timestamp >= rangeStart);
  const dateKeys = getRangeDateKeys(rangeDays, now);
  const dateSet = new Set(dateKeys);
  const activeDateKeys = new Set(filtered.map((row) => row.dateKey));
  const threadIds = new Set(filtered.map((row) => row.threadId));
  const totalTokens = filtered.reduce((sum, row) => sum + row.tokens, 0);
  const totalMessages = filtered.reduce((sum, row) => sum + row.messages, 0);
  const estimated = filtered.some((row) => row.estimated);

  const byDate = new Map<string, UsageRow[]>();
  const byModel = new Map<string, UsageModelRow>();

  for (const key of dateKeys) byDate.set(key, []);
  for (const row of filtered) {
    if (!dateSet.has(row.dateKey)) continue;
    byDate.get(row.dateKey)?.push(row);
    const current = byModel.get(row.model) ?? { model: row.model, tokens: 0, messages: 0, turns: 0 };
    current.tokens += row.tokens;
    current.messages += row.messages;
    current.turns += 1;
    byModel.set(row.model, current);
  }

  const modelRows = Array.from(byModel.values()).sort((a, b) => b.tokens - a.tokens);
  const maxDailyTokens = Math.max(1, ...Array.from(byDate.values()).map((items) => items.reduce((sum, row) => sum + row.tokens, 0)));
  const yAxisTicks = Array.from({ length: 5 }, (_, index) => Math.round((maxDailyTokens * (4 - index)) / 4));
  const today = startOfLocalDay(now);
  let streak = 0;
  for (let day = today; day >= today - 365 * 86400000; day -= 86400000) {
    if (!activeDateKeys.has(dateKey(day))) break;
    streak += 1;
  }

  return {
    rows: filtered,
    dateKeys,
    byDate,
    modelRows,
    totalTokens,
    totalMessages,
    totalThreads: threadIds.size,
    activeDays: activeDateKeys.size,
    streak,
    maxDailyTokens,
    yAxisTicks,
    estimated,
  };
}
