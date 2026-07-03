/**
 * 使用统计 — 基于本地会话历史聚合
 */

import React, { useEffect, useMemo, useState } from "react";
import {
  BarChart3,
  Bot,
  Clock,
  Flame,
  MessageSquare,
  RefreshCw,
} from "../common/IconMap";
import { useSettingsStore } from "../../store/settingsStore";
import { ipcApi } from "../../services/ipcApi";

type RangeDays = 7 | 30;

interface UsageRow {
  dateKey: string;
  timestamp: number;
  model: string;
  threadId: string;
  messages: number;
  tokens: number;
  estimated: boolean;
}

const MODEL_COLORS = ["#1683f7", "#1f8f45", "#8b6ff0", "#ef2d2d", "#e68600", "#12a6a6", "#64748b"];

const USAGE_TEXT = {
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

function dateKey(timestamp: number): string {
  return new Date(timestamp).toISOString().slice(0, 10);
}

function startOfLocalDay(timestamp: number): number {
  const date = new Date(timestamp);
  date.setHours(0, 0, 0, 0);
  return date.getTime();
}

function formatNumber(value: number): string {
  if (value >= 100_000_000) return `${(value / 100_000_000).toFixed(value >= 1_000_000_000 ? 1 : 2)}亿`;
  if (value >= 10_000) return `${(value / 10_000).toFixed(value >= 100_000 ? 1 : 2)}万`;
  return value.toLocaleString("zh-CN");
}

function formatExactNumber(value: number): string {
  return Math.round(value).toLocaleString("zh-CN");
}

function formatPercent(value: number): string {
  if (!Number.isFinite(value)) return "0%";
  if (value < 1 && value > 0) return `${value.toFixed(1)}%`;
  return `${Math.round(value)}%`;
}

function formatDayLabel(dateKeyValue: string): string {
  const date = new Date(`${dateKeyValue}T00:00:00`);
  return `${date.getMonth() + 1}月${date.getDate()}日`;
}

/** 从 stats:getSummary 返回的聚合数据直接构建 UsageRow（单次 IPC 调用） */
function buildRowsFromStats(stats: Array<{
  turnId: string;
  threadId: string;
  model: string;
  timestamp: number;
  messages: number;
  tokens: number;
  estimated: boolean;
}>): UsageRow[] {
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

function getRangeDateKeys(days: RangeDays): string[] {
  const today = startOfLocalDay(Date.now());
  return Array.from({ length: days }, (_, index) => dateKey(today - (days - 1 - index) * 86400000));
}

export const UsageStats: React.FC = () => {
  const { language } = useSettingsStore();
  const text = USAGE_TEXT[language];
  const [rangeDays, setRangeDays] = useState<RangeDays>(30);
  const [rows, setRows] = useState<UsageRow[]>([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState("");

  const loadUsage = async () => {
    setLoading(true);
    setError("");
    try {
      const stats = await ipcApi.stats.getSummary();
      setRows(buildRowsFromStats(stats));
    } catch (err) {
      setError(err instanceof Error ? err.message : "读取使用统计失败");
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    loadUsage();
  }, []);

  const data = useMemo(() => {
    const rangeStart = startOfLocalDay(Date.now()) - (rangeDays - 1) * 86400000;
    const filtered = rows.filter((row) => row.timestamp >= rangeStart);
    const dateKeys = getRangeDateKeys(rangeDays);
    const dateSet = new Set(dateKeys);
    const activeDateKeys = new Set(filtered.map((row) => row.dateKey));
    const threadIds = new Set(filtered.map((row) => row.threadId));
    const totalTokens = filtered.reduce((sum, row) => sum + row.tokens, 0);
    const totalMessages = filtered.reduce((sum, row) => sum + row.messages, 0);
    const estimated = filtered.some((row) => row.estimated);

    const byDate = new Map<string, UsageRow[]>();
    const byModel = new Map<string, { model: string; tokens: number; messages: number; turns: number }>();

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
    const today = startOfLocalDay(Date.now());
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
  }, [rows, rangeDays]);

  const modelColorMap = useMemo(() => {
    const map = new Map<string, string>();
    data.modelRows.forEach((row, index) => map.set(row.model, MODEL_COLORS[index % MODEL_COLORS.length]));
    return map;
  }, [data.modelRows]);

  const dominantModel = data.modelRows[0];
  const donutGradient = data.modelRows.length
    ? data.modelRows.reduce<{ parts: string[]; offset: number }>((acc, row) => {
        const start = acc.offset;
        const size = data.totalTokens ? (row.tokens / data.totalTokens) * 100 : 0;
        const end = start + size;
        acc.parts.push(`${modelColorMap.get(row.model)} ${start}% ${end}%`);
        acc.offset = end;
        return acc;
      }, { parts: [], offset: 0 }).parts.join(", ")
    : "#e5e7eb 0% 100%";

  return (
    <div className="settings-section-content usage-stats-page">
      <div className="usage-header">
        <div>
          <div className="usage-title-row">
            <h2>{text.title}</h2>
            <span className="usage-tab">{text.tab}</span>
          </div>
          <p className="section-desc">{text.desc}</p>
        </div>
        <div className="usage-actions">
          <div className="usage-range-toggle" aria-label={text.rangeLabel}>
            <button className={rangeDays === 7 ? "active" : ""} onClick={() => setRangeDays(7)}>{text.last7}</button>
            <button className={rangeDays === 30 ? "active" : ""} onClick={() => setRangeDays(30)}>{text.last30}</button>
          </div>
          <button className="usage-refresh-btn" onClick={loadUsage} disabled={loading} title={text.refresh}>
            <RefreshCw size={15} className={loading ? "spin" : ""} />
          </button>
        </div>
      </div>

      {error && <div className="usage-error">{error}</div>}

      <div className="usage-metric-grid">
        <MetricCard icon={<Flame size={16} />} label={text.tokenUsage} value={formatNumber(data.totalTokens)} hint={data.estimated ? text.estimated : text.realStats} />
        <MetricCard icon={<MessageSquare size={16} />} label={text.sessions} value={formatNumber(data.totalThreads)} />
        <MetricCard icon={<BarChart3 size={16} />} label={text.messages} value={formatNumber(data.totalMessages)} />
        <MetricCard icon={<Clock size={16} />} label={text.activeDays} value={formatNumber(data.activeDays)} />
        <MetricCard icon={<Clock size={16} />} label={text.streak} value={formatNumber(data.streak)} />
        <MetricCard
          icon={<Bot size={16} />}
          label={text.topModel}
          value={dominantModel?.model || text.none}
          hint={dominantModel ? text.share(formatPercent((dominantModel.tokens / Math.max(1, data.totalTokens)) * 100)) : undefined}
          compact
        />
      </div>

      <section className="usage-panel">
        <div className="usage-panel-title-row">
          <h3>{text.heatmap}</h3>
          <div className="usage-heat-legend">
            <span>{text.less}</span>
            {[0, 1, 2, 3, 4].map((level) => <i key={level} className={`level-${level}`} />)}
            <span>{text.more}</span>
          </div>
        </div>
        <div className="usage-heatmap" style={{ gridTemplateColumns: `repeat(${data.dateKeys.length}, 1fr)` }}>
          {data.dateKeys.map((key) => {
            const dayTokens = data.byDate.get(key)?.reduce((sum, row) => sum + row.tokens, 0) ?? 0;
            const level = dayTokens === 0 ? 0 : Math.min(4, Math.ceil((dayTokens / data.maxDailyTokens) * 4));
            return (
              <div
                key={key}
                className={`usage-heat-cell level-${level}`}
                title={`${formatDayLabel(key)}：${formatNumber(dayTokens)} tokens`}
              />
            );
          })}
        </div>
      </section>

      <section className="usage-panel">
        <h3>{text.dailyTrend}</h3>
        <div className="usage-chart-frame">
          <div className="usage-y-axis">
            {data.yAxisTicks.map((tick) => (
              <span key={tick}>{formatNumber(tick)}</span>
            ))}
          </div>
          <div className="usage-bar-chart">
            {data.yAxisTicks.map((tick, index) => (
              <div key={`${tick}-${index}`} className="usage-grid-line" style={{ top: `${index * 25}%` }} />
            ))}
            {data.dateKeys.map((key, index) => {
              const rows = data.byDate.get(key) ?? [];
              const dayTotal = rows.reduce((sum, row) => sum + row.tokens, 0);
              const byModel = Array.from(rows.reduce<Map<string, number>>((map, row) => {
                map.set(row.model, (map.get(row.model) ?? 0) + row.tokens);
                return map;
              }, new Map()).entries()).sort((a, b) => b[1] - a[1]);
              const showAxisLabel = index === 0 || index === data.dateKeys.length - 1 || index % (rangeDays === 30 ? 5 : 1) === 0;
              return (
                <div
                  key={key}
                  className="usage-bar-column"
                  aria-label={`${formatDayLabel(key)}：${formatNumber(dayTotal)} tokens`}
                >
                  <div className="usage-bar-stack">
                    {byModel.map(([model, tokens]) => (
                      <span
                        key={model}
                        style={{
                          height: `${Math.max(2, (tokens / data.maxDailyTokens) * 100)}%`,
                          background: modelColorMap.get(model),
                        }}
                      />
                    ))}
                  </div>
                  {dayTotal > 0 && (
                    <div className="usage-bar-tooltip">
                      <strong>{formatDayLabel(key)} - {formatNumber(dayTotal)} tokens</strong>
                      {byModel.map(([model, tokens]) => (
                        <div key={model}>
                          <span><i style={{ background: modelColorMap.get(model) }} />{model}</span>
                          <em>{formatExactNumber(tokens)}</em>
                        </div>
                      ))}
                    </div>
                  )}
                  <span className="usage-bar-label">{showAxisLabel ? formatDayLabel(key) : ""}</span>
                </div>
              );
            })}
          </div>
        </div>
        <div className="usage-model-legend">
          {data.modelRows.slice(0, 6).map((row) => (
            <span key={row.model}>
              <i style={{ background: modelColorMap.get(row.model) }} />
              <strong>{row.model}</strong>
              <em>{formatNumber(row.tokens)}</em>
            </span>
          ))}
        </div>
      </section>

      <section className="usage-panel usage-model-panel">
        <h3>{text.modelUsage}</h3>
        {data.modelRows.length === 0 ? (
          <div className="usage-empty">{text.empty}</div>
        ) : (
          <div className="usage-model-content">
            <div className="usage-donut" style={{ background: `conic-gradient(${donutGradient})` }}>
              <div>
                <strong>{formatNumber(data.totalTokens)}</strong>
                <span>tokens</span>
              </div>
            </div>
            <div className="usage-model-list">
              {data.modelRows.slice(0, 8).map((row) => (
                <div key={row.model} className="usage-model-row">
                  <i style={{ background: modelColorMap.get(row.model) }} />
                  <div>
                    <strong>{row.model}</strong>
                    <span>{formatNumber(row.tokens)} tokens</span>
                  </div>
                  <em>{formatPercent((row.tokens / Math.max(1, data.totalTokens)) * 100)}</em>
                </div>
              ))}
            </div>
          </div>
        )}
      </section>
    </div>
  );
};

function MetricCard({
  icon,
  label,
  value,
  hint,
  compact,
}: {
  icon: React.ReactNode;
  label: string;
  value: string;
  hint?: string;
  compact?: boolean;
}) {
  return (
    <div className="usage-metric-card">
      <div className="usage-metric-label">{icon}<span>{label}</span></div>
      <div className={compact ? "usage-metric-value compact" : "usage-metric-value"}>{value}</div>
      {hint && <div className="usage-metric-hint">{hint}</div>}
    </div>
  );
}
