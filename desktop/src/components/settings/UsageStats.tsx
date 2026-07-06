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
import {
  MODEL_COLORS,
  USAGE_TEXT,
  buildRowsFromStats,
  buildUsageStatsData,
  formatDayLabel,
  formatExactNumber,
  formatNumber,
  formatPercent,
  type RangeDays,
  type UsageRow,
} from "./usageStatsData";

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

  const data = useMemo(() => buildUsageStatsData(rows, rangeDays), [rows, rangeDays]);

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
