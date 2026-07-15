/**
 * 使用统计 — 基于本地会话历史聚合
 */

import React, { useEffect, useMemo, useState } from "react";
import { BarChart3, Bot, Clock, Flame, MessageSquare, RefreshCw } from "../common/IconMap";
import { useSettingsStore } from "../../store/settingsStore";
import { ipcApi } from "../../services/ipcApi";
import {
  MODEL_COLORS,
  USAGE_TEXT,
  buildRowsFromStats,
  buildUsageStatsData,
  formatNumber,
  formatPercent,
  type RangeDays,
  type UsageRow,
} from "./usageStatsData";
import { UsageStatsCharts } from "./UsageStatsCharts";

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
    data.modelRows.forEach((row, index) =>
      map.set(row.model, MODEL_COLORS[index % MODEL_COLORS.length]),
    );
    return map;
  }, [data.modelRows]);

  const dominantModel = data.modelRows[0];
  const donutGradient = data.modelRows.length
    ? data.modelRows
        .reduce<{ parts: string[]; offset: number }>(
          (acc, row) => {
            const start = acc.offset;
            const size = data.totalTokens ? (row.tokens / data.totalTokens) * 100 : 0;
            const end = start + size;
            acc.parts.push(`${modelColorMap.get(row.model)} ${start}% ${end}%`);
            acc.offset = end;
            return acc;
          },
          { parts: [], offset: 0 },
        )
        .parts.join(", ")
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
            <button className={rangeDays === 7 ? "active" : ""} onClick={() => setRangeDays(7)}>
              {text.last7}
            </button>
            <button className={rangeDays === 30 ? "active" : ""} onClick={() => setRangeDays(30)}>
              {text.last30}
            </button>
          </div>
          <button
            className="usage-refresh-btn"
            onClick={loadUsage}
            disabled={loading}
            title={text.refresh}
          >
            <RefreshCw size={15} className={loading ? "spin" : ""} />
          </button>
        </div>
      </div>

      {error && <div className="usage-error">{error}</div>}

      <div className="usage-metric-grid">
        <MetricCard
          icon={<Flame size={16} />}
          label={text.tokenUsage}
          value={formatNumber(data.totalTokens)}
          hint={data.estimated ? text.estimated : text.realStats}
        />
        <MetricCard
          icon={<MessageSquare size={16} />}
          label={text.sessions}
          value={formatNumber(data.totalThreads)}
        />
        <MetricCard
          icon={<BarChart3 size={16} />}
          label={text.messages}
          value={formatNumber(data.totalMessages)}
        />
        <MetricCard
          icon={<Clock size={16} />}
          label={text.activeDays}
          value={formatNumber(data.activeDays)}
        />
        <MetricCard
          icon={<Clock size={16} />}
          label={text.streak}
          value={formatNumber(data.streak)}
        />
        <MetricCard
          icon={<Bot size={16} />}
          label={text.topModel}
          value={dominantModel?.model || text.none}
          hint={
            dominantModel
              ? text.share(
                  formatPercent((dominantModel.tokens / Math.max(1, data.totalTokens)) * 100),
                )
              : undefined
          }
          compact
        />
      </div>

      <UsageStatsCharts
        data={data}
        rangeDays={rangeDays}
        modelColorMap={modelColorMap}
        donutGradient={donutGradient}
        text={{
          heatmap: text.heatmap,
          less: text.less,
          more: text.more,
          dailyTrend: text.dailyTrend,
          modelUsage: text.modelUsage,
          empty: text.empty,
        }}
      />
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
      <div className="usage-metric-label">
        {icon}
        <span>{label}</span>
      </div>
      <div className={compact ? "usage-metric-value compact" : "usage-metric-value"}>{value}</div>
      {hint && <div className="usage-metric-hint">{hint}</div>}
    </div>
  );
}
