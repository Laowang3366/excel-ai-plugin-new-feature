import React from "react";
import {
  formatDayLabel,
  formatExactNumber,
  formatNumber,
  formatPercent,
  type RangeDays,
  type UsageRow,
  type UsageStatsData,
} from "./usageStatsData";

interface UsageStatsChartsProps {
  data: UsageStatsData;
  rangeDays: RangeDays;
  modelColorMap: Map<string, string>;
  donutGradient: string;
  text: {
    heatmap: string;
    less: string;
    more: string;
    dailyTrend: string;
    modelUsage: string;
    empty: string;
  };
}

export function UsageStatsCharts({
  data,
  rangeDays,
  modelColorMap,
  donutGradient,
  text,
}: UsageStatsChartsProps) {
  return (
    <>
      <section className="usage-panel">
        <div className="usage-panel-title-row">
          <h3>{text.heatmap}</h3>
          <div className="usage-heat-legend">
            <span>{text.less}</span>
            {[0, 1, 2, 3, 4].map((level) => (
              <i key={level} className={`level-${level}`} />
            ))}
            <span>{text.more}</span>
          </div>
        </div>
        <div
          className="usage-heatmap"
          style={{ gridTemplateColumns: `repeat(${data.dateKeys.length}, 1fr)` }}
        >
          {data.dateKeys.map((key) => {
            const dayTokens =
              data.byDate.get(key)?.reduce((sum: number, row: UsageRow) => sum + row.tokens, 0) ??
              0;
            const level =
              dayTokens === 0 ? 0 : Math.min(4, Math.ceil((dayTokens / data.maxDailyTokens) * 4));
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
              <div
                key={`${tick}-${index}`}
                className="usage-grid-line"
                style={{ top: `${index * 25}%` }}
              />
            ))}
            {data.dateKeys.map((key, index) => {
              const rows: UsageRow[] = data.byDate.get(key) ?? [];
              const dayTotal = rows.reduce((sum: number, row: UsageRow) => sum + row.tokens, 0);
              const byModel: Array<[string, number]> = Array.from(
                rows
                  .reduce((map: Map<string, number>, row: UsageRow) => {
                    map.set(row.model, (map.get(row.model) ?? 0) + row.tokens);
                    return map;
                  }, new Map<string, number>())
                  .entries(),
              ).sort((a, b) => b[1] - a[1]);
              const showAxisLabel =
                index === 0 ||
                index === data.dateKeys.length - 1 ||
                index % (rangeDays === 30 ? 5 : 1) === 0;
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
                      <strong>
                        {formatDayLabel(key)} - {formatNumber(dayTotal)} tokens
                      </strong>
                      {byModel.map(([model, tokens]) => (
                        <div key={model}>
                          <span>
                            <i style={{ background: modelColorMap.get(model) }} />
                            {model}
                          </span>
                          <em>{formatExactNumber(tokens)}</em>
                        </div>
                      ))}
                    </div>
                  )}
                  <span className="usage-bar-label">
                    {showAxisLabel ? formatDayLabel(key) : ""}
                  </span>
                </div>
              );
            })}
          </div>
        </div>
        <div className="usage-model-legend">
          {data.modelRows.slice(0, 6).map((row: { model: string; tokens: number }) => (
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
              {data.modelRows.slice(0, 8).map((row: { model: string; tokens: number }) => (
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
    </>
  );
}
