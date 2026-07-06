import { describe, expect, it } from "vitest";

import {
  buildRowsFromStats,
  buildUsageStatsData,
  formatDayLabel,
  formatNumber,
  formatPercent,
} from "./usageStatsData";

describe("usageStatsData", () => {
  it("builds sorted usage rows and ignores zero-token turns", () => {
    const rows = buildRowsFromStats([
      {
        turnId: "turn-zero",
        threadId: "thread-1",
        model: "gpt-4o",
        timestamp: Date.parse("2026-07-05T08:00:00Z"),
        messages: 1,
        tokens: 0,
        estimated: false,
      },
      {
        turnId: "turn-new",
        threadId: "thread-2",
        model: "gpt-4o",
        timestamp: Date.parse("2026-07-06T08:00:00Z"),
        messages: 2,
        tokens: 200,
        estimated: true,
      },
      {
        turnId: "turn-old",
        threadId: "thread-1",
        model: "qwen",
        timestamp: Date.parse("2026-07-04T08:00:00Z"),
        messages: 3,
        tokens: 100,
        estimated: false,
      },
    ]);

    expect(rows.map((row) => row.tokens)).toEqual([100, 200]);
    expect(rows.map((row) => row.dateKey)).toEqual(["2026-07-04", "2026-07-06"]);
  });

  it("aggregates usage rows for the selected range", () => {
    const now = Date.parse("2026-07-06T23:00:00Z");
    const rows = buildRowsFromStats([
      {
        turnId: "turn-1",
        threadId: "thread-1",
        model: "gpt-4o",
        timestamp: Date.parse("2026-07-05T08:00:00Z"),
        messages: 2,
        tokens: 300,
        estimated: false,
      },
      {
        turnId: "turn-2",
        threadId: "thread-2",
        model: "qwen",
        timestamp: Date.parse("2026-07-06T08:00:00Z"),
        messages: 1,
        tokens: 100,
        estimated: true,
      },
    ]);

    const data = buildUsageStatsData(rows, 7, now);

    expect(data.dateKeys).toHaveLength(7);
    expect(data.totalTokens).toBe(400);
    expect(data.totalMessages).toBe(3);
    expect(data.totalThreads).toBe(2);
    expect(data.activeDays).toBe(2);
    expect(data.streak).toBe(2);
    expect(data.estimated).toBe(true);
    expect(data.modelRows.map((row) => [row.model, row.tokens])).toEqual([
      ["gpt-4o", 300],
      ["qwen", 100],
    ]);
  });

  it("formats numbers, percentages, and day labels", () => {
    expect(formatNumber(1536)).toBe("1,536");
    expect(formatPercent(0.5)).toBe("0.5%");
    expect(formatPercent(42.2)).toBe("42%");
    expect(formatDayLabel("2026-07-06")).toBe("7月6日");
  });
});
