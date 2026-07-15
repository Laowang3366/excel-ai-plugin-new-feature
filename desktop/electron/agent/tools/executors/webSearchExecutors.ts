import type { ToolExecutor } from "../../shared/types";
import { clampNumber } from "../../shared/numberLimits";
import {
  assertRemoteDataProcessingAllowed,
  toRemoteDataPolicyResult,
  type RemoteDataTransferSummary,
} from "../../../shared/egressPolicy";
import { validateArgs } from "./validation";
import {
  getSearchDestination,
  searchWeb,
  type SearchFreshness,
  type WebSearchResponse as WebSearchProviderResponse,
} from "./webSearchProviders";

export type { WebSearchResultItem } from "./webSearchHtmlParsers";

export interface WebSearchResponse extends WebSearchProviderResponse {
  remoteProcessing?: RemoteDataTransferSummary;
}

export interface WebSearchExecutorDeps {
  isRemoteDataProcessingEnabled?: () => boolean;
}

export function addWebSearchExecutors(
  target: Map<string, ToolExecutor>,
  deps: WebSearchExecutorDeps = {},
): void {
  target.set("web.search", {
    name: "web.search",
    execute: async (args: Record<string, unknown>) => {
      const err = validateArgs(args, { query: "string" });
      if (err) return { success: false, error: err };
      const optionalErr = validateOptionalSearchArgs(args);
      if (optionalErr) return { success: false, error: optionalErr };

      const query = (args.query as string).trim();
      if (!query) return { success: false, error: "参数 query 不能为空" };

      const maxResults = clampNumber(args.maxResults, { fallback: 5, min: 1, max: 10 });
      const freshness = normalizeFreshness(args.freshness);

      try {
        assertRemoteDataProcessingAllowed({
          enabled: deps.isRemoteDataProcessingEnabled?.() === true,
          operation: "web-search",
          texts: [query],
        });
        const data = await searchWeb(query, maxResults, freshness);
        return {
          success: true,
          data: {
            ...data,
            remoteProcessing: {
              operation: "web-search",
              service: data.provider,
              destination: getSearchDestination(data.provider),
              dataSummary: `搜索查询，${query.length} 个字符`,
            },
          },
        };
      } catch (e: any) {
        const policyResult = toRemoteDataPolicyResult(e);
        if (policyResult) return policyResult;
        return { success: false, error: `联网搜索失败: ${e.message}` };
      }
    },
  });
}

function validateOptionalSearchArgs(args: Record<string, unknown>): string | null {
  if (args.maxResults !== undefined && typeof args.maxResults !== "number") {
    return "参数 maxResults 必须是 number";
  }
  if (args.freshness !== undefined && typeof args.freshness !== "string") {
    return "参数 freshness 必须是 string";
  }
  const freshness = normalizeFreshness(args.freshness);
  if (args.freshness !== undefined && freshness === "any" && args.freshness !== "any") {
    return "参数 freshness 必须是 day、week、month、year 或 any";
  }
  return null;
}

function normalizeFreshness(value: unknown): SearchFreshness {
  return value === "day" ||
    value === "week" ||
    value === "month" ||
    value === "year" ||
    value === "any"
    ? value
    : "any";
}
