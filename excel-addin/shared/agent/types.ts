import type { ToolDefinition, ToolResult } from "../tools/types";

export type AgentFinishReason =
  | "stop"
  | "tool_calls"
  | "length"
  | "content_filter"
  | "error"
  | "unknown";

export interface AgentTokenUsage {
  inputTokens: number;
  outputTokens: number;
  cachedInputTokens?: number;
  reasoningOutputTokens?: number;
}

export type AgentStreamEvent =
  | { type: "text_delta"; delta: string }
  | { type: "tool_call_begin"; toolCallId: string; toolName: string }
  | { type: "tool_call_delta"; toolCallId: string; argumentsDelta: string }
  | {
      type: "tool_call_end";
      toolCallId: string;
      toolName?: string;
      argumentsJson?: string;
    }
  | { type: "usage"; usage: AgentTokenUsage }
  | { type: "finish"; reason: AgentFinishReason; rawReason?: string }
  | {
      type: "error";
      message: string;
      kind?:
        | "missing_key"
        | "http"
        | "network"
        | "cors"
        | "parse"
        | "aborted"
        | "unsupported"
        | "provider";
      status?: number;
      url?: string;
    };

export type AgentRole = "system" | "user" | "assistant" | "tool";

export interface AgentToolCall {
  id: string;
  name: string;
  argumentsJson: string;
}

export interface AgentMessage {
  role: AgentRole;
  content: string;
  toolCallId?: string;
  name?: string;
  toolCalls?: AgentToolCall[];
}

export interface ParsedToolCall {
  id: string;
  name: string;
  argumentsJson: string;
  arguments?: Record<string, unknown>;
  parseError?: string;
}

export type AgentToolOutcome =
  | { kind: "host"; toolName: string; result: ToolResult }
  | { kind: "unknown_tool"; toolName: string; error: string }
  | {
      kind: "invalid_arguments";
      toolName: string;
      error: string;
      argumentsJson: string;
    };

export type AgentErrorKind =
  | "missing_key"
  | "http"
  | "network"
  | "cors"
  | "parse"
  | "aborted"
  | "unsupported"
  | "provider";

export interface AgentStreamError {
  message: string;
  kind?: AgentErrorKind;
  status?: number;
  url?: string;
}

export interface AgentRoundStreamResult {
  assistantText: string;
  toolCalls: AgentToolCall[];
  finishReason: AgentFinishReason;
  usage?: AgentTokenUsage;
  error?: AgentStreamError;
}

export type AgentRunStatus = "completed" | "failed" | "aborted" | "max_rounds";

export interface AgentRunResult {
  status: AgentRunStatus;
  assistantText: string;
  messages: AgentMessage[];
  rounds: number;
  usage: AgentTokenUsage;
  lastFinishReason?: AgentFinishReason | "aborted" | "max_rounds";
  error?: AgentStreamError & { message: string };
}

export type LoopEvent =
  | { type: "round_start"; round: number }
  | { type: "text_delta"; delta: string; round: number }
  | { type: "tool_call_parsed"; call: ParsedToolCall; round: number }
  | { type: "tool_outcome"; toolCallId: string; outcome: AgentToolOutcome; round: number }
  | {
      type: "round_end";
      round: number;
      finishReason: AgentFinishReason;
      toolCallCount: number;
    }
  | { type: "run_end"; result: AgentRunResult };

export interface StreamChatRequest {
  systemPrompt: string;
  messages: AgentMessage[];
  tools: ToolDefinition[];
  signal?: AbortSignal;
}

export interface AgentStreamProvider {
  streamChat(request: StreamChatRequest): AsyncIterable<AgentStreamEvent>;
}

export interface AgentLoopOptions {
  provider: AgentStreamProvider;
  executor: {
    execute(call: {
      name: import("../tools/types").ToolName;
      arguments: Record<string, unknown>;
    }): Promise<ToolResult>;
  };
  systemPrompt: string;
  tools?: ToolDefinition[];
  maxRounds?: number;
  signal?: AbortSignal;
  onEvent?: (event: LoopEvent) => void;
}

export interface AgentRunInput {
  userMessage: string;
  history?: AgentMessage[];
}
