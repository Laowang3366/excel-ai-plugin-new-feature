import { TOOL_DEFINITIONS, TOOL_DEFINITION_MAP } from "../tools";
import type { ToolName, ToolResult } from "../tools/types";
import { collectAgentStream, emptyUsage, sumUsage } from "./collectStream";
import { trimMessagesForRequest } from "./historyBudget";
import { isAbortError, throwIfAborted } from "./streamProvider";
import type {
  AgentLoopOptions,
  AgentMessage,
  AgentRunInput,
  AgentRunResult,
  AgentTokenUsage,
  AgentToolOutcome,
  LoopEvent,
  ParsedToolCall,
} from "./types";

function serializeToolOutcome(outcome: AgentToolOutcome): string {
  return JSON.stringify(outcome);
}

function isPlainObject(value: unknown): value is Record<string, unknown> {
  return value != null && typeof value === "object" && !Array.isArray(value);
}

function parseToolCall(raw: {
  id: string;
  name: string;
  argumentsJson: string;
}): ParsedToolCall {
  const argumentsJson = raw.argumentsJson || "{}";
  try {
    const parsed: unknown = JSON.parse(argumentsJson);
    if (!isPlainObject(parsed)) {
      return {
        id: raw.id,
        name: raw.name,
        argumentsJson,
        parseError: "arguments must be a plain object",
      };
    }
    return { id: raw.id, name: raw.name, argumentsJson, arguments: parsed };
  } catch (error) {
    return {
      id: raw.id,
      name: raw.name,
      argumentsJson,
      parseError: error instanceof Error ? error.message : String(error),
    };
  }
}

function isKnownToolName(name: string): name is ToolName {
  return Object.prototype.hasOwnProperty.call(TOOL_DEFINITION_MAP, name);
}

function emit(onEvent: AgentLoopOptions["onEvent"], event: LoopEvent): void {
  onEvent?.(event);
}

export class AgentLoop {
  private readonly options: Required<
    Pick<AgentLoopOptions, "provider" | "executor" | "systemPrompt">
  > &
    AgentLoopOptions;

  constructor(options: AgentLoopOptions) {
    if (
      options.maxRounds != null &&
      (!Number.isInteger(options.maxRounds) || options.maxRounds < 1)
    ) {
      throw new Error("maxRounds must be an integer >= 1");
    }
    this.options = options;
  }

  async run(input: AgentRunInput): Promise<AgentRunResult> {
    const tools = this.options.tools ?? TOOL_DEFINITIONS;
    const activeNames = new Set<string>(tools.map((t) => t.name));
    const maxRounds = this.options.maxRounds ?? 8;
    const signal = this.options.signal;
    const historyLen = input.history?.length ?? 0;
    const messages: AgentMessage[] = [
      ...(input.history ? input.history.slice() : []),
      { role: "user", content: input.userMessage },
    ];
    // Current-turn user message index; history + current-turn tool chains after it are protected.
    const protectFromIndex = historyLen;
    let rounds = 0;
    let assistantText = "";
    let usage: AgentTokenUsage = emptyUsage();
    let lastFinishReason: AgentRunResult["lastFinishReason"];
    let runEnded = false;

    const finish = (result: AgentRunResult): AgentRunResult => {
      if (!runEnded) {
        runEnded = true;
        emit(this.options.onEvent, { type: "run_end", result });
      }
      return result;
    };

    try {
      while (true) {
        throwIfAborted(signal);
        if (rounds >= maxRounds) {
          return finish({
            status: "max_rounds",
            assistantText,
            messages,
            rounds,
            usage,
            lastFinishReason: "max_rounds",
          });
        }

        const nextRound = rounds + 1;
        emit(this.options.onEvent, { type: "round_start", round: nextRound });

        let collected;
        try {
          const requestMessages = this.buildRequestMessages(
            messages,
            tools,
            protectFromIndex,
          );
          const stream = this.options.provider.streamChat({
            systemPrompt: this.options.systemPrompt,
            messages: requestMessages,
            tools: tools.slice(),
            signal,
          });
          collected = await collectAgentStream(stream, {
            signal,
            onTextDelta: (delta) => {
              emit(this.options.onEvent, {
                type: "text_delta",
                delta,
                round: nextRound,
              });
            },
          });
        } catch (error) {
          if (isAbortError(error) || signal?.aborted) {
            return finish({
              status: "aborted",
              assistantText,
              messages,
              rounds,
              usage,
              lastFinishReason: "aborted",
              error: { message: "aborted", kind: "aborted" },
            });
          }
          const message = error instanceof Error ? error.message : String(error);
          return finish({
            status: "failed",
            assistantText,
            messages,
            rounds,
            usage,
            lastFinishReason: "error",
            error: { message, kind: "provider" },
          });
        }

        if (collected.error) {
          const aborted = collected.error.kind === "aborted";
          return finish({
            status: aborted ? "aborted" : "failed",
            assistantText,
            messages,
            rounds,
            usage,
            lastFinishReason: aborted ? "aborted" : "error",
            error: collected.error,
          });
        }

        // Successful collect: count round and append assistant (raw toolCalls).
        rounds += 1;
        usage = sumUsage(usage, collected.usage);
        if (collected.assistantText) {
          assistantText += collected.assistantText;
        }
        lastFinishReason = collected.finishReason;
        messages.push({
          role: "assistant",
          content: collected.assistantText,
          toolCalls:
            collected.toolCalls.length > 0 ? collected.toolCalls.slice() : undefined,
        });
        emit(this.options.onEvent, {
          type: "round_end",
          round: rounds,
          finishReason: collected.finishReason,
          toolCallCount: collected.toolCalls.length,
        });

        if (collected.toolCalls.length === 0) {
          return finish({
            status: "completed",
            assistantText,
            messages,
            rounds,
            usage,
            lastFinishReason,
          });
        }

        for (const raw of collected.toolCalls) {
          throwIfAborted(signal);

          const parsed = parseToolCall(raw);
          emit(this.options.onEvent, {
            type: "tool_call_parsed",
            call: parsed,
            round: rounds,
          });

          let outcome: AgentToolOutcome;
          if (parsed.parseError) {
            outcome = {
              kind: "invalid_arguments",
              toolName: parsed.name,
              error: parsed.parseError,
              argumentsJson: parsed.argumentsJson,
            };
          } else if (!isKnownToolName(parsed.name) || !activeNames.has(parsed.name)) {
            outcome = {
              kind: "unknown_tool",
              toolName: parsed.name,
              error: `unknown or inactive tool: ${parsed.name}`,
            };
          } else {
            // Known + active + plain object: execute. Abort mid-await: settle then stop.
            const toolName: ToolName = parsed.name;
            let hostResult: ToolResult;
            try {
              hostResult = await this.options.executor.execute({
                name: toolName,
                arguments: parsed.arguments ?? {},
              });
            } catch (error) {
              if (isAbortError(error)) {
                // Should be rare from executor; treat as aborted without host outcome.
                return finish({
                  status: "aborted",
                  assistantText,
                  messages,
                  rounds,
                  usage,
                  lastFinishReason: "aborted",
                  error: { message: "aborted", kind: "aborted" },
                });
              }
              throw error;
            }
            outcome = {
              kind: "host",
              toolName: toolName,
              result: hostResult,
            };
          }

          emit(this.options.onEvent, {
            type: "tool_outcome",
            toolCallId: parsed.id,
            outcome,
            round: rounds,
          });
          messages.push({
            role: "tool",
            content: serializeToolOutcome(outcome),
            toolCallId: parsed.id,
            name: parsed.name,
          });

          // If abort happened during await, stop after writing this tool message.
          if (signal?.aborted) {
            return finish({
              status: "aborted",
              assistantText,
              messages,
              rounds,
              usage,
              lastFinishReason: "aborted",
              error: { message: "aborted", kind: "aborted" },
            });
          }
        }
      }
    } catch (error) {
      if (isAbortError(error) || signal?.aborted) {
        return finish({
          status: "aborted",
          assistantText,
          messages,
          rounds,
          usage,
          lastFinishReason: "aborted",
          error: { message: "aborted", kind: "aborted" },
        });
      }
      const message = error instanceof Error ? error.message : String(error);
      return finish({
        status: "failed",
        assistantText,
        messages,
        rounds,
        usage,
        lastFinishReason: "error",
        error: { message, kind: "provider" },
      });
    }
  }

  private buildRequestMessages(
    messages: AgentMessage[],
    tools: import("../tools/types").ToolDefinition[],
    protectFromIndex: number,
  ): AgentMessage[] {
    const contextWindowSize = this.options.contextWindowSize;
    if (
      contextWindowSize == null ||
      !Number.isFinite(contextWindowSize) ||
      contextWindowSize <= 0
    ) {
      return messages.slice();
    }
    return trimMessagesForRequest({
      messages,
      systemPrompt: this.options.systemPrompt,
      tools,
      contextWindowSize,
      protectFromIndex,
    });
  }
}
